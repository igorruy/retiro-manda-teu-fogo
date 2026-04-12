#!/usr/bin/env python3
"""
Gerador de Plaquetas - Retiro de Crisma 2026
Lê lista de pessoas de um Excel e gera PDF A4 com plaquetas para portas dos quartos.
Cada metade de A4 = uma plaqueta de um quarto.

Uso:
    python gerar_plaquetas.py                          # usa lista_quartos.xlsx padrão
    python gerar_plaquetas.py minha_lista.xlsx         # usa outro arquivo
    python gerar_plaquetas.py lista.xlsx saida.pdf     # define nome da saída

Colunas obrigatórias no Excel:
    Nome   | Tipo (Crismando/Servo) | Equipe | Quarto
"""

import sys
import re
import os
from pathlib import Path
from collections import defaultdict

import pandas as pd
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor


# ─── Configurações ───────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).parent

# Templates das plaquetas (imagens PNG)
TEMPLATE_CRISMANDOS = SCRIPT_DIR / "template_crismandos.png"
TEMPLATE_SERVOS     = SCRIPT_DIR / "template_servos.png"

# Área de texto dentro da plaqueta (em proporção 0–1 da imagem)
# Ajuste se precisar reposicionar os nomes dentro do retângulo branco
TEXT_AREA = {
    "crismando": {
        "x_start": 0.350,  # começo horizontal do box de nomes
        "x_end":   0.985,
        "y_start": 0.195,  # topo do box
        "y_end":   0.820,
    },
    "servo": {
        "x_start": 0.350,
        "x_end":   0.980,
        "y_start": 0.195,
        "y_end":   0.820,
    },
}

# Cor e tamanho do texto
TEXT_COLOR   = HexColor("#4A1010")   # marrom escuro (combina com o layout)
FONT_NAME    = "Helvetica-Bold"
MAX_FONT_PT  = 18
MIN_FONT_PT  = 7

# ─── Funções auxiliares ───────────────────────────────────────────────────────

def load_data(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path)
    df.columns = [c.strip() for c in df.columns]

    col_map = {}
    for col in df.columns:
        low = col.lower()
        if "nome" in low:                col_map[col] = "Nome"
        elif "tipo" in low:              col_map[col] = "Tipo"
        elif "equipe" in low:            col_map[col] = "Equipe"
        elif "quarto" in low or "room" in low: col_map[col] = "Quarto"
    df = df.rename(columns=col_map)

    required = {"Nome", "Tipo", "Quarto"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Colunas não encontradas no Excel: {missing}\n"
                         f"Colunas disponíveis: {list(df.columns)}")

    df = df.dropna(subset=["Nome", "Quarto"])
    df["Nome"]   = df["Nome"].astype(str).str.strip()
    df["Tipo"]   = df["Tipo"].astype(str).str.strip().str.lower()
    # Converte para string preservando sufixos como "54S";
    # remove ".0" que o pandas adiciona ao ler números do Excel
    df["Quarto"] = df["Quarto"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    if "Equipe" in df.columns:
        df["Equipe"] = df["Equipe"].astype(str).str.strip()
    else:
        df["Equipe"] = ""
    return df


def classify(tipo_str: str) -> str:
    """Retorna 'crismando' ou 'servo'."""
    t = tipo_str.lower()
    if "crismand" in t:
        return "crismando"
    return "servo"


def group_by_room(df: pd.DataFrame):
    """Agrupa pessoas por (quarto, tipo) -> {'crismando': {...}, 'servo': {...}}"""
    rooms = defaultdict(lambda: {"crismando": [], "servo": []})
    for _, row in df.iterrows():
        tipo = classify(row["Tipo"])
        quarto = row["Quarto"]
        entry = {"nome": row["Nome"], "equipe": row.get("Equipe", "")}
        rooms[quarto][tipo].append(entry)
    return rooms


def best_font_size(c: canvas.Canvas, names: list[str], area_w: float, area_h: float) -> int:
    """Encontra o maior font size que cabe todos os nomes na área."""
    for pt in range(MAX_FONT_PT, MIN_FONT_PT - 1, -1):
        c.setFont(FONT_NAME, pt)
        line_h = pt * 1.4
        total_h = line_h * len(names)
        max_w   = max(c.stringWidth(n, FONT_NAME, pt) for n in names)
        if total_h <= area_h and max_w <= area_w:
            return pt
    return MIN_FONT_PT


def draw_placard(c: canvas.Canvas, template_path: Path, names: list[str],
                 x: float, y: float, w: float, h: float, tipo: str):
    """Desenha uma plaqueta na posição (x,y) com largura w e altura h."""
    # Fundo: imagem template
    c.drawImage(str(template_path), x, y, width=w, height=h,
                preserveAspectRatio=False, mask="auto")

    # Área de texto
    ta = TEXT_AREA[tipo]
    tx = x + ta["x_start"] * w
    ty_top = y + (1 - ta["y_start"]) * h
    ty_bot = y + (1 - ta["y_end"])   * h
    area_w = (ta["x_end"] - ta["x_start"]) * w
    area_h = ty_top - ty_bot

    if not names:
        return

    pt = best_font_size(c, names, area_w, area_h)
    c.setFont(FONT_NAME, pt)
    c.setFillColor(TEXT_COLOR)

    line_h = pt * 1.4
    total_h = line_h * len(names)
    # Centraliza verticalmente
    start_y = ty_top - (area_h - total_h) / 2 - pt

    for name in names:
        c.drawCentredString(tx + area_w / 2, start_y, name)
        start_y -= line_h


def generate_pdf(rooms: dict, output_path: str):
    """Gera o PDF A4 com duas plaquetas por página."""
    A4_W, A4_H = A4
    MARGIN = 5 * mm
    plaq_w = A4_W - 2 * MARGIN
    plaq_h = (A4_H - 3 * MARGIN) / 2   # duas por página

    c = canvas.Canvas(output_path, pagesize=A4)

    # Coleta todas as plaquetas a gerar: (quarto, tipo, nomes_list)
    placards = []
    def room_sort_key(q):
        # Separa parte numérica do sufixo (ex: "54S" -> (54, "S"), "10" -> (10, ""))
        m = re.match(r'^(\d+)(.*)', q.strip())
        if m:
            return (int(m.group(1)), m.group(2).upper())
        return (9999, q.upper())
    for quarto in sorted(rooms.keys(), key=room_sort_key):
        for tipo in ("crismando", "servo"):
            pessoas = rooms[quarto][tipo]
            if not pessoas:
                continue
            names = [p["nome"] for p in pessoas]
            equipe = pessoas[0]["equipe"] if pessoas[0]["equipe"] else ""
            template = TEMPLATE_CRISMANDOS if tipo == "crismando" else TEMPLATE_SERVOS
            placards.append((quarto, tipo, names, template))

    if not placards:
        print("⚠️  Nenhuma pessoa encontrada. Verifique a planilha.")
        return

    # Renderiza 2 plaquetas por página
    for i in range(0, len(placards), 2):
        batch = placards[i:i+2]

        # Plaqueta superior
        quarto, tipo, names, template = batch[0]
        y_top = MARGIN + plaq_h + MARGIN
        draw_placard(c, template, names, MARGIN, y_top, plaq_w, plaq_h, tipo)
        # Rótulo de quarto (pequeno, fora da plaqueta)
        c.setFont("Helvetica", 7)
        c.setFillColor(HexColor("#888888"))
        c.drawString(MARGIN, y_top + plaq_h + 1 * mm, f"Quarto {quarto} – {tipo.capitalize()}")

        # Plaqueta inferior (se existir)
        if len(batch) > 1:
            quarto2, tipo2, names2, template2 = batch[1]
            draw_placard(c, template2, names2, MARGIN, MARGIN, plaq_w, plaq_h, tipo2)
            c.setFont("Helvetica", 7)
            c.setFillColor(HexColor("#888888"))
            c.drawString(MARGIN, MARGIN + plaq_h + 1 * mm, f"Quarto {quarto2} – {tipo2.capitalize()}")

        c.showPage()

    c.save()
    print(f"✅  PDF gerado: {output_path}  ({len(placards)} plaquetas, {-(-len(placards)//2)} páginas)")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    excel_path  = sys.argv[1] if len(sys.argv) > 1 else "lista_quartos.xlsx"
    output_path = sys.argv[2] if len(sys.argv) > 2 else "plaquetas_quartos.pdf"

    if not os.path.exists(excel_path):
        print(f"❌  Arquivo não encontrado: {excel_path}")
        print("    Crie um arquivo Excel com as colunas: Nome | Tipo | Equipe | Quarto")
        sys.exit(1)

    for tmpl in (TEMPLATE_CRISMANDOS, TEMPLATE_SERVOS):
        if not tmpl.exists():
            print(f"❌  Template não encontrado: {tmpl}")
            sys.exit(1)

    print(f"📋  Lendo: {excel_path}")
    df = load_data(excel_path)
    print(f"    {len(df)} pessoas encontradas")

    rooms = group_by_room(df)
    print(f"    {len(rooms)} quartos com pessoas")

    generate_pdf(rooms, output_path)


if __name__ == "__main__":
    main()
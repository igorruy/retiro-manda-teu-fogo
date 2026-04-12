# Retiro de Crisma 2026 — Manda teu Fogo

Este repositório reúne:

- A página do cronograma do retiro (GitHub Pages): [index.html](index.html)
- Arquivos de referência (PDF/JPEG/XLSX): [Arquivos de referência/](Arquivos%20de%20refer%C3%AAncia/)
- Ferramentas de impressão:
  - Plaquetas de quartos: [gerador_plaquetas/](gerador_plaquetas/)
  - Crachás: [gerador_crachas/](gerador_crachas/)

## Site (index.html)

Página HTML estática para acompanhamento do retiro, com:

- Cronograma (cronograma geral + tarefas por equipe)
- Atividades dos crismandos (programação geral)
- Equipes (coordenação e responsabilidades)
- Arquivos úteis (PDF/JPEG para consulta e download)

### Rodar localmente

Na raiz do projeto:

```bash
python -m http.server 8000
```

Abra:

- http://localhost:8000/

## Instalação (dependências Python)

Recomendado criar um ambiente virtual na raiz:

```bash
python -m venv .venv
```

Ativar no Windows (PowerShell):

```bash
.\.venv\Scripts\Activate.ps1
```

Instalar dependências (cobre os dois geradores):

```bash
python -m pip install --upgrade pip
python -m pip install pandas openpyxl pillow reportlab
```

## Gerador de plaquetas (quartos)

Pasta: [gerador_plaquetas/](gerador_plaquetas/)

Gera um PDF A4 com plaquetas para portas dos quartos a partir de um Excel.
Cada página A4 contém 2 plaquetas (metade superior e metade inferior).

### Arquivos

- Script: `gerador_plaquetas/gerar_plaquetas.py`
- Entrada: `gerador_plaquetas/lista_quartos.xlsx`
- Templates:
  - `gerador_plaquetas/template_crismandos.png`
  - `gerador_plaquetas/template_servos.png`

### Formato do Excel

Colunas obrigatórias (nomes podem variar, o script tenta reconhecer por similaridade):

- `Nome`
- `Tipo` (Crismando/Servo)
- `Quarto`

Coluna opcional:

- `Equipe`

### Como usar

No diretório `gerador_plaquetas`:

```bash
python gerar_plaquetas.py
```

Usando outra planilha:

```bash
python gerar_plaquetas.py minha_lista.xlsx
```

Definindo o nome do PDF de saída:

```bash
python gerar_plaquetas.py minha_lista.xlsx saida.pdf
```

### Saída

Por padrão: `plaquetas_quartos.pdf`.

### Ajustes de layout

No `gerador_plaquetas/gerar_plaquetas.py`, ajuste:

- `TEXT_AREA` (área proporcional do texto no template)
- `MAX_FONT_PT` / `MIN_FONT_PT` (limites do tamanho da fonte)
- `TEXT_COLOR` (cor do texto)

## Gerador de crachás

Pasta: [gerador_crachas/](gerador_crachas/)

Gera um PDF A4 com crachás prontos para corte a partir de um Excel.
Se a planilha não existir, o script cria um modelo automaticamente.

### Arquivos

- Script: `gerador_crachas/design/gerador_crachas.py`
- Entrada: `gerador_crachas/design/lista_crachas.xlsx`
- Templates PNG: `gerador_crachas/design/*.png`

### Formato do Excel

Colunas obrigatórias:

- `Nome`
- `Equipe`

### Como usar

No diretório `gerador_crachas/design`:

```bash
python gerador_crachas.py
```

Usando outra planilha:

```bash
python gerador_crachas.py minha_lista.xlsx
```

Definindo o nome do PDF de saída:

```bash
python gerador_crachas.py minha_lista.xlsx saida.pdf
```

### Saída

Por padrão: `crachas.pdf`.

### Observações

- O script tenta mapear o nome da equipe para um PNG em `design/` (ex.: Guardiões, Missão, Providência, Movimento etc.).
- Se algum template não for encontrado, o crachá é gerado com fundo branco e a equipe é listada no terminal.


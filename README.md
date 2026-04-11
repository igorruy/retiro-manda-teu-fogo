# Gerador de Plaquetas (Retiro de Crisma 2026)

Gera um PDF A4 com plaquetas para portas de quartos a partir de uma planilha Excel.

Cada página A4 contém 2 plaquetas (metade superior e metade inferior).

## Pré-requisitos

- Python 3
- Dependências:
  - pandas
  - openpyxl
  - pillow
  - reportlab

## Arquivos do diretório

- `gerar_plaquetas.py`: script principal
- `lista_quartos.xlsx`: exemplo/base de entrada
- `template_crismandos.png`: template de plaqueta para crismandos
- `template_servos.png`: template de plaqueta para servos

## Formato da planilha (obrigatório)

Colunas obrigatórias (nomes podem variar, o script tenta reconhecer por similaridade):

- `Nome`
- `Tipo` (Crismando/Servo)
- `Quarto`

Coluna opcional:

- `Equipe`

## Como usar

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

## Saída

O PDF gerado (por padrão `plaquetas_quartos.pdf`) contém uma plaqueta por tipo dentro de cada quarto:

- `Crismando` (usa `template_crismandos.png`)
- `Servo` (usa `template_servos.png`)

Se um quarto não tiver pessoas de um tipo, a plaqueta daquele tipo não é gerada.

## Ajustes de layout

Se precisar reposicionar os nomes dentro do retângulo branco, ajuste no `gerar_plaquetas.py`:

- `TEXT_AREA` (área proporcional do texto no template)
- `MAX_FONT_PT` / `MIN_FONT_PT` (limites de tamanho da fonte)
- `TEXT_COLOR` (cor do texto)

## Problemas comuns

- `ModuleNotFoundError: No module named ...`:
  - Instale as dependências, por exemplo:

```bash
python -m pip install pandas openpyxl pillow reportlab
```

- `❌ Template não encontrado`:
  - Confirme se `template_crismandos.png` e `template_servos.png` estão na mesma pasta do script.


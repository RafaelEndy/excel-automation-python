# Excel Report Automation

This is a Python script built to automate the cleaning and organization of sales spreadsheets.

It takes a "raw" spreadsheet — with inconsistent date formats, differently-spelled categories, missing values, and duplicate rows — and turns it into a clean, formatted Excel report with a chart.

## What the script does

1. Reads the raw data from `data/vendas_brutas.csv`
2. Cleans the data:
   - removes extra whitespace from text
   - standardizes categories (`PERIFERICOS`, `perifericos` → `Perifericos`)
   - unifies dates in different formats (`01/03/2026`, `2026-03-02`, `04-03-2026`)
   - fills missing quantity/price with the product's median (or overall median, as a fallback)
   - removes duplicate rows
3. Calculates total revenue by category
4. Generates:
   - `output/relatorio_vendas.xlsx` — spreadsheet with a Summary tab (with a native Excel chart) and a Cleaned Data tab, with formatted headers
   - `output/grafico_receita_categoria.png` — standalone bar chart image

## How to run

```bash
pip install -r requirements.txt
python src/gerar_relatorio.py
```

## Structure

```
excel-automation-python/
├── data/
│   └── vendas_brutas.csv       # input data (messy example)
├── src/
│   └── gerar_relatorio.py      # main script
├── output/                     # generated automatically when run
└── requirements.txt
```

## Tech stack

- **pandas** — reading, cleaning, and aggregating the data
- **openpyxl** — generating the formatted Excel file (colors, fonts, native chart)
- **matplotlib** — generating the chart image

## Possible improvements

- Support `.xlsx` input in addition to `.csv`
- Add a period filter (month/year) via command line
- Also export to PDF

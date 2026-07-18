# Excel Report Automation

Este é um script Python desenvolvido para automatizar a limpeza e organização de planilhas de vendas.

O script recebe uma planilha "crua" — com formatos de data inconsistentes, categorias com grafia diferente, valores ausentes e linhas duplicadas — e transforma em um relatório Excel limpo, formatado e com gráfico.

## O que o script faz

1. Lê os dados brutos de `data/vendas_brutas.csv`
2. Limpa os dados:
   - remove espaços extras de texto
   - padroniza categorias (`PERIFERICOS`, `perifericos` → `Perifericos`)
   - unifica datas em formatos diferentes (`01/03/2026`, `2026-03-02`, `04-03-2026`)
   - preenche quantidade/valor ausentes com a mediana do produto (ou geral, como fallback)
   - remove linhas duplicadas
3. Calcula a receita total por categoria
4. Gera:
   - `output/relatorio_vendas.xlsx` — planilha com aba de Resumo (com gráfico nativo do Excel) e aba de Dados Tratados, com cabeçalhos formatados
   - `output/grafico_receita_categoria.png` — gráfico de barras separado, em imagem

## Como rodar

```bash
pip install -r requirements.txt
python src/gerar_relatorio.py
```

## Estrutura

```
excel-automation-python/
├── data/
│   └── vendas_brutas.csv       # dados de entrada (exemplo bagunçado)
├── src/
│   └── gerar_relatorio.py      # script principal
├── output/                     # gerado automaticamente ao rodar
└── requirements.txt
```

## Tecnologias

- **pandas** — leitura, limpeza e agregação dos dados
- **openpyxl** — geração do Excel formatado (cores, fontes, gráfico nativo)
- **matplotlib** — geração do gráfico em imagem

## Possíveis evoluções

- Suportar entrada em `.xlsx` além de `.csv`
- Adicionar filtro por período (mês/ano) via linha de comando
- Exportar também em PDF

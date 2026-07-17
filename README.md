# Excel Report Automation

Script Python que le uma planilha de vendas "crua" — com formatos de data
inconsistentes, categorias com grafia diferente, valores ausentes e linhas
duplicadas — e transforma em um relatorio Excel limpo, formatado e com
grafico, pronto para ser enviado a um cliente ou gestor.

## O que o script faz

1. **Le** os dados brutos de `data/vendas_brutas.csv`
2. **Limpa**:
   - remove espacos extras de texto
   - padroniza categorias (`PERIFERICOS`, `perifericos` → `Perifericos`)
   - unifica datas em formatos diferentes (`01/03/2026`, `2026-03-02`, `04-03-2026`)
   - preenche quantidade/valor ausentes com a mediana do produto (ou geral, como fallback)
   - remove linhas duplicadas
3. **Calcula** a receita total por categoria
4. **Gera**:
   - `output/relatorio_vendas.xlsx` — planilha com aba de Resumo (com grafico
     nativo do Excel) e aba de Dados Tratados, com cabecalhos formatados
   - `output/grafico_receita_categoria.png` — grafico de barras standalone

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

- **pandas** — leitura, limpeza e agregacao dos dados
- **openpyxl** — geracao do Excel formatado (cores, fontes, grafico nativo)
- **matplotlib** — geracao do grafico em imagem

## Possiveis evoluções

- Suportar entrada em `.xlsx` alem de `.csv`
- Adicionar filtro por periodo (mes/ano) via linha de comando
- Exportar tambem em PDF

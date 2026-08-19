# Resumo Saldo Contrato x Saldo de Empenho

Utilitário desktop em **Python + Tkinter** que consolida, em uma única planilha-resumo, os saldos de contrato e de empenho extraídos de múltiplas planilhas de obras.

## Funcionalidades

- Seleção múltipla de planilhas `.xlsx`/`.xlsm` via interface gráfica
- Localização automática da linha "TOTAL" (varrendo de baixo para cima) para extrair os saldos finais de cada obra
- Extração da "Data Base do Orçamento" a partir da aba "Dados da Obra"
- Geração de planilha-resumo formatada (cabeçalho estilizado, bordas, formato numérico, largura de colunas, congelamento de painel e auto-filtro)

## Tecnologias

- Python 3
- Tkinter (interface gráfica)
- openpyxl (leitura, escrita e formatação de Excel)

# 📊 DB - Controle de Despesas em Excel

Este projeto é um utilitário em **Python** para gerenciar um arquivo Excel de controle financeiro pessoal.  
Ele permite **adicionar novas linhas** com dados de gastos e visualizar a tabela diretamente no terminal.

---

## ⚙️ Funcionalidades

- Inserir novas entradas na planilha (`ControleMensal_v2.xlsx`):
  - Data (no formato `ddmm`, ex.: `2408` → 24/08/2025)
  - Tipo (ex.: `Out`, `In`)
  - Valor em Euro, se houver.
  - Valor em CHF
  - Local
  - Descrição
  - Categoria
- Loop interativo: você pode adicionar várias linhas de uma vez.
- Visualizar últimas linhas da tabela (`view_table.py`).
- Geração de executável `.exe` com **PyInstaller** para rodar sem precisar do Python.

---

## 📂 Estrutura do Projeto


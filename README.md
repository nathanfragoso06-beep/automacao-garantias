# 📊 MVP: Automação de Relatórios Gerenciais de Vendas e Estoque
**Projeto de Extensão Acadêmica — Análise e Desenvolvimento de Sistemas**

> Sistema desenvolvido em Python para automatizar a consolidação de dados operacionais e a geração de relatórios executivos mensais para a **Loja Hawara** (comércio varejista de vestuário), eliminando gargalos manuais e otimizando o controle de giro de estoque.

---

## 🚀 Sobre o Projeto
Na rotina de gestão de um comércio varejista, a consolidação manual de planilhas de vendas e o cruzamento com dados de estoque para identificar rupturas e produtos de alto giro consomem horas valiosas da equipe. 

Este projeto consiste em um **Produto Mínimo Viável (MVP)** de automação que processa os dados brutos de movimentação da Loja Hawara, aplica regras de agrupamento analítico e gera de forma automatizada uma apresentação executiva formatada.

---

## 🛠️ Tecnologias e Bibliotecas Utilizadas
* **Python 3.x** — Linguagem principal do script de automação.
* **Pandas** — Manipulação, limpeza, tratamento e agrupamento de dados tabulares.
* **Matplotlib** — Geração programática de gráficos analíticos e visuais.
* **Python-PPtx** — Automação e geração programática de slides de apresentação gerencial (.pptx).
* **OpenPyXL** — Leitura e escrita de arquivos de planilha (`.xlsx`).
---

## ⚙️ Arquitetura e Fluxo de Execução
1. **Ingetão:** O script lê a planilha de vendas e estoque mensal da Loja Hawara (`dados_loja_hawara.xlsx`).
2. **Processamento (ETL):** O `Pandas` limpa os dados, calcula indicadores de giro por categoria de vestuário e identifica itens críticos de estoque.
3. **Geração de Saída:** O motor em `python-pptx` preenche automaticamente os templates visuais e gera o relatório executivo pronto para a diretoria.

---

## 📁 Estrutura do Repositório
```text
├── src/
│   ├── main.py              # Script principal de automação
│   └── processamento.py     # Funções de manipulação com Pandas
├── data/
│   └── dados_hawara_exemplo.xlsx # Planilha fictícia de teste (Mock Data)
├── output/
│   └── relatorio_gerencial.pptx # Relatório gerencial gerado automaticamente
├── requirements.txt         # Dependências do projeto
└── README.md

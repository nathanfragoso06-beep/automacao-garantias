# Explicação
# MVP - Automação de Relatório de Insuficiência de Garantias

Repositório desenvolvido para o projeto prático da faculdade, contendo a solução de automação em Python para o processamento de relatórios e geração de apresentações executivas corporativas.

## 📁 Estrutura do Repositório
* `gerar_relatorio.py`: Script principal em Python que realiza a ingestão de dados, cruzamento temporal e geração dos slides executivos `.pptx` com gráficos por rede.
* `instalar_pendencias.bat`: Script de configuração automática do ambiente (verifica o Python, atualiza o pip e instala as bibliotecas necessárias).

## 🚀 Como Executar

1. **Configurar o ambiente:**
   Dê um duplo clique no arquivo `instalar_pendencias.bat` para instalar automaticamente as dependências (`pandas`, `python-pptx`, `matplotlib`, `openpyxl`).

2. **Gerar o relatório:**
   Execute o script Python na mesma pasta:
   ```bash
   python gerar_relatorio.py
## 📊 Saída Esperada
Geração automática do arquivo `Relatorio_Executivo_Garantias_Segmentos.pptx` com gráficos de barras horizontais por Rede.

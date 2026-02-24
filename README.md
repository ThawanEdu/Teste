# 📈 Automação de Atualização de Preços de Investimentos

Este projeto foi desenvolvido para automatizar a atualização de cotações e o gerenciamento de uma carteira de investimentos em Excel utilizando **Python** e a biblioteca **Pandas**.

## 🚀 Funcionalidades
- **Processamento de Dados:** Leitura automatizada de múltiplas abas de planilhas de investimentos (Renda Fixa, Renda Variável, Aportes).
- **Segurança de Dados:** Utilização de scripts para geração de dados sintéticos (Mock Data), permitindo a demonstração da lógica sem exposição de dados sensíveis ou financeiros reais.
- **Escalabilidade:** Estrutura preparada para integração com APIs de cotações em tempo real.

## 📂 Estrutura de Arquivos
- `Att_preços.py`: Script principal responsável pela lógica de processamento e atualização dos valores.
- `gerar_teste.py`: Utilitário para geração de uma planilha de exemplo (`investimentos_teste.xlsx`) com dados fictícios.
- `investimentos_teste.xlsx`: Planilha modelo utilizada para demonstrar o funcionamento do sistema.

## 🛠️ Tecnologias
- **Linguagem:** Python 3.x.
- **Biblioteca Principal:** Pandas (Manipulação de DataFrames).

## 🛡️ Aviso de Privacidade
Este repositório segue boas práticas de **CyberSegurança**. Arquivos contendo dados financeiros reais são ignorados via `.gitignore`, garantindo que apenas o ambiente de teste seja compartilhado publicamente.
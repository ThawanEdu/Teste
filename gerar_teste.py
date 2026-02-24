import pandas as pd

# Estrutura baseada nas abas que você me enviou
dados_teste = {
    'Aportes e Retiradas': {
        'Data': ['2026-01-10', '2026-02-15'],
        'Tipo': ['Aporte', 'Aporte'],
        'Valor': [500.00, 750.00]
    },
    'Renda Variável': {
        'Ticker': ['PETR4', 'VALE3', 'ITUB4'],
        'Quantidade': [10, 5, 15],
        'Preço Médio': [35.20, 68.00, 27.50]
    },
    'Renda Fixa': {
        'Ativo': ['CDB Banco X', 'Tesouro IPCA+'],
        'Valor Aplicado': [2000.00, 3000.00],
        'Vencimento': ['2027-12-31', '2029-05-15']
    },
    'Dashboard(Resumo)': {
        'Categoria': ['Renda Fixa', 'Renda Variável', 'Liquidez'],
        'Total': [5000.00, 1500.00, 1000.00]
    }
}

# Criando o Excel com as abas corretas
with pd.ExcelWriter('investimentos_teste.xlsx') as writer:
    for aba, info in dados_teste.items():
        pd.DataFrame(info).to_excel(writer, sheet_name=aba, index=False)

print("✅ Planilha 'investimentos_teste.xlsx' gerada com sucesso para portfólio!")
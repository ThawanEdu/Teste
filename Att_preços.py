import pandas as pd
import yfinance as yf
import os

def atualizar_planilha(arquivo_excel):
    nome_aba = 'Renda Variável'  # Nome exato da aba onde estão os ativos
    
    try:
        print(f"📂 Abrindo arquivo: {arquivo_excel}")
        
        # O SEGREDO: Especificar a sheet_name correta
        df = pd.read_excel(arquivo_excel, sheet_name=nome_aba)
        
        # Remove espaços extras dos nomes das colunas
        df.columns = [str(c).strip() for c in df.columns]

        if 'Ativo' not in df.columns:
            print(f"❌ ERRO: Coluna 'Ativo' não encontrada na aba '{nome_aba}'.")
            print(f"Colunas encontradas: {list(df.columns)}")
            return

        print(f"📊 Total de ativos para atualizar: {len(df)}")

        def buscar_preco(ticker):
            ticker = str(ticker).strip().upper()
            if not ticker or ticker == "NAN": return None
            
            try:
                # Adiciona .SA se for ação/FII brasileira
                ticker_full = f"{ticker}.SA" if not ticker.endswith(".SA") else ticker
                papel = yf.Ticker(ticker_full)
                hist = papel.history(period="7d")
                
                if not hist.empty:
                    preco = hist['Close'].iloc[-1]
                    print(f"   ✅ {ticker_full}: R$ {preco:.2f}")
                    return round(float(preco), 2)
                else:
                    return None
            except:
                return None

        # Atualizando os preços
        df['Preço Atual'] = df['Ativo'].apply(buscar_preco)

        # SALVANDO SEM APAGAR AS OUTRAS ABAS
        with pd.ExcelWriter(arquivo_excel, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)
            
        print(f"\n✨ Aba '{nome_aba}' atualizada com sucesso!")

    except Exception as e:
        print(f"💥 Erro: {e}")

# Caminho do seu arquivo (ajuste se necessário)
meu_arquivo = r"C:\Users\thawa\OneDrive\Desktop\Projetos\financas.py\Investimentos.xlsx"
atualizar_planilha(meu_arquivo)
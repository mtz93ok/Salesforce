import streamlit as st
import pandas as pd
import requests
from io import BytesIO, StringIO

st.title("🔍 Consulta de Relatório por Nome")

# Entrada do nome do colaborador
nome = st.text_input("Digite seu nome exatamente como aparece no relatório (coluna 'Proprietário da conta')")

# URL fixa do relatório
salesforce_url = "https://secil.my.salesforce.com/00O7S000001kByi?export=1&enc=UTF-8&xf=csv"

if nome:
    try:
        st.info("🔄 Baixando e processando o relatório...")

        response = requests.get(salesforce_url)
        response.raise_for_status()

        # Força o encoding ISO-8859-1
        df = pd.read_csv(BytesIO(response.content), encoding="ISO-8859-1")

        # Corrige espaços extras nos nomes de colunas
        df.columns = df.columns.str.strip()

        # Exibe as colunas detectadas (para debug)
        # st.write("Colunas:", df.columns.tolist())

        # Verifica se a coluna existe
        if "Proprietário da conta" not in df.columns:
            st.error("⚠️ Coluna 'Proprietário da conta' não encontrada no relatório.")
        else:
            df_filtrado = df[df["Proprietário da conta"] == nome]

            if df_filtrado.empty:
                st.warning("Nenhum resultado encontrado para esse nome.")
            else:
                st.success(f"{len(df_filtrado)} linha(s) encontradas para {nome}.")
                st.dataframe(df_filtrado)

                # Converte para CSV e oferece botão de download
                csv = df_filtrado.to_csv(index=False, sep=';', encoding="ISO-8859-1")
                st.download_button(
                    label="📥 Baixar relatório filtrado (.csv)",
                    data=csv,
                    file_name=f"relatorio_{nome}.csv",
                    mime="text/csv"
                )

    except Exception as e:
        st.error(f"Erro: {e}")


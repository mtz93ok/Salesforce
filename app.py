import streamlit as st
import pandas as pd
import requests
from io import BytesIO, StringIO

REPORT_URL = "https://secil.my.salesforce.com/00O7S000001kByi?export=1&enc=UTF-8&xf=csv"

st.title("🔎 Exportar Relatório do Salesforce")

nome = st.text_input("Digite seu nome (Proprietário da conta)")
sid = st.text_input("SID do Salesforce", type="password")

if st.button("Gerar e baixar relatório"):
    if not nome or not sid:
        st.warning("Preencha todos os campos.")
    else:
        headers = {"Authorization": f"Bearer {sid}"}
        response = requests.get(REPORT_URL, headers=headers)

        if response.status_code == 200:
            try:
                # Lê o CSV corretamente como UTF-8
                df = pd.read_csv(BytesIO(response.content), encoding="utf-8")
                df.columns = df.columns.str.strip().str.lower()
                nome = nome.strip().lower()

                # Detecta a coluna correta do proprietário
                coluna_proprietario = [col for col in df.columns if "propriet" in col and "conta" in col]
                if not coluna_proprietario:
                    st.error("Coluna 'Proprietário da conta' não encontrada.")
                else:
                    col = coluna_proprietario[0]
                    df[col] = df[col].astype(str).str.strip().str.lower()
                    df_filtrado = df[df[col] == nome]

                    if df_filtrado.empty:
                        st.warning("Nenhum dado encontrado para esse nome.")
                    else:
                        st.dataframe(df_filtrado)

                        # Exporta para ISO-8859-1 para compatibilidade com Excel
                        csv_buffer = StringIO()
                        df_filtrado.to_csv(csv_buffer, index=False, encoding="ISO-8859-1")
                        csv_data = csv_buffer.getvalue()

                        st.download_button(
                            label="📥 Baixar CSV em ISO-8859-1",
                            data=csv_data.encode("ISO-8859-1"),
                            file_name=f"relatorio_{nome}.csv",
                            mime="text/csv"
                        )

            except Exception as e:
                st.error(f"Erro ao processar o CSV: {e}")
        else:
            st.error("Erro ao baixar o relatório. Verifique o SID.")

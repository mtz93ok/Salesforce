import streamlit as st
import pandas as pd
import requests
from io import BytesIO, StringIO

# URL fixa do relatório Salesforce
REPORT_URL = "https://secil.my.salesforce.com/00O7S000001kByi?export=1&enc=UTF-8&xf=csv"

st.title("🔎 Exportar Relatório do Salesforce")

# Entradas do usuário
nome = st.text_input("Digite seu nome (Proprietário da conta)")
sid = st.text_input("SID do Salesforce", type="password")

if st.button("Gerar e baixar relatório"):
    if not nome or not sid:
        st.warning("Por favor, preencha o nome e o SID.")
    else:
        # Faz o request autenticado com o SID
        headers = {
            "Authorization": f"Bearer {sid}"
        }
        response = requests.get(REPORT_URL, headers=headers)

        if response.status_code == 200:
            try:
                df = pd.read_csv(BytesIO(response.content), encoding='utf-8')
                st.write("Pré-visualização dos dados (com filtro):")

                # Filtra pelo nome na coluna correta
                df_filtrado = df[df["Proprietário da conta"] == nome]

                if df_filtrado.empty:
                    st.warning("Nenhum resultado encontrado para esse nome.")
                else:
                    st.dataframe(df_filtrado)

                    # Converte para CSV em memória
                    csv_buffer = StringIO()
                    df_filtrado.to_csv(csv_buffer, index=False)
                    csv_data = csv_buffer.getvalue()

                    # Botão para baixar CSV
                    st.download_button(
                        label="📥 Baixar relatório filtrado",
                        data=csv_data,
                        file_name=f"relatorio_{nome}.csv",
                        mime="text/csv"
                    )
            except Exception as e:
                st.error(f"Erro ao processar o CSV: {e}")
        else:
            st.error("Erro ao baixar o relatório. Verifique se o SID está válido.")

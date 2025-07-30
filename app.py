import streamlit as st
import pandas as pd
import requests
from io import BytesIO

st.set_page_config(page_title="Download de Relatório Salesforce", layout="centered")
st.title("🔽 Baixar Relatório Individual - Salesforce")

# Formulário
with st.form("relatorio_form"):
    nome = st.text_input("Digite seu nome (como aparece no Salesforce)", max_chars=50)
    email = st.text_input("Digite seu e-mail", max_chars=100)
    sid = st.text_input("Cole aqui seu SID (você precisa estar logado no Salesforce)", type="password")
    submitted = st.form_submit_button("🔍 Gerar Relatório")

if submitted:
    if not nome or not sid:
        st.error("Por favor, preencha todos os campos.")
    else:
        try:
            # URL do relatório
            url = "https://secil.my.salesforce.com/00O7S000001kByi?export=1&enc=UTF-8&xf=csv"
            headers = {"Authorization": f"Bearer {sid}"}
            response = requests.get(url, headers=headers)

            if response.status_code != 200:
                st.error("Erro ao baixar o relatório. Verifique se o SID está correto e se você está logado.")
            else:
                # Verificar se resposta HTML (erro de autenticação)
                content = response.content.decode("utf-8", errors="ignore")
                if "<html" in content.lower():
                    st.error("O SID fornecido é inválido ou expirou. Faça login no Salesforce e cole o SID válido.")
                else:
                    # Ler CSV com encoding correto
                    df = pd.read_csv(BytesIO(response.content), encoding="utf-8", sep=",")
                    
                    # Tentar encontrar a coluna do proprietário
                    coluna_proprietario = next((col for col in df.columns if "Propriet" in col), None)

                    if not coluna_proprietario:
                        st.error("Coluna 'Proprietário da conta' não encontrada no relatório.")
                    else:
                        # Filtrar pelo nome informado
                        df_filtrado = df[df[coluna_proprietario].astype(str).str.strip().str.lower() == nome.strip().lower()]
                        
                        if df_filtrado.empty:
                            st.warning("Nenhum dado encontrado para esse nome.")
                        else:
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                df_filtrado.to_excel(writer, index=False, sheet_name="Relatório")
                            output.seek(0)

                            st.success("Relatório gerado com sucesso!")
                            st.download_button(
                                label="📥 Baixar relatório (.xlsx)",
                                data=output,
                                file_name=f"relatorio_{nome.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o relatório: {str(e)}")



import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(page_title="Download de Relatório Salesforce", layout="centered")
st.title("📥 Baixar Relatório Individual")

# Formulário de entrada
with st.form("relatorio_form"):
    nome = st.text_input("Digite seu nome (como aparece no relatório)", max_chars=100)
    submitted = st.form_submit_button("🔍 Buscar dados")

if submitted:
    if not nome:
        st.error("Por favor, preencha o nome.")
    else:
        try:
            # Caminho UNC da rede (substitui o mapeamento K:\)
            caminho_arquivo = r"\\pom-srv-fs-01\sc\BI_IM_POM\Fontes de Dados XLS\pd_em_massa\pdemmassa.csv"

            # # Verifica se o arquivo existe
            # if not os.path.exists(caminho_arquivo):
            #     st.error("Arquivo não encontrado no caminho da rede.")
            # else:
            #     # Lê o CSV
            #     df = pd.read_csv(caminho_arquivo, encoding='utf-8')

                # Tenta identificar uma coluna que contenha "Proprietário" ou similar
                coluna_proprietario = next((col for col in df.columns if "Proprietário" in col or "Nome" in col), None)

                if not coluna_proprietario:
                    st.error("Não foi possível identificar a coluna de nome no relatório.")
                else:
                    # Filtra pelo nome informado (sem acento e ignorando maiúsculas)
                    df_filtrado = df[df[coluna_proprietario].str.strip().str.lower() == nome.strip().lower()]

                    if df_filtrado.empty:
                        st.warning("Nenhum dado encontrado para esse nome.")
                    else:
                        # Converter para XLSX
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df_filtrado.to_excel(writer, index=False, sheet_name="Relatório")

                        output.seek(0)
                        st.success("Relatório gerado com sucesso!")

                        # Botão de download
                        st.download_button(
                            label="📥 Baixar arquivo (.xlsx)",
                            data=output,
                            file_name=f"relatorio_{nome.replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

        except Exception as e:
            st.error("Erro ao processar o relatório.")
            st.exception(e)




import streamlit as st
import pandas as pd

st.title("PD EM MASSA")

# Input nome do cidadão
nome = st.text_input("Digite o nome do cidadão para filtro")

# Caminho do arquivo na rede - ajuste para o caminho correto da sua rede
caminho_arquivo = r"K:\BI_IM_POM\Fontes de Dados XLS\pd_em_massa\pdemmassa.csv"  # Exemplo Windows
# caminho_arquivo = "/mnt/rede/relatorio.xlsx"  # Exemplo Linux

if st.button("Filtrar e baixar"):
    try:
        # Ler o arquivo Excel (ou CSV)
        df = pd.read_excel(caminho_arquivo)  # ou pd.read_csv(caminho_arquivo) se for CSV

        # Filtrar pela coluna "Proprietário da conta" - ajuste o nome conforme seu arquivo
        coluna_proprietario = next((c for c in df.columns if "Propriet" in c), None)
        if not coluna_proprietario:
            st.error("Coluna 'Proprietário da conta' não encontrada no arquivo.")
        else:
            df_filtrado = df[df[coluna_proprietario].astype(str).str.contains(nome, case=False, na=False)]

            if df_filtrado.empty:
                st.warning("Nenhum registro encontrado para esse nome.")
            else:
                st.dataframe(df_filtrado)

                # Gerar Excel para download
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_filtrado.to_excel(writer, index=False, sheet_name="Filtrado")
                output.seek(0)

                st.download_button(
                    label="📥 Baixar dados filtrados (.xlsx)",
                    data=output,
                    file_name=f"relatorio_filtrado_{nome.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo da rede: {e}")




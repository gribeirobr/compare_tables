import streamlit as st
import pandas as pd
import datacompy
import re
import io
from base64 import b64encode

def limpar_texto(texto):
    if isinstance(texto, str):
        texto = texto.replace('\u00A0', ' ')
        texto = re.sub(r"[^\x20-\x7EÀ-ÿ]", "", texto)
        return texto.strip()
    return texto

def gerar_download(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        df.to_excel(writer, index=False)
    b64 = b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Baixar {filename}</a>'
    return href

st.title("🔍 Comparador de Abas do Google Sheets (com seleção de colunas)")

spreadsheet_id = st.text_input("📄 ID da Planilha", "")
gid_original = st.text_input("📑 GID da Aba Original", "")
gid_atualizado = st.text_input("📑 GID da Aba Atualizada", "")

carregado = False

if spreadsheet_id and gid_original and gid_atualizado:
    try:
        url_original = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid_original}"
        url_atualizado = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid_atualizado}"

        df_original = pd.read_csv(url_original, encoding='utf-8-sig')
        df_atualizado = pd.read_csv(url_atualizado, encoding='utf-8-sig')

        st.success("✅ Abas carregadas com sucesso!")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Colunas da Tabela A**")
            colunas_original = st.multiselect("Selecione as colunas de junção da tabela A", df_original.columns)

        with col2:
            st.markdown("**Colunas da Tabela B**")
            colunas_atualizado = st.multiselect("Selecione as colunas correspondentes da tabela B", df_atualizado.columns)

        if len(colunas_original) == len(colunas_atualizado) and len(colunas_original) > 0:
            if st.button("🔎 Comparar"):
                try:
                    column_mapping = dict(zip(colunas_original, colunas_atualizado))

                    comparacao = datacompy.Compare(
                        df1=df_original,
                        df2=df_atualizado,
                        join_columns=column_mapping,
                        df1_name='Original',
                        df2_name='Atualizado',
                        abs_tol=0,
                        rel_tol=0
                    )

                    df1_unq = comparacao.df1_unq_rows
                    df2_unq = comparacao.df2_unq_rows

                    for df in [df1_unq, df2_unq]:
                        for col in df.select_dtypes(include='object').columns:
                            df[col] = df[col].apply(limpar_texto)

                    st.subheader("🔸 Linhas somente na aba original")
                    st.dataframe(df1_unq)
                    st.markdown(gerar_download(df1_unq, "somente_original.xlsx"), unsafe_allow_html=True)

                    st.subheader("🔹 Linhas somente na aba atualizada")
                    st.dataframe(df2_unq)
                    st.markdown(gerar_download(df2_unq, "somente_atualizado.xlsx"), unsafe_allow_html=True)

                except Exception as e:
                    st.error(f"Erro ao comparar: {e}")
        elif len(colunas_original) != len(colunas_atualizado):
            st.warning("⚠️ As listas de colunas devem ter o mesmo número de elementos.")

    except Exception as e:
        st.error(f"Erro ao carregar as abas: {e}")


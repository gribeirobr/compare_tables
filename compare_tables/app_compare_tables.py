import streamlit as st
import pandas as pd
import datacompy
import re
import io
from base64 import b64encode

# Função para limpar texto
def limpar_texto(texto):
    if isinstance(texto, str):
        texto = texto.replace('\u00A0', ' ')
        texto = re.sub(r"[^\x20-\x7EÀ-ÿ]", "", texto)
        return texto.strip()
    return texto

# Função principal de comparação
def comparar_abas(spreadsheet_id, gid_original, gid_atualizado):
    url_original = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid_original}"
    url_atualizado = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid_atualizado}"

    df_original = pd.read_csv(url_original, encoding='utf-8-sig')
    df_atualizado = pd.read_csv(url_atualizado, encoding='utf-8-sig')

    # Colunas comuns
    colunas_comuns = list(set(df_original.columns) & set(df_atualizado.columns))
    if not colunas_comuns:
        raise ValueError("Não há colunas comuns entre as abas para realizar a junção.")

    # Comparação com pandas
    comparacao = datacompy.Compare(
        df1=df_original,
        df2=df_atualizado,
        join_columns=colunas_comuns,
        df1_name='Original',
        df2_name='Atualizado',
        abs_tol=0,
        rel_tol=0,
        ignore_extra_columns=False
    )

    df1_unq = comparacao.df1_unq_rows
    df2_unq = comparacao.df2_unq_rows

    for df in [df1_unq, df2_unq]:
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].apply(limpar_texto)

    return df1_unq, df2_unq

# Função para gerar link de download
def gerar_download(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        df.to_excel(writer, index=False)
    b64 = b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Baixar {filename}</a>'
    return href

# Streamlit App
st.title("📊 Comparador de Abas do Google Sheets")

st.markdown("Compare duas abas de uma mesma planilha do Google Sheets via `spreadsheet_id` e `gid`.")

spreadsheet_id = st.text_input("🆔 ID da Planilha", "")
gid_original = st.text_input("📄 GID da Aba onde está a Tabela A", "")
gid_atualizado = st.text_input("📄 GID da Aba onde está a Tabela B", "")

if st.button("🔍 Comparar"):
    if spreadsheet_id and gid_original and gid_atualizado:
        with st.spinner("Comparando..."):
            try:
                df1, df2 = comparar_abas(spreadsheet_id, gid_original, gid_atualizado)
                st.success("✅ Comparação concluída!")

                st.subheader("🔸 Linhas somente na Tabela A")
                st.dataframe(df1)
                st.markdown(gerar_download(df1, "linhas_somente_tabela_a.xlsx"), unsafe_allow_html=True)

                st.subheader("🔹 Linhas somente na Tabela B")
                st.dataframe(df2)
                st.markdown(gerar_download(df2, "linhas_somente_tabela_b.xlsx"), unsafe_allow_html=True)
            except Exception as e:
                st.error(f"❌ Erro: {e}")
    else:
        st.warning("⚠️ Preencha todos os campos para continuar.")

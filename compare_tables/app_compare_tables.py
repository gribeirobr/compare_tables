import streamlit as st
import pandas as pd
import datacompy
import re
import io
from pyspark.sql import SparkSession
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
    # Spark session
    spark = SparkSession.builder.appName("CompararPlanilhas").getOrCreate()

    url_original = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid_original}"
    url_atualizado = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid_atualizado}"

    df_original = pd.read_csv(url_original, encoding='utf-8-sig')
    df_atualizado = pd.read_csv(url_atualizado, encoding='utf-8-sig')

    df_A = spark.createDataFrame(df_original)
    df_B = spark.createDataFrame(df_atualizado)

    join_cols = list(set(df_original.columns) & set(df_atualizado.columns))
    if not join_cols:
        raise ValueError("Não há colunas comuns entre as abas para realizar a junção.")

    comparacao = datacompy.SparkSQLCompare(
        spark,
        df_A,
        df_B,
        join_columns=join_cols
    )

    df1 = comparacao.df1_unq_rows.toPandas()
    df2 = comparacao.df2_unq_rows.toPandas()

    for df in [df1, df2]:
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].apply(limpar_texto)

    return df1, df2

# Função para gerar link de download
def gerar_download(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        df.to_excel(writer, index=False)
    b64 = b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Baixar {filename}</a>'
    return href

# Streamlit App
st.title("🧾 Comparador de Abas do Google Sheets")

spreadsheet_id = st.text_input("ID da Planilha", "")
gid_original = st.text_input("GID da Aba Original", "")
gid_atualizado = st.text_input("GID da Aba Atualizada", "")

if st.button("Comparar"):
    if spreadsheet_id and gid_original and gid_atualizado:
        with st.spinner("Comparando..."):
            try:
                df1, df2 = comparar_abas(spreadsheet_id, gid_original, gid_atualizado)
                st.success("Comparação realizada!")

                st.subheader("📄 Linhas somente na aba original")
                st.write(df1)
                st.markdown(gerar_download(df1, "somente_original.xlsx"), unsafe_allow_html=True)

                st.subheader("📄 Linhas somente na aba atualizada")
                st.write(df2)
                st.markdown(gerar_download(df2, "somente_atualizado.xlsx"), unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Erro: {e}")
    else:
        st.warning("Preencha todos os campos.")

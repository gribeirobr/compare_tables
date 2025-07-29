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
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Baixar {filename}</a>'
    return href

def dados_preenchidos():
    return all([spreadsheet_id, gid_original, gid_atualizado])


USUARIOS_CADASTRADOS = {
    "admin": "1234",
    "guilherme": "senha_guilherme",
    "GABI_REZENDE": "gabiR#123",
    "ana": "abc123"
}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None

if not st.session_state.autenticado:
    st.subheader("Login")
    usuario = st.text_input("Usuário")
    senha = st.text_input("Senha", type="password")

    if st.button("Entrar"):
        if usuario in USUARIOS_CADASTRADOS and senha == USUARIOS_CADASTRADOS[usuario]:
            st.session_state.autenticado = True
            st.session_state.usuario_logado = usuario
            st.success(f"Bem-vindo, {usuario}!")
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")
    st.stop()

if st.session_state.autenticado:
    st.title("Comparador de Tabelas do Google Sheets")

    col1, col2 = st.columns([3,1])
    with col1:
        st.markdown(f"👤 Usuário logado: **{st.session_state.usuario_logado}**")
    with col2:
        if st.button("Logout"):
            st.session_state.autenticado = False
            st.session_state.usuario_logado = None
            st.rerun()

    spreadsheet_id = st.text_input("ID da Planilha", "")
    gid_original = st.text_input("GID da Aba da Tabela A", "")
    gid_atualizado = st.text_input("GID da Aba da Tabela B", "")

    def dados_preenchidos():
        return all([spreadsheet_id, gid_original, gid_atualizado])

    if "manual_mode" not in st.session_state:
        st.session_state.manual_mode = False

    if not st.session_state.manual_mode:
        if st.button("Selecionar colunas manualmente"):
            if not dados_preenchidos():
                st.warning("Preencha o ID da planilha e os GIDs antes de selecionar colunas.")
            else:
                st.session_state.manual_mode = True

    if st.session_state.manual_mode:
        st.markdown("### Selecione as colunas de junção manualmente")

    if dados_preenchidos():
        try:
            url_original = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid_original}"
            url_atualizado = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid_atualizado}"

            df_original = pd.read_csv(url_original, encoding='utf-8-sig')
            df_atualizado = pd.read_csv(url_atualizado, encoding='utf-8-sig')

            st.success("Abas carregadas com sucesso!")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**Colunas da Tabela A**")
                colunas_original = st.multiselect("Selecione as colunas de junção da tabela A", df_original.columns)

            with col2:
                st.markdown("**Colunas da Tabela B**")
                colunas_atualizado = st.multiselect("Selecione as colunas correspondentes da tabela B", df_atualizado.columns)

            if st.button("Comparar"):
                if not dados_preenchidos():
                    st.warning("Preencha o ID da planilha e os GIDs antes de comparar.")
                elif len(colunas_original) != len(colunas_atualizado) or len(colunas_original) == 0:
                    st.warning("As listas de colunas devem ter o mesmo número de elementos.")
                else:
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
                        df_intersecao = comparacao.intersect_rows

                        for df in [df1_unq, df2_unq, df_intersecao]:
                            for col in df.select_dtypes(include='object').columns:
                                df[col] = df[col].apply(limpar_texto)

                        st.subheader("Registros somente na tabela A")
                        st.dataframe(df1_unq)
                        st.markdown(gerar_download(df1_unq, "registros_somente_tabela_a.xlsx"), unsafe_allow_html=True)

                        st.subheader("Registros somente na tabela B")
                        st.dataframe(df2_unq)
                        st.markdown(gerar_download(df2_unq, "registros_somente_tabela_b.xlsx"), unsafe_allow_html=True)

                        st.subheader("Registros iguais nas duas tabelas")
                        st.dataframe(df_intersecao)
                        st.markdown(gerar_download(df_intersecao, "registros_conciliados.xlsx"), unsafe_allow_html=True)

                        total_iguais = df_intersecao.shape[0]
                        total_diferentes = df1_unq.shape[0] + df2_unq.shape[0]

                        st.markdown("### Resumo da Comparação")
                        st.markdown(f"- **Linhas iguais:** {total_iguais}")
                        st.markdown(f"- **Linhas diferentes:** {total_diferentes}")

                    except Exception as e:
                        st.error(f"Erro ao comparar: {e}")
            elif len(colunas_original) != len(colunas_atualizado):
                st.warning("As listas de colunas devem ter o mesmo número de elementos.")

        except Exception as e:
            st.error(f"Erro ao carregar as abas: {e}")

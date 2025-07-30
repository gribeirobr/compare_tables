import streamlit as st
import pandas as pd
import io

USUARIOS_CADASTRADOS = {
    "admin": "1234",
    "guilherme": "senha_guilherme",
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
    st.set_page_config(
        page_title="Renomeador de Colunas de Planilhas Google",
        page_icon="✨",
        layout="wide"
    )

    st.title("Renomear Colunas de Planilhas Google")
    st.markdown("""
    Esta aplicação permite que você carregue uma aba de uma Planilha Google, 
    selecione as colunas que deseja manter e as renomeie de forma fácil e rápida.
    """)
        
    st.set_page_config(
        page_title="Renomeador de Colunas de Planilhas Google",
        page_icon="✨",
        layout="wide"
    )


    @st.cache_data
    def carregar_planilha(sheet_id, gid):
        """
        Carrega os dados de uma Planilha Google pública usando o ID e o GID.
        """
        try:
            url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}'
            df = pd.read_csv(url)
            return df, None 
        except Exception as e:
            return None, f"Ocorreu um erro ao carregar a planilha: {e}"

    @st.cache_data
    def to_excel(df):
        """
        Converte um DataFrame para um arquivo Excel em bytes.
        """
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Dados')
        processed_data = output.getvalue()
        return processed_data


    if 'df_original' not in st.session_state:
        st.session_state.df_original = None
    if 'df_final' not in st.session_state:
        st.session_state.df_final = None


    st.header("1. Informe os dados da sua planilha")
    st.warning(
        "*Importante:* Sua planilha precisa estar com o compartilhamento definido como '*Qualquer pessoa com o link*' para que a aplicação consiga acessá-la.",
        icon="⚠️"
    )

    col1, col2 = st.columns(2)

    with col1:
        google_sheet_id = st.text_input(
            "Cole o ID da sua Planilha Google",
            help="O ID é a parte longa no meio da URL da sua planilha. Ex: 1aBcDeFgHiJkLmNoPqRsTuVwXyZ_1234567890"
        )

    with col2:
        google_sheet_gid = st.text_input(
            "Cole o GID da aba que você quer usar",
            help="O GID fica no final da URL, depois de '#gid='. Se não houver, geralmente é '0' para a primeira aba."
        )

    if st.button("Carregar Planilha", type="primary"):
        if google_sheet_id and google_sheet_gid:
            with st.spinner("Buscando e carregando os dados..."):
                st.session_state.df_original = None
                st.session_state.df_final = None

                df, erro = carregar_planilha(google_sheet_id, google_sheet_gid)
                if df is not None:
                    st.session_state.df_original = df
                    st.success("Planilha carregada com sucesso!")
                else:
                    st.error(erro)
        else:
            st.error("Por favor, preencha o ID e o GID da planilha.")

    if st.session_state.df_original is not None:
        st.header("2. Selecione e renomeie as colunas")
        
        df_original = st.session_state.df_original
        colunas_originais = df_original.columns.tolist()

        st.markdown("Primeiro, escolha as colunas que você deseja manter na nova planilha:")
        colunas_selecionadas = st.multiselect(
            "Selecione as colunas:",
            options=colunas_originais,
            label_visibility="collapsed" 
        )

        mapa_renomeacao = {}

        if colunas_selecionadas:
            st.markdown("Agora, defina os novos nomes para as colunas selecionadas. (Deixe como está se não quiser renomear).")
            
            col_nome_original, col_nome_novo = st.columns(2)
            col_nome_original.markdown("*Nome Original*")
            col_nome_novo.markdown("*Novo Nome*")

            for coluna in colunas_selecionadas:
                with col_nome_original:
                    st.text(coluna) 
                with col_nome_novo:
                    novo_nome = st.text_input(f"novo_nome_{coluna}", value=coluna, label_visibility="collapsed")
                    mapa_renomeacao[coluna] = novo_nome
            
            if st.button("Gerar Nova Planilha", type="primary"):
                with st.spinner("Processando..."):
                    df_filtrado = df_original[colunas_selecionadas]
                    
                    df_renomeado = df_filtrado.rename(columns=mapa_renomeacao)
                    
                    st.session_state.df_final = df_renomeado

    if st.session_state.df_final is not None:
        st.header("3. Resultado Final")
        st.success("Sua nova planilha está pronta!")
        
        df_final = st.session_state.df_final
        st.dataframe(df_final)

        st.subheader("Faça o download do resultado")

        csv_data = df_final.to_csv(index=False).encode('utf-8')
        excel_data = to_excel(df_final)  

        dl_col1, dl_col2 = st.columns(2)

        with dl_col1:
            st.download_button(
            label="📥 Baixar como CSV",
            data=csv_data,
            file_name='planilha_renomeada.csv',
            mime='text/csv',
            use_container_width=True
            )

        with dl_col2:
            st.download_button(
            label="📊 Baixar como XLSX (Excel)",
            data=excel_data,
            file_name='planilha_renomeada.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True
            )
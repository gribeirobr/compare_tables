import streamlit as st
import pandas as pd
import io

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
    # Configuração da página
    st.set_page_config(page_title="Filtro e Soma de Planilhas", page_icon="📊", layout="wide")
    st.title("📊 Filtro e Soma de Planilha Google Sheets")

    # Entrada de dados do usuário
    with st.expander("🔑 Informações da Planilha", expanded=True):
        spreadsheet_id = st.text_input("ID da Planilha", "")
        gid = st.text_input("GID da Aba", "")
        st.caption("Exemplo de URL: `https://docs.google.com/spreadsheets/d/[ID_AQUI]/edit#gid=[GID_AQUI]`")

    # Carregar dados da planilha
    def carregar_dados(spreadsheet_id, gid):
        try:
            url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"
            df = pd.read_csv(url, encoding='utf-8-sig')
            st.session_state.df_original = df
            return df
        except Exception as e:
            st.error(f"Erro ao carregar dados: {str(e)}")
            return None

    # Processar dados e gerar resultados
    def processar_dados(df, coluna_filtro, valor_filtro, coluna_soma):
        try:
            # Filtragem dos dados
            df_filtrado = df[df[coluna_filtro].astype(str) == str(valor_filtro)]
            
            # Cálculo da soma
            soma = df_filtrado[coluna_soma].sum()
            
            return df_filtrado, soma
        except KeyError:
            st.error("Erro: Uma ou mais colunas selecionadas não existem no DataFrame")
            return None, None
        except Exception as e:
            st.error(f"Erro no processamento: {str(e)}")
            return None, None

    # Gerar arquivo para download
    def gerar_download(df, formato):
        if formato == 'CSV':
            return df.to_csv(index=False).encode('utf-8-sig')
        elif formato == 'XLSX':
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Dados Filtrados')
            return output.getvalue()

    # Interface principal
    if st.button("Carregar Planilha") or 'df_original' in st.session_state:
        if not spreadsheet_id or not gid:
            st.warning("⚠️ Por favor, informe o ID da planilha e o GID da aba")
        else:
            df = carregar_dados(spreadsheet_id, gid)
            
            if df is not None:
                st.success(f"✅ Dados carregados com sucesso! ({df.shape[0]} linhas, {df.shape[1]} colunas)")
                
                # Seleção das colunas
                col1, col2 = st.columns(2)
                
                with col1:
                    coluna_filtro = st.selectbox("Coluna para filtrar:", df.columns)
                    # Obter valores únicos para o filtro
                    valores_unicos = df[coluna_filtro].unique()
                    valor_filtro = st.selectbox(f"Valor para filtrar em '{coluna_filtro}':", valores_unicos)
                
                with col2:
                    colunas_numericas = df.select_dtypes(include=['number']).columns
                    if len(colunas_numericas) == 0:
                        st.warning("⚠️ Nenhuma coluna numérica encontrada para soma!")
                        coluna_soma = None
                    else:
                        coluna_soma = st.selectbox("Coluna para somar:", colunas_numericas)
                
                if st.button("Aplicar Filtro e Calcular Soma") and coluna_soma is not None:
                    df_filtrado, soma = processar_dados(df, coluna_filtro, valor_filtro, coluna_soma)
                    
                    if df_filtrado is not None:
                        st.subheader("Resultados")
                        
                        # Mostrar resultado da soma
                        st.metric(f"Soma de '{coluna_soma}'", f"{soma:,.2f}")
                        
                        # Mostrar tabela filtrada
                        st.dataframe(df_filtrado)
                        
                        # Download dos dados
                        st.subheader("Download dos Dados Filtrados")
                        col3, col4 = st.columns(2)
                        
                        with col3:
                            csv = gerar_download(df_filtrado, 'CSV')
                            st.download_button(
                                label="Baixar como CSV",
                                data=csv,
                                file_name='dados_filtrados.csv',
                                mime='text/csv'
                            )
                        
                        with col4:
                            excel = gerar_download(df_filtrado, 'XLSX')
                            st.download_button(
                                label="Baixar como Excel",
                                data=excel,
                                file_name='dados_filtrados.xlsx',
                                mime='application/vnd.ms-excel'
                            )

    # Instruções de uso
    st.markdown("""
    ---

    ### 📌 Instruções de Uso:
    1. Obtenha o **ID da Planilha** e **GID da Aba** da URL do Google Sheets:
    - Formato da URL: `https://docs.google.com/spreadsheets/d/[ID_DA_PLANILHA]/edit#gid=[GID_DA_ABA]`
    2. Cole os valores nos campos acima
    3. Clique em **"Carregar Planilha"**
    4. Selecione a coluna para filtrar e o valor desejado
    5. Escolha a coluna numérica para soma
    6. Clique em **"Aplicar Filtro e Calcular Soma"**
    7. Faça o download dos resultados nos formatos disponíveis

    > ⚠️ A planilha precisa estar configurada com acesso público (qualquer pessoa com o link pode visualizar)
    """)

    # Rodapé
    st.caption("Desenvolvido com Streamlit | [Documentação do Streamlit](https://docs.streamlit.io/)")
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
    st.title("Filtro e Soma de Dados de Arquivo CSV")

    st.markdown("---")

    # Função para carregar dados de arquivo CSV
    @st.cache_data(ttl=600)  # Cache os dados por 10 minutos
    def load_data_from_csv(uploaded_file):
        try:
            # Ler o arquivo CSV
            df = pd.read_csv(uploaded_file)
            
            # Remover linhas e colunas completamente vazias
            df.dropna(how='all', inplace=True)
            df.dropna(axis=1, how='all', inplace=True)
            
            # Se houver cabeçalhos em branco, renomeá-los para evitar conflitos
            df.columns = [f"Coluna_{i+1}" if col is None or str(col).strip() == "" else col for i, col in enumerate(df.columns)]
            
            return df
        except Exception as e:
            st.error(f"Ocorreu um erro ao carregar os dados: {e}")
            return None

    # Upload do arquivo CSV
    uploaded_file = st.file_uploader("Carregue seu arquivo CSV", type=["csv"])

    df = None
    if uploaded_file is not None:
        df = load_data_from_csv(uploaded_file)
        
        if df is not None and not df.empty:
            st.success("Dados do arquivo carregados com sucesso!")
            st.write("Prévia dos dados:")
            st.dataframe(df.head())
            
            st.markdown("---")
            
            # Opções de filtro
            st.header("Configurações de Filtro e Soma")
            
            columns = df.columns.tolist()
            
            col_to_filter = st.selectbox("Selecione a coluna para filtrar:", options=columns)
            filter_value = st.text_input(f"Informe o valor para filtrar na coluna '{col_to_filter}':")
            
            col_to_sum = st.selectbox("Selecione a coluna para somar:", options=columns)
            
            st.markdown("---")
            
            # Aplicar filtro e soma
            if st.button("Aplicar Filtro e Somar"):
                if not filter_value:
                    st.warning("Por favor, informe um valor para filtrar.")
                else:
                    try:
                        # Tentar converter a coluna de soma para numérico, ignorando erros
                        df[col_to_sum] = pd.to_numeric(df[col_to_sum], errors='coerce')
                        
                        # Filtrar o DataFrame
                        # Converte ambas as colunas para string para comparação consistente e case-insensitive
                        filtered_df = df[df[col_to_filter].astype(str).str.contains(filter_value, case=False, na=False)]
                        
                        if not filtered_df.empty:
                            st.subheader("Resultado do Filtro:")
                            st.dataframe(filtered_df)
                            
                            # Somar a coluna selecionada
                            total_sum = filtered_df[col_to_sum].sum()
                            st.success(f"Soma dos valores na coluna '{col_to_sum}' após o filtro: **{total_sum:,.2f}**")
                            
                            st.markdown("---")
                            
                            # Opções de Download
                            st.subheader("Opções de Download")
                            
                            # Download CSV
                            csv_buffer = io.StringIO()
                            filtered_df.to_csv(csv_buffer, index=False)
                            st.download_button(
                                label="Download CSV",
                                data=csv_buffer.getvalue(),
                                file_name="dados_filtrados.csv",
                                mime="text/csv"
                            )
                            
                            # Download XLSX
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                filtered_df.to_excel(writer, index=False, sheet_name='Dados Filtrados')
                            excel_buffer.seek(0)
                            st.download_button(
                                label="Download XLSX",
                                data=excel_buffer.getvalue(),
                                file_name="dados_filtrados.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.warning("Nenhum registro encontrado com o filtro aplicado.")
                    except KeyError:
                        st.error(f"Coluna '{col_to_filter}' ou '{col_to_sum}' não encontrada. Verifique se os nomes das colunas estão corretos.")
                    except Exception as e:
                        st.error(f"Ocorreu um erro ao aplicar o filtro ou soma: {e}")
        elif df is not None and df.empty:
            st.warning("O arquivo carregado está vazio ou não contém dados válidos.")
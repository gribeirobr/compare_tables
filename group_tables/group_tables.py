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
    st.set_page_config(page_title="Filtro Avançado de Planilhas", page_icon="📊", layout="wide")
    st.title("📊 Filtro Avançado de Planilha Google Sheets")

    # Inicialização de variáveis de sessão
    if 'filtros' not in st.session_state:
        st.session_state.filtros = []
    if 'df_original' not in st.session_state:
        st.session_state.df_original = None

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
            # Converter colunas numéricas que podem estar como strings
            for col in df.columns:
                if df[col].dtype == object:
                    try:
                        # Tenta substituir vírgula por ponto e converter para float
                        df[col] = pd.to_numeric(df[col].str.replace(',', '.'), errors='ignore')
                    except:
                        # Se falhar, mantém como string
                        pass
            st.session_state.df_original = df
            return df
        except Exception as e:
            st.error(f"Erro ao carregar dados: {str(e)}")
            return None

    # Adicionar novo filtro
    def adicionar_filtro():
        st.session_state.filtros.append({"coluna": "", "valor": ""})

    # Remover filtro
    def remover_filtro(index):
        st.session_state.filtros.pop(index)

    # Processar dados e gerar resultados
    def processar_dados(df, filtros, coluna_soma):
        try:
            # Aplicar todos os filtros
            for filtro in filtros:
                col = filtro["coluna"]
                val = filtro["valor"]
                if col and val and col in df.columns:
                    # Converter valor para string para comparação segura
                    df = df[df[col].astype(str) == str(val)]
            
            # Verificar se há dados após filtragem
            if df.empty:
                st.warning("⚠️ Nenhum dado encontrado após aplicação dos filtros!")
                return df, None
            
            # Calcular soma se possível
            soma = None
            if coluna_soma and coluna_soma in df.columns:
                try:
                    # Tentar converter para numérico se necessário
                    if not pd.api.types.is_numeric_dtype(df[coluna_soma]):
                        df[coluna_soma] = pd.to_numeric(df[coluna_soma], errors='coerce')
                    
                    soma = df[coluna_soma].sum()
                except Exception as e:
                    st.warning(f"⚠️ Não foi possível somar a coluna '{coluna_soma}': {str(e)}")
            
            return df, soma
        except KeyError as ke:
            st.error(f"Erro: Coluna não encontrada - {str(ke)}")
            return None, None
        except Exception as e:
            st.error(f"Erro no processamento: {str(e)}")
            return None, None

    # Gerar arquivo para download
    def gerar_download(df, formato):
        if df is None or df.empty:
            return None
            
        if formato == 'CSV':
            return df.to_csv(index=False).encode('utf-8-sig')
        elif formato == 'XLSX':
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Dados Filtrados')
            return output.getvalue()

    # Interface principal
    if st.button("Carregar Planilha") or st.session_state.df_original is not None:
        if not spreadsheet_id or not gid:
            st.warning("⚠️ Por favor, informe o ID da planilha e o GID da aba")
        else:
            df = carregar_dados(spreadsheet_id, gid)
            
            if df is not None:
                st.success(f"✅ Dados carregados com sucesso! ({df.shape[0]} linhas, {df.shape[1]} colunas)")
                
                # Botão para adicionar novos filtros
                st.button("➕ Adicionar Filtro", on_click=adicionar_filtro)
                
                # Interface de filtros
                with st.container():
                    st.subheader("Filtros")
                    
                    # Exibir filtros existentes
                    for i, filtro in enumerate(st.session_state.filtros):
                        col1, col2, col3 = st.columns([3, 3, 1])
                        with col1:
                            # Selecionar coluna para filtro
                            coluna = st.selectbox(
                                f"Coluna #{i+1}", 
                                df.columns, 
                                key=f"col_{i}",
                                index=df.columns.get_loc(filtro["coluna"]) if filtro["coluna"] in df.columns else 0
                            )
                            st.session_state.filtros[i]["coluna"] = coluna
                        
                        with col2:
                            # Selecionar valor para filtro - CORREÇÃO DO ERRO
                            if coluna and coluna in df.columns:
                                valores = df[coluna].astype(str).unique()
                                
                                # Encontrar índice do valor atual
                                try:
                                    index_valor = np.where(valores == str(filtro["valor"]))[0]
                                    index_valor = index_valor[0] if len(index_valor) > 0 else 0
                                except:
                                    index_valor = 0
                                
                                valor = st.selectbox(
                                    f"Valor #{i+1}", 
                                    valores, 
                                    key=f"val_{i}",
                                    index=index_valor
                                )
                                st.session_state.filtros[i]["valor"] = valor
                        
                        with col3:
                            # Botão para remover filtro
                            st.write(" ")
                            st.button("❌", key=f"del_{i}", on_click=remover_filtro, args=(i,))

                # Seleção da coluna para soma (todas as colunas disponíveis)
                st.subheader("Soma")
                coluna_soma = st.selectbox("Coluna para somar:", [""] + list(df.columns))
                
                if coluna_soma and coluna_soma != "":
                    # Verificar se a coluna é numérica
                    if not pd.api.types.is_numeric_dtype(df[coluna_soma]):
                        # Tentar converter automaticamente
                        try:
                            df[coluna_soma] = pd.to_numeric(df[coluna_soma], errors='coerce')
                            st.info(f"✅ Coluna '{coluna_soma}' convertida para numérica")
                        except:
                            st.warning(f"⚠️ A coluna '{coluna_soma}' não é numérica e não pode ser convertida")
                
                # Botão para processar
                if st.button("Aplicar Filtros e Calcular Soma"):
                    df_filtrado, soma = processar_dados(
                        st.session_state.df_original.copy(),
                        st.session_state.filtros,
                        coluna_soma if coluna_soma != "" else None
                    )
                    
                    if df_filtrado is not None and not df_filtrado.empty:
                        st.subheader("Resultados")
                        
                        # Mostrar resultado da soma se aplicável
                        if soma is not None:
                            st.metric(f"Soma de '{coluna_soma}'", f"{soma:,.2f}")
                        
                        # Mostrar tabela filtrada
                        st.dataframe(df_filtrado)
                        
                        # Download dos dados
                        st.subheader("Download dos Dados Filtrados")
                        
                        # Criar abas para diferentes formatos
                        tab1, tab2 = st.tabs(["📥 CSV", "📊 Excel"])
                        
                        with tab1:
                            csv = gerar_download(df_filtrado, 'CSV')
                            if csv:
                                st.download_button(
                                    label="Baixar como CSV",
                                    data=csv,
                                    file_name='dados_filtrados.csv',
                                    mime='text/csv'
                                )
                        
                        with tab2:
                            excel = gerar_download(df_filtrado, 'XLSX')
                            if excel:
                                st.download_button(
                                    label="Baixar como Excel",
                                    data=excel,
                                    file_name='dados_filtrados.xlsx',
                                    mime='application/vnd.ms-excel'
                                )
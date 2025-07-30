import streamlit as st
import pandas as pd
import datacompy
import re
import io
import numpy as np
from base64 import b64encode
from io import BytesIO

# --- Configuração Inicial ---
st.set_page_config(
    page_title="Suite de Ferramentas de Planilhas",
    layout="wide"
)

# --- Sistema de Autenticação Unificado ---
USUARIOS_CADASTRADOS = {
    "guilherme": "senha_guilherme",
    "GABI_REZENDE": "gabiR#123",
    "AMANDA_MASSON": "amandaM#123",
    "ERICA_NAS": "ericaN#123",
    "admin": "admin123"
}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None
if "pagina_atual" not in st.session_state:
    st.session_state.pagina_atual = "login"

# --- Funções Utilitárias Comuns ---
def limpar_texto(texto):
    if isinstance(texto, str):
        texto = texto.replace('\u00A0', ' ')
        texto = re.sub(r"[^\x20-\x7EÀ-ÿ]", "", texto)
        return texto.strip()
    return texto

def gerar_download(df, filename, formato='excel'):
    """
    Gera links de download para CSV e Excel
    Retorna HTML com ambos os links formatados
    """
    if df is None or df.empty:
        return ""
    
    # Download para Excel
    if formato == 'excel' or formato == 'ambos':
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel) as writer:
            df.to_excel(writer, index=False)
        b64_excel = b64encode(output_excel.getvalue()).decode()
        link_excel = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="{filename}.xlsx">📊 Excel</a>'
    
    # Download para CSV
    if formato == 'csv' or formato == 'ambos':
        csv = df.to_csv(index=False).encode('utf-8-sig')
        b64_csv = b64encode(csv).decode()
        link_csv = f'<a href="data:file/csv;base64,{b64_csv}" download="{filename}.csv">📥 CSV</a>'
    
    # Retorna os dois links formatados
    if formato == 'ambos':
        return f'<div style="display:flex;gap:10px;margin-top:10px;">{link_csv} {link_excel}</div>'
    elif formato == 'excel':
        return link_excel
    elif formato == 'csv':
        return link_csv

def carregar_planilha(sheet_id, gid):
    try:
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        df = pd.read_csv(url, encoding='utf-8-sig')
        
        # Converter colunas numéricas que podem estar como strings
        for col in df.columns:
            if df[col].dtype == object:
                try:
                    df[col] = pd.to_numeric(df[col].str.replace(',', '.'), errors='ignore')
                except:
                    pass
        return df, None
    except Exception as e:
        return None, str(e)

# --- Módulo 1: Comparador de Tabelas ---
def modulo_comparador():
    st.title("Comparador de Tabelas do Google Sheets")
    st.write("Compare duas abas de uma planilha Google Sheets e identifique diferenças")
    
    if st.button("← Voltar ao Menu Principal"):
        st.session_state.pagina_atual = "menu"
        st.rerun()

    spreadsheet_id = st.text_input("ID da Planilha", key="comp_id")
    gid_original = st.text_input("GID da Aba da Tabela A", key="comp_gid_a")
    gid_atualizado = st.text_input("GID da Aba da Tabela B", key="comp_gid_b")

    def dados_preenchidos():
        return all([spreadsheet_id, gid_original, gid_atualizado])

    if dados_preenchidos():
        try:
            df_original = carregar_planilha(spreadsheet_id, gid_original)[0]
            df_atualizado = carregar_planilha(spreadsheet_id, gid_atualizado)[0]

            if df_original is not None and df_atualizado is not None:
                st.success("Abas carregadas com sucesso!")

                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Colunas da Tabela A**")
                    colunas_original = st.multiselect("Selecione colunas para comparar - Tabela A", 
                                                     df_original.columns,
                                                     key="ms_orig")

                with col2:
                    st.markdown("**Colunas da Tabela B**")
                    colunas_atualizado = st.multiselect("Selecione colunas para comparar - Tabela B", 
                                                      df_atualizado.columns,
                                                      key="ms_atual")

                if st.button("Comparar Tabelas", type="primary"):
                    if len(colunas_original) != len(colunas_atualizado):
                        st.warning("Selecione o mesmo número de colunas em ambas as tabelas!")
                    else:
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

                        col_res1, col_res2, col_res3 = st.columns(3)
                        
                        with col_res1:
                            st.subheader("Registros encontrados apenas na Tabela A")
                            st.dataframe(df1_unq)
                            st.markdown(gerar_download(df1_unq, "registros_apenas_tabela_a", 'ambos'), unsafe_allow_html=True)
                        
                        with col_res2:
                            st.subheader("Registros encontrados apenas na Tabela B")
                            st.dataframe(df2_unq)
                            st.markdown(gerar_download(df2_unq, "registros_apenas_tabela_b", 'ambos'), unsafe_allow_html=True)
                        
                        with col_res3:
                            st.subheader("Registros encontrados em ambas tabelas")
                            st.dataframe(df_intersecao)
                            st.markdown(gerar_download(df_intersecao, "registros_encontrados_em_ambas_tabelas", 'ambos'), unsafe_allow_html=True)

                        st.divider()
                        st.metric("Total de Linhas Iguais", df_intersecao.shape[0])
                        st.metric("Total de Linhas Diferentes", df1_unq.shape[0] + df2_unq.shape[0])
            else:
                st.error("Erro ao carregar as abas. Verifique os IDs e permissões.")
        except Exception as e:
            st.error(f"Erro na comparação: {str(e)}")

# --- Módulo 2: Filtro Avançado ---
def modulo_filtro():
    st.title("Filtro Avançado de Planilhas")
    st.write("Filtre dados de uma planilha Google Sheets com múltiplos critérios e gere somas")
    
    if st.button("← Voltar ao Menu Principal"):
        st.session_state.pagina_atual = "menu"
        st.rerun()
    
    # Inicialização de variáveis de sessão
    if 'filtros' not in st.session_state:
        st.session_state.filtros = []
    if 'df_original' not in st.session_state:
        st.session_state.df_original = None

    # Entrada de dados do usuário
    with st.expander("Informações da Planilha", expanded=True):
        spreadsheet_id = st.text_input("ID da Planilha", key="filtro_id")
        gid = st.text_input("GID da Aba", key="filtro_gid")
        st.caption("Exemplo de URL: `https://docs.google.com/spreadsheets/d/[ID_AQUI]/edit#gid=[GID_AQUI]`")

    # Botões de controle
    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("Carregar Planilha", type="primary"):
            if spreadsheet_id and gid:
                df, error = carregar_planilha(spreadsheet_id, gid)
                if df is not None:
                    st.session_state.df_original = df
                    st.success(f"Dados carregados! ({df.shape[0]} linhas)")
                else:
                    st.error(f"Erro: {error}")
            else:
                st.warning("Preencha ID e GID primeiro")
    
    with col2:
        if st.button("➕ Adicionar Novo Filtro"):
            st.session_state.filtros.append({"coluna": "", "valor": ""})

    # Exibir filtros se houver dados
    if st.session_state.df_original is not None:
        df = st.session_state.df_original
        
        st.subheader("Filtros Aplicados")
        if not st.session_state.filtros:
            st.info("Nenhum filtro aplicado. Clique em 'Adicionar Novo Filtro' para começar.")
        
        # Interface de filtros
        for i, filtro in enumerate(st.session_state.filtros):
            col_f1, col_f2, col_f3 = st.columns([4, 4, 1])
            with col_f1:
                coluna = st.selectbox(
                    f"Coluna #{i+1}", 
                    df.columns, 
                    key=f"fcol_{i}",
                    index=df.columns.get_loc(filtro["coluna"]) if filtro["coluna"] in df.columns else 0
                )
                st.session_state.filtros[i]["coluna"] = coluna
            
            with col_f2:
                if coluna and coluna in df.columns:
                    valores = df[coluna].astype(str).unique()
                    try:
                        index_valor = np.where(valores == str(filtro["valor"]))[0]
                        index_valor = int(index_valor[0]) if index_valor.size > 0 else 0
                    except:
                        index_valor = 0
                    
                    valor = st.selectbox(
                        f"Valor #{i+1}", 
                        valores, 
                        key=f"fval_{i}",
                        index=index_valor
                    )
                    st.session_state.filtros[i]["valor"] = valor
            
            with col_f3:
                st.write(" ")
                if st.button("❌", key=f"fdel_{i}"):
                    st.session_state.filtros.pop(i)
                    st.rerun()
        
        # Seleção da coluna para soma
        st.subheader("Cálculos")
        coluna_soma = st.selectbox("Coluna para somar:", [""] + list(df.columns))
        
        # Botão para processar
        if st.button("Aplicar Filtros e Calcular", type="primary") and st.session_state.filtros:
            try:
                df_filtrado = df.copy()
                for filtro in st.session_state.filtros:
                    col = filtro["coluna"]
                    val = filtro["valor"]
                    if col and val and col in df_filtrado.columns:
                        df_filtrado = df_filtrado[df_filtrado[col].astype(str) == str(val)]
                
                if df_filtrado.empty:
                    st.warning("Nenhum dado encontrado após filtragem!")
                else:
                    st.subheader("Resultados Filtrados")
                    st.dataframe(df_filtrado)
                    
                    # Cálculo da soma se aplicável
                    if coluna_soma and coluna_soma in df_filtrado.columns:
                        try:
                            if not pd.api.types.is_numeric_dtype(df_filtrado[coluna_soma]):
                                df_filtrado[coluna_soma] = pd.to_numeric(df_filtrado[coluna_soma], errors='coerce')
                            
                            soma = df_filtrado[coluna_soma].sum()
                            st.metric(f"Total da coluna '{coluna_soma}'", f"{soma:,.2f}")
                        except:
                            st.warning(f"Não foi possível somar a coluna '{coluna_soma}'")
                    
                    # Download
                    st.subheader("Exportar Resultados")
                    st.markdown(gerar_download(df_filtrado, "dados_filtrados", 'ambos'), unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Erro no processamento: {str(e)}")

# --- Módulo 3: Renomeador de Colunas ---
def modulo_renomeador():
    st.title("Renomeador de Colunas")
    st.write("Selecione e renomeie colunas de uma planilha Google Sheets")
    
    if st.button("← Voltar ao Menu Principal"):
        st.session_state.pagina_atual = "menu"
        st.rerun()
    
    if 'df_original' not in st.session_state:
        st.session_state.df_original = None
    if 'df_final' not in st.session_state:
        st.session_state.df_final = None
    if 'mapa_renomeacao' not in st.session_state:
        st.session_state.mapa_renomeacao = {}

    st.header("Informações da Planilha")
    col1, col2 = st.columns(2)
    with col1:
        google_sheet_id = st.text_input("ID da Planilha", key="ren_id")
    with col2:
        google_sheet_gid = st.text_input("GID da Aba", key="ren_gid")

    if st.button("Carregar Planilha", type="primary"):
        if google_sheet_id and google_sheet_gid:
            with st.spinner("Carregando dados..."):
                df, erro = carregar_planilha(google_sheet_id, google_sheet_gid)
                if df is not None:
                    st.session_state.df_original = df
                    st.success("Dados carregados com sucesso!")
                else:
                    st.error(f"Erro: {erro}")
        else:
            st.warning("Preencha ID e GID primeiro")

    if st.session_state.df_original is not None:
        df_original = st.session_state.df_original
        
        st.header("Seleção de Colunas")
        colunas_selecionadas = st.multiselect(
            "Selecione as colunas que deseja manter:",
            options=df_original.columns.tolist(),
            default=list(st.session_state.mapa_renomeacao.keys())
        )

        if colunas_selecionadas:
            st.header("Renomear Colunas")
            st.write("Defina os novos nomes para cada coluna selecionada:")
            
            # Atualizar mapa de renomeação
            for col in colunas_selecionadas:
                if col not in st.session_state.mapa_renomeacao:
                    st.session_state.mapa_renomeacao[col] = col

            # Interface de renomeação
            for col in colunas_selecionadas:
                novo_nome = st.text_input(
                    f"Nome para '{col}'", 
                    value=st.session_state.mapa_renomeacao[col],
                    key=f"rn_{col}"
                )
                st.session_state.mapa_renomeacao[col] = novo_nome

            if st.button("Gerar Planilha Renomeada", type="primary"):
                df_filtrado = df_original[colunas_selecionadas]
                df_renomeado = df_filtrado.rename(columns=st.session_state.mapa_renomeacao)
                st.session_state.df_final = df_renomeado
                st.success("Planilha renomeada com sucesso!")

        if st.session_state.df_final is not None:
            st.header("Resultado Final")
            st.dataframe(st.session_state.df_final.head())
            
            st.subheader("Baixar Planilha Renomeada")
            st.markdown(gerar_download(st.session_state.df_final, "planilha_renomeada", 'ambos'), unsafe_allow_html=True)

# --- Página de Login ---
def pagina_login():
    st.title("Bem-vindo à Suite de Planilhas")
    st.subheader("Autenticação")
    
    usuario = st.text_input("Usuário")
    senha = st.text_input("Senha", type="password")
    
    if st.button("Entrar", type="primary"):
        if usuario in USUARIOS_CADASTRADOS and senha == USUARIOS_CADASTRADOS[usuario]:
            st.session_state.autenticado = True
            st.session_state.usuario_logado = usuario
            st.session_state.pagina_atual = "menu"
            st.rerun()
        else:
            st.error("Credenciais inválidas. Tente novamente.")

# --- Menu Principal ---
def menu_principal():
    st.title(f"Suite de Ferramentas de Planilhas")
    st.markdown(f"**Usuário:** {st.session_state.usuario_logado}")
    
    st.write("Selecione a ferramenta que deseja utilizar:")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Comparador")
        st.write("Compare duas tabelas e identifique diferenças")
        if st.button("Acessar Comparador", key="btn1"):
            st.session_state.pagina_atual = "comparador"
            st.rerun()
    
    with col2:
        st.subheader("Filtro Avançado")
        st.write("Filtre dados com múltiplos critérios")
        if st.button("Acessar Filtro", key="btn2"):
            st.session_state.pagina_atual = "filtro"
            st.rerun()
    
    with col3:
        st.subheader("Renomeador")
        st.write("Selecione e renomeie colunas")
        if st.button("Acessar Renomeador", key="btn3"):
            st.session_state.pagina_atual = "renomeador"
            st.rerun()
    
    st.divider()
    if st.button("Sair do Sistema", type="secondary"):
        st.session_state.autenticado = False
        st.session_state.usuario_logado = None
        st.session_state.pagina_atual = "login"
        st.rerun()

# --- Navegação Principal ---
if st.session_state.pagina_atual == "login":
    pagina_login()
elif st.session_state.pagina_atual == "menu":
    menu_principal()
elif st.session_state.pagina_atual == "comparador":
    modulo_comparador()
elif st.session_state.pagina_atual == "filtro":
    modulo_filtro()
elif st.session_state.pagina_atual == "renomeador":
    modulo_renomeador()
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
    """Carrega dados de uma planilha Google Sheets."""
    if not sheet_id or not gid:
        return None, "O ID da Planilha e o GID da Aba não podem estar vazios."
    try:
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        df = pd.read_csv(url, encoding='utf-8-sig')
        
        for col in df.columns:
            if df[col].dtype == object:
                try:
                    df[col] = pd.to_numeric(df[col].str.replace(',', '.'), errors='coerce').fillna(df[col])
                except:
                    pass
        return df, None
    except Exception as e:
        return None, f"Erro ao carregar do Google Sheets: Verifique o ID, GID e as permissões de compartilhamento da planilha. (Detalhe: {e})"

def carregar_arquivo_local(uploaded_file):
    """Lê um arquivo CSV ou XLSX enviado pelo usuário e retorna um DataFrame."""
    if uploaded_file is None:
        return None, "Nenhum arquivo enviado."
    
    nome_arquivo = uploaded_file.name
    try:
        if nome_arquivo.endswith('.csv'):
            uploaded_file.seek(0)
            # **MUDANÇA PRINCIPAL:** Usar o 'sniffer' do pandas para detectar o separador
            # O `sep=None` com `engine='python'` ativa a detecção automática.
            df = pd.read_csv(uploaded_file, sep=None, engine='python', encoding='utf-8-sig')
            
            # Checagem final para garantir que a leitura foi bem sucedida
            if df.shape[1] <= 1:
                st.warning(
                    f"A leitura do arquivo '{nome_arquivo}' resultou em apenas uma coluna. "
                    "O pandas não conseguiu identificar o separador automaticamente. "
                    "Verifique se o arquivo está bem formatado, usando vírgulas (,) ou ponto e vírgula (;) como delimitador."
                )

        elif nome_arquivo.endswith(('.xlsx', '.xls')):
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file)
        else:
            return None, "Formato de arquivo inválido. Por favor, envie um arquivo .csv ou .xlsx."
        
        # Limpeza e conversão de tipos
        # Remove espaços em branco dos nomes das colunas, um erro comum
        df.rename(columns={col: col.strip() for col in df.columns}, inplace=True)

        for col in df.columns:
            if df[col].dtype == object:
                try:
                    # Tenta converter para numérico, tratando vírgula como decimal
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='ignore')
                except:
                    pass # Ignora erros se a conversão não for possível
        return df, None
    except Exception as e:
        return None, f"Erro ao processar o arquivo: {e}"

# --- Módulo 1: Comparador de Tabelas ---
def modulo_comparador():
    st.title("Comparador de Tabelas")
    st.write("Compare duas tabelas (do Google Sheets ou de arquivos locais) e identifique as diferenças.")
    
    if st.button("← Voltar ao Menu Principal"):
        # Limpar o estado do módulo ao sair para evitar dataframes "fantasmas"
        for key in ['df_original_comp', 'df_atualizado_comp']:
            if key in st.session_state:
                del st.session_state[key]
        st.session_state.pagina_atual = "menu"
        st.rerun()

    # Inicialização dos dataframes
    df_original = None
    df_atualizado = None

    col1, col2 = st.columns(2)

    with col1:
        st.header("Tabela A (Original)")
        origem_a = st.radio("Origem da Tabela A", ["Google Sheets", "Upload de Arquivo"], key="origem_a")
        if origem_a == "Google Sheets":
            spreadsheet_id_a = st.text_input("ID da Planilha (Tabela A)", key="comp_id_a")
            gid_original = st.text_input("GID da Aba (Tabela A)", key="comp_gid_a")
            if st.button("Carregar Tabela A", key="load_a"):
                with st.spinner("Carregando Tabela A..."):
                    df, erro = carregar_planilha(spreadsheet_id_a, gid_original)
                    if erro: st.error(f"Erro ao carregar Tabela A: {erro}")
                    else: 
                        st.session_state.df_original_comp = df
                        st.success(f"Tabela A carregada ({df.shape[0]} linhas).")
        else:
            upload_a = st.file_uploader("Carregar arquivo da Tabela A (CSV, XLSX)", type=['csv', 'xlsx'], key="upload_a")
            if upload_a:
                with st.spinner("Processando Arquivo A..."):
                    df, erro = carregar_arquivo_local(upload_a)
                    if erro: st.error(f"Erro no arquivo A: {erro}")
                    else: 
                        st.session_state.df_original_comp = df
                        st.success(f"Arquivo da Tabela A carregado ({df.shape[0]} linhas).")

    with col2:
        st.header("Tabela B (Atualizada)")
        origem_b = st.radio("Origem da Tabela B", ["Google Sheets", "Upload de Arquivo"], key="origem_b")
        if origem_b == "Google Sheets":
            spreadsheet_id_b = st.text_input("ID da Planilha (Tabela B)", key="comp_id_b")
            gid_atualizado = st.text_input("GID da Aba (Tabela B)", key="comp_gid_b")
            if st.button("Carregar Tabela B", key="load_b"):
                 with st.spinner("Carregando Tabela B..."):
                    df, erro = carregar_planilha(spreadsheet_id_b, gid_atualizado)
                    if erro: st.error(f"Erro ao carregar Tabela B: {erro}")
                    else: 
                        st.session_state.df_atualizado_comp = df
                        st.success(f"Tabela B carregada ({df.shape[0]} linhas).")
        else:
            upload_b = st.file_uploader("Carregar arquivo da Tabela B (CSV, XLSX)", type=['csv', 'xlsx'], key="upload_b")
            if upload_b:
                with st.spinner("Processando Arquivo B..."):
                    df, erro = carregar_arquivo_local(upload_b)
                    if erro: st.error(f"Erro no arquivo B: {erro}")
                    else: 
                        st.session_state.df_atualizado_comp = df
                        st.success(f"Arquivo da Tabela B carregado ({df.shape[0]} linhas).")

    # Recupera os dataframes do estado da sessão
    if 'df_original_comp' in st.session_state:
        df_original = st.session_state.df_original_comp
    if 'df_atualizado_comp' in st.session_state:
        df_atualizado = st.session_state.df_atualizado_comp

    if df_original is not None and df_atualizado is not None:
        st.divider()
        st.header("Mapeamento de Colunas para Comparação")
        st.info("Selecione as colunas que servem como identificador único dos registros (chaves). Ex: ID do cliente, CPF, etc.")
        
        col_map1, col_map2 = st.columns(2)
        with col_map1:
            st.markdown("**Colunas-Chave da Tabela A**")
            colunas_original = st.multiselect("Selecione as colunas da Tabela A", df_original.columns, key="ms_orig")
        with col_map2:
            st.markdown("**Colunas-Chave da Tabela B**")
            colunas_atualizado = st.multiselect("Selecione as colunas da Tabela B (na mesma ordem da Tabela A)", df_atualizado.columns, key="ms_atual")

        if st.button("Comparar Tabelas", type="primary"):
            if not colunas_original or not colunas_atualizado:
                st.warning("Selecione pelo menos uma coluna em cada tabela para usar como chave de comparação.")
            elif len(colunas_original) != len(colunas_atualizado):
                st.warning("O número de colunas-chave deve ser o mesmo em ambas as tabelas!")
            else:
                try:
                    with st.spinner("Preparando dados e comparando..."):
                        # Criar uma cópia do df2 para não alterar o original no session_state
                        df_atualizado_temp = df_atualizado.copy()

                        # Criar o dicionário para renomear as colunas de B para que coincidam com A
                        mapa_renomeacao = {col_b: col_a for col_a, col_b in zip(colunas_original, colunas_atualizado)}
                        
                        # Renomear as colunas na cópia do df2
                        df_atualizado_temp.rename(columns=mapa_renomeacao, inplace=True)

                        # Agora, a comparação usa `colunas_original` como a chave para AMBOS os DataFrames
                        comparacao = datacompy.Compare(
                            df1=df_original,
                            df2=df_atualizado_temp, # Usar o DF com colunas renomeadas
                            join_columns=colunas_original, # A lista de chaves agora é a mesma para ambos
                            df1_name='Original',
                            df2_name='Atualizado',
                            abs_tol=0.0001, # Tolerância para números de ponto flutuante
                            rel_tol=0
                        )
                        comparacao.matches(ignore_extra_columns=True)

                        df1_unq = comparacao.df1_unq_rows
                        df2_unq = comparacao.df2_unq_rows
                        df_intersecao = comparacao.intersect_rows
                        
                        # Limpa textos em todos os dataframes de resultado
                        for df in [df1_unq, df2_unq, df_intersecao]:
                            for col in df.select_dtypes(include='object').columns:
                                df[col] = df[col].apply(limpar_texto)
                        
                        st.header("Resultados da Comparação")
                        col_res1, col_res2, col_res3 = st.columns(3)
                        
                        with col_res1:
                            st.subheader(f"Registros únicos na Tabela A ({df1_unq.shape[0]})")
                            st.dataframe(df1_unq)
                            st.markdown(gerar_download(df1_unq, "registros_apenas_tabela_a", 'ambos'), unsafe_allow_html=True)
                        
                        with col_res2:
                            st.subheader(f"Registros únicos na Tabela B ({df2_unq.shape[0]})")
                            st.dataframe(df2_unq)
                            st.markdown(gerar_download(df2_unq, "registros_apenas_tabela_b", 'ambos'), unsafe_allow_html=True)
                        
                        with col_res3:
                            st.subheader(f"Registros iguais em ambas ({df_intersecao.shape[0]})")
                            st.dataframe(df_intersecao)
                            st.markdown(gerar_download(df_intersecao, "registros_em_ambas_tabelas", 'ambos'), unsafe_allow_html=True)

                        st.divider()
                        with st.expander("Ver Relatório Detalhado da Comparação"):
                            st.text(comparacao.report())

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
    if 'df_original_filtro' not in st.session_state:
        st.session_state.df_original_filtro = None

    with st.expander("Passo 1: Carregar Dados", expanded=True):
        origem = st.radio("Origem da Planilha", ["Google Sheets", "Upload de Arquivo"], key="filtro_origem")
        
        if origem == "Google Sheets":
            spreadsheet_id = st.text_input("ID da Planilha", key="filtro_id")
            gid = st.text_input("GID da Aba", key="filtro_gid")
            if st.button("Carregar do Google Sheets", type="primary"):
                if spreadsheet_id and gid:
                    with st.spinner("Carregando..."):
                        df, error = carregar_planilha(spreadsheet_id, gid)
                        if df is not None:
                            st.session_state.df_original_filtro = df
                            st.success(f"Dados carregados! ({df.shape[0]} linhas)")
                            st.rerun()
                        else:
                            st.error(f"Erro: {error}")
                else:
                    st.warning("Preencha o ID da Planilha e o GID da Aba.")
        else:
            uploaded_file = st.file_uploader("Selecione um arquivo (CSV ou XLSX)", type=['csv', 'xlsx'], key="filtro_upload")
            if uploaded_file:
                with st.spinner("Processando arquivo..."):
                    df, error = carregar_arquivo_local(uploaded_file)
                    if df is not None:
                        st.session_state.df_original_filtro = df
                        st.success(f"Arquivo carregado! ({df.shape[0]} linhas)")
                        st.rerun()
                    else:
                        st.error(f"Erro: {error}")
    
    if st.session_state.df_original_filtro is not None:
        df = st.session_state.df_original_filtro
        st.dataframe(df.head())
        
        st.header("Passo 2: Configurar Filtros")
        if st.button("➕ Adicionar Novo Filtro"):
            st.session_state.filtros.append({"coluna": df.columns[0], "valor": ""})
            st.rerun()

        if not st.session_state.filtros:
            st.info("Nenhum filtro aplicado. Clique em 'Adicionar Novo Filtro' para começar.")
        
        # Interface de filtros
        for i, filtro in enumerate(st.session_state.filtros):
            col_f1, col_f2, col_f3 = st.columns([4, 4, 1])
            with col_f1:
                coluna = st.selectbox(f"Coluna #{i+1}", df.columns, key=f"fcol_{i}", index=list(df.columns).index(filtro["coluna"]))
                st.session_state.filtros[i]["coluna"] = coluna
            
            with col_f2:
                valores_unicos = sorted(df[coluna].dropna().astype(str).unique())
                valor_selecionado = st.selectbox(f"Valor #{i+1}", valores_unicos, key=f"fval_{i}")
                st.session_state.filtros[i]["valor"] = valor_selecionado
            
            with col_f3:
                st.write("")
                st.write("")
                if st.button("❌", key=f"fdel_{i}", help="Remover filtro"):
                    st.session_state.filtros.pop(i)
                    st.rerun()
        
        st.header("Passo 3: Aplicar Filtros e Calcular")
        coluna_soma = st.selectbox("Selecione uma coluna para somar (opcional):", [""] + [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])])
        
        if st.button("Aplicar Filtros e Calcular", type="primary"):
            if not st.session_state.filtros:
                st.warning("Adicione pelo menos um filtro para continuar.")
            else:
                try:
                    df_filtrado = df.copy()
                    for filtro in st.session_state.filtros:
                        col, val = filtro["coluna"], filtro["valor"]
                        if col and val:
                            df_filtrado = df_filtrado[df_filtrado[col].astype(str) == str(val)]
                    
                    st.session_state.df_filtrado_final = df_filtrado
                except Exception as e:
                    st.error(f"Erro no processamento: {str(e)}")

        if 'df_filtrado_final' in st.session_state:
            df_filtrado = st.session_state.df_filtrado_final
            st.header("Resultados")
            if df_filtrado.empty:
                st.warning("Nenhum dado encontrado com os filtros aplicados.")
            else:
                st.dataframe(df_filtrado)
                st.markdown(f"**Total de registros encontrados:** {df_filtrado.shape[0]}")
                
                if coluna_soma:
                    try:
                        soma = df_filtrado[coluna_soma].sum()
                        st.metric(f"Soma da coluna '{coluna_soma}'", f"{soma:,.2f}")
                    except Exception as e:
                        st.warning(f"Não foi possível somar a coluna '{coluna_soma}': {e}")
                
                st.subheader("Exportar Resultados")
                st.markdown(gerar_download(df_filtrado, "dados_filtrados", 'ambos'), unsafe_allow_html=True)


# --- Módulo 3: Renomeador de Colunas ---
def modulo_renomeador():
    st.title("Renomeador de Colunas")
    st.write("Selecione e renomeie colunas de uma planilha Google Sheets")
    
    if st.button("← Voltar ao Menu Principal"):
        st.session_state.pagina_atual = "menu"
        st.rerun()
    
    # Resetar estados se a página for recarregada
    if 'df_original_ren' not in st.session_state:
        st.session_state.df_original_ren = None
    if 'df_final_ren' not in st.session_state:
        st.session_state.df_final_ren = None
    if 'mapa_renomeacao' not in st.session_state:
        st.session_state.mapa_renomeacao = {}

    with st.expander("Passo 1: Carregar Dados", expanded=True):
        origem = st.radio("Origem da Planilha", ["Google Sheets", "Upload de Arquivo"], key="ren_origem")
        
        if origem == "Google Sheets":
            google_sheet_id = st.text_input("ID da Planilha", key="ren_id")
            google_sheet_gid = st.text_input("GID da Aba", key="ren_gid")
            if st.button("Carregar do Google Sheets", type="primary"):
                if google_sheet_id and google_sheet_gid:
                    with st.spinner("Carregando..."):
                        df, erro = carregar_planilha(google_sheet_id, google_sheet_gid)
                        if df is not None:
                            st.session_state.df_original_ren = df
                            st.session_state.mapa_renomeacao = {col: col for col in df.columns}
                            st.success("Dados carregados com sucesso!")
                            st.rerun()
                        else:
                            st.error(f"Erro: {erro}")
                else:
                    st.warning("Preencha o ID da Planilha e o GID da Aba.")
        else:
            uploaded_file = st.file_uploader("Selecione um arquivo (CSV ou XLSX)", type=['csv', 'xlsx'], key="ren_upload")
            if uploaded_file:
                with st.spinner("Processando arquivo..."):
                    df, erro = carregar_arquivo_local(uploaded_file)
                    if df is not None:
                        st.session_state.df_original_ren = df
                        st.session_state.mapa_renomeacao = {col: col for col in df.columns}
                        st.success("Arquivo carregado com sucesso!")
                        st.rerun()
                    else:
                        st.error(f"Erro: {erro}")

    if st.session_state.df_original_ren is not None:
        df_original = st.session_state.df_original_ren
        st.dataframe(df_original.head())
        
        st.header("Passo 2: Selecionar e Renomear Colunas")
        colunas_selecionadas = st.multiselect(
            "Selecione as colunas que deseja manter:",
            options=df_original.columns.tolist(),
            default=df_original.columns.tolist()
        )

        if colunas_selecionadas:
            st.write("Defina os novos nomes para cada coluna selecionada:")
            
            novos_nomes = {}
            for col in colunas_selecionadas:
                novo_nome = st.text_input(
                    f"'{col}' → Novo nome:", 
                    value=st.session_state.mapa_renomeacao.get(col, col),
                    key=f"rn_{col}"
                )
                novos_nomes[col] = novo_nome
            
            st.session_state.mapa_renomeacao = novos_nomes

            if st.button("Gerar Planilha Final", type="primary"):
                try:
                    df_filtrado = df_original[colunas_selecionadas]
                    df_renomeado = df_filtrado.rename(columns=st.session_state.mapa_renomeacao)
                    st.session_state.df_final_ren = df_renomeado
                    st.success("Planilha processada com sucesso!")
                except Exception as e:
                    st.error(f"Ocorreu um erro: {e}")

        if st.session_state.df_final_ren is not None:
            st.header("Passo 3: Resultado")
            st.dataframe(st.session_state.df_final_ren.head())
            
            st.subheader("Baixar Planilha Renomeada")
            st.markdown(gerar_download(st.session_state.df_final_ren, "planilha_renomeada", 'ambos'), unsafe_allow_html=True)

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
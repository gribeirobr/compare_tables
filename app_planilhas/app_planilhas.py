import streamlit as st
import pandas as pd
import datacompy
import re
import io
import numpy as np
from base64 import b64encode
from io import BytesIO
import extra_streamlit_components as stc
import datetime
import zipfile

cookie_manager = stc.CookieManager()

HELP_TEXT_SHEET_ID = "Você encontra o ID na URL da sua planilha. É a longa sequência de caracteres entre `/spreadsheets/d/` e `/edit`."
HELP_TEXT_GID = "O GID identifica a aba específica (página) da sua planilha. Você o encontra no final da URL, logo após `#gid=` (geralmente é um número como 0, 1, 2...)."

# --- Configuração Inicial ---
st.set_page_config(
    page_title="Ferramentas de Planilhas - GRB",
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
    
def gerar_download_zip(dfs_dict, formato_arquivo):
    """Gera um arquivo zip em memória contendo múltiplos DataFrames."""
    zip_buffer = io.BytesIO()
    ext = 'csv' if formato_arquivo == 'CSV' else 'xlsx'
    
    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
        for filename, df in dfs_dict.items():
            if formato_arquivo == 'CSV':
                # Salva como CSV
                csv_data = df.to_csv(index=False).encode('utf-8-sig')
                zip_file.writestr(f'{filename}.{ext}', csv_data)
            else:
                # Salva como XLSX
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Dados')
                excel_buffer.seek(0)
                zip_file.writestr(f'{filename}.{ext}', excel_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer.getvalue()
    
def reset_comparador_state():
    keys_to_delete = ['df_original_comp', 'df_atualizado_comp']
    for key in keys_to_delete:
        if key in st.session_state:
            del st.session_state[key]

def reset_filtro_state():
    keys_to_delete = ['filtros', 'df_original_filtro', 'df_filtrado_final']
    for key in keys_to_delete:
        if key in st.session_state:
            del st.session_state[key]

def reset_renomeador_state():
    keys_to_delete = ['df_original_ren', 'df_final_ren', 'mapa_renomeacao']
    for key in keys_to_delete:
        if key in st.session_state:
            del st.session_state[key]

def reset_unificador_state():
    for key in ['df_a_unif', 'df_b_unif', 'df_unificado_final']:
        if key in st.session_state: del st.session_state[key]

def reset_agrupador_state():
    for key in ['df_original_agrup', 'df_agrupado_final']:
        if key in st.session_state: del st.session_state[key]

def reset_limpador_state():
    for key in ['df_original_limp', 'df_limpo_final', 'limpeza_acoes']:
        if key in st.session_state: del st.session_state[key]

def reset_divisor_state():
    for key in ['df_original_div', 'dados_zip_div']:
        if key in st.session_state: del st.session_state[key]

# --- Módulo 1: Comparador de Tabelas ---
def modulo_comparador():
    st.title("Comparador de Tabelas")
    st.write("Compare duas tabelas (do Google Sheets ou de arquivos locais) e identifique as diferenças.")
    
    if st.button("← Voltar ao Menu Principal"):
        reset_comparador_state() # Limpa o estado deste módulo
        st.session_state.pagina_atual = "menu"
        st.rerun()

    st.warning("Sua planilha do Google Sheets deve estar em modo público. Acesse \"Compartilhar\" e ative a opção \"Qualquer pessoa com o link pode visualizar\".")

    # Inicialização dos dataframes
    df_original = None
    df_atualizado = None

    col1, col2 = st.columns(2)

    with col1:
        st.header("Tabela A")
        origem_a = st.radio("Origem da Tabela A", ["Google Sheets", "Upload de Arquivo"], key="origem_a")
        if origem_a == "Google Sheets":
            spreadsheet_id_a = st.text_input("ID da Planilha (Tabela A)", key="comp_id_a", help=HELP_TEXT_SHEET_ID)
            gid_original = st.text_input("GID da Aba (Tabela A)", key="comp_gid_a", help=HELP_TEXT_GID)
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
        st.header("Tabela B")
        origem_b = st.radio("Origem da Tabela B", ["Google Sheets", "Upload de Arquivo"], key="origem_b")
        if origem_b == "Google Sheets":
            spreadsheet_id_b = st.text_input("ID da Planilha (Tabela B)", key="comp_id_b", help=HELP_TEXT_SHEET_ID)
            gid_atualizado = st.text_input("GID da Aba (Tabela B)", key="comp_gid_b", help=HELP_TEXT_GID)
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
    st.write("Filtre dados de uma planilha (do Google Sheets ou de arquivos locais) com múltiplos critérios e gere somas")
    
    if st.button("← Voltar ao Menu Principal"):
        reset_filtro_state() # Limpa o estado deste módulo
        st.session_state.pagina_atual = "menu"
        st.rerun()

    st.warning("Sua planilha do Google Sheets deve estar em modo público. Acesse \"Compartilhar\" e ative a opção \"Qualquer pessoa com o link pode visualizar\".")
    
    # Inicialização de variáveis de sessão
    if 'filtros' not in st.session_state:
        st.session_state.filtros = []
    if 'df_original_filtro' not in st.session_state:
        st.session_state.df_original_filtro = None

    with st.expander("Passo 1: Carregar Dados", expanded=True):
        origem = st.radio("Origem da Planilha", ["Google Sheets", "Upload de Arquivo"], key="filtro_origem")
        
        if origem == "Google Sheets":
            spreadsheet_id = st.text_input("ID da Planilha", key="filtro_id", help=HELP_TEXT_SHEET_ID)
            gid = st.text_input("GID da Aba", key="filtro_gid", help=HELP_TEXT_GID)
            if st.button("Carregar do Google Sheets", type="primary"):
                if spreadsheet_id and gid:
                    with st.spinner("Carregando..."):
                        df, error = carregar_planilha(spreadsheet_id, gid)
                        if df is not None:
                            st.session_state.df_original_filtro = df
                            st.success(f"Dados carregados! ({df.shape[0]} linhas)")
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
    st.write("Selecione e renomeie colunas de uma planilha (do Google Sheets ou de arquivos locais)")
    
    if st.button("← Voltar ao Menu Principal"):
        reset_renomeador_state() # Limpa o estado deste módulo
        st.session_state.pagina_atual = "menu"
        st.rerun()
    
    # Resetar estados se a página for recarregada
    if 'df_original_ren' not in st.session_state:
        st.session_state.df_original_ren = None
    if 'df_final_ren' not in st.session_state:
        st.session_state.df_final_ren = None
    if 'mapa_renomeacao' not in st.session_state:
        st.session_state.mapa_renomeacao = {}
        
    st.warning("Sua planilha do Google Sheets deve estar em modo público. Acesse \"Compartilhar\" e ative a opção \"Qualquer pessoa com o link pode visualizar\".")

    with st.expander("Passo 1: Carregar Dados", expanded=True):
        origem = st.radio("Origem da Planilha", ["Google Sheets", "Upload de Arquivo"], key="ren_origem")
        
        if origem == "Google Sheets":
            google_sheet_id = st.text_input("ID da Planilha", key="ren_id", help=HELP_TEXT_SHEET_ID)
            google_sheet_gid = st.text_input("GID da Aba", key="ren_gid", help=HELP_TEXT_GID)
            if st.button("Carregar do Google Sheets", type="primary"):
                if google_sheet_id and google_sheet_gid:
                    with st.spinner("Carregando..."):
                        df, erro = carregar_planilha(google_sheet_id, google_sheet_gid)
                        if df is not None:
                            st.session_state.df_original_ren = df
                            st.session_state.mapa_renomeacao = {col: col for col in df.columns}
                            st.success("Dados carregados com sucesso!")
                            
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

def modulo_unificador():
    st.title("Unificador de Tabelas (PROCV/VLOOKUP Inteligente)")
    st.write("Cruze informações de duas tabelas com base em uma coluna em comum.")
    
    if st.button("← Voltar ao Menu Principal"):
        reset_unificador_state()
        st.session_state.pagina_atual = "menu"
        st.rerun()

    st.warning("Sua planilha do Google Sheets deve estar em modo público. Acesse \"Compartilhar\" e ative a opção \"Qualquer pessoa com o link pode visualizar\".")

    # Passo 1: Carregar as duas tabelas
    col1, col2 = st.columns(2)
    with col1:
        st.header("Tabela A (Esquerda)")
        origem_a = st.radio("Origem", ["Google Sheets", "Upload"], key="unif_origem_a")
        if origem_a == "Google Sheets":
            id_a = st.text_input("ID da Planilha A", key="unif_id_a", help=HELP_TEXT_SHEET_ID)
            gid_a = st.text_input("GID da Aba A", key="unif_gid_a", help=HELP_TEXT_GID)
            if st.button("Carregar Tabela A"):
                df, erro = carregar_planilha(id_a, gid_a)
                if erro: st.error(erro)
                else: 
                    st.session_state.df_a_unif = df
                    st.success(f"Tabela A carregada ({df.shape[0]} linhas).")
        else:
            up_a = st.file_uploader("Arquivo da Tabela A", type=['csv','xlsx'], key="unif_up_a")
            if up_a:
                df, erro = carregar_arquivo_local(up_a)
                if erro: st.error(erro)
                else: 
                    st.session_state.df_a_unif = df
                    st.success(f"Tabela A carregada ({df.shape[0]} linhas).")

    with col2:
        st.header("Tabela B (Direita)")
        origem_b = st.radio("Origem", ["Google Sheets", "Upload"], key="unif_origem_b")
        if origem_b == "Google Sheets":
            id_b = st.text_input("ID da Planilha B", key="unif_id_b", help=HELP_TEXT_SHEET_ID)
            gid_b = st.text_input("GID da Aba B", key="unif_gid_b", help=HELP_TEXT_GID)
            if st.button("Carregar Tabela B"):
                df, erro = carregar_planilha(id_b, gid_b)
                if erro: st.error(erro)
                else:
                    st.session_state.df_b_unif = df
                    st.success(f"Tabela B carregada ({df.shape[0]} linhas).")
        else:
            up_b = st.file_uploader("Arquivo da Tabela B", type=['csv','xlsx'], key="unif_up_b")
            if up_b:
                df, erro = carregar_arquivo_local(up_b)
                if erro: st.error(erro)
                else:
                    st.session_state.df_b_unif = df
                    st.success(f"Tabela B carregada ({df.shape[0]} linhas).")

    # Passo 2: Configurar a unificação
    if 'df_a_unif' in st.session_state and 'df_b_unif' in st.session_state:
        df_a = st.session_state.df_a_unif
        df_b = st.session_state.df_b_unif
        st.divider()
        st.header("Passo 2: Configurar a Unificação")
        
        col_cfg1, col_cfg2 = st.columns(2)
        with col_cfg1:
            key_a = st.selectbox("Coluna-chave da Tabela A", df_a.columns)
            tipo_uniao = st.radio("Tipo de União", 
                options=['Manter todos da Tabela A (Left Join)', 'Manter apenas correspondências (Inner Join)'],
                captions=["Traz informações da Tabela B para a A. O mais comum.", "Mantém apenas as linhas que existem em ambas as tabelas."])
        with col_cfg2:
            key_b = st.selectbox("Coluna-chave da Tabela B", df_b.columns)
            cols_b = st.multiselect("Quais colunas da Tabela B você quer adicionar?", df_b.columns, default=list(df_b.columns))

        if st.button("Unificar Tabelas", type="primary"):
            if not key_a or not key_b or not cols_b:
                st.warning("Por favor, selecione as colunas-chave e as colunas da Tabela B.")
            else:
                try:
                    with st.spinner("Unificando..."):
                        # Mapeia a opção do rádio para o parâmetro do pandas
                        how_param = 'left' if 'Left Join' in tipo_uniao else 'inner'
                        
                        # Garante que a coluna chave da Tabela B esteja na lista de colunas a serem puxadas
                        if key_b not in cols_b:
                            cols_b.insert(0, key_b)

                        df_final = pd.merge(
                            left=df_a,
                            right=df_b[cols_b],
                            left_on=key_a,
                            right_on=key_b,
                            how=how_param,
                            suffixes=('_A', '_B') # Adiciona sufixo se houver colunas com mesmo nome
                        )
                        st.session_state.df_unificado_final = df_final
                        st.success("Tabelas unificadas com sucesso!")
                except Exception as e:
                    st.error(f"Ocorreu um erro na unificação: {e}")

    # Passo 3: Mostrar resultado
    if 'df_unificado_final' in st.session_state:
        st.divider()
        st.header("Resultado da Unificação")
        df_resultado = st.session_state.df_unificado_final
        st.write(f"A tabela final possui **{df_resultado.shape[0]} linhas** e **{df_resultado.shape[1]} colunas**.")
        st.dataframe(df_resultado)
        st.markdown(gerar_download(df_resultado, "planilha_unificada"), unsafe_allow_html=True)

def modulo_agrupador():
    st.title("Agrupador e Sumarizador (Tabela Dinâmica)")
    st.write("Agrupe dados por categorias e realize cálculos como soma, média, contagem, etc.")

    if st.button("← Voltar ao Menu Principal"):
        reset_agrupador_state()
        st.session_state.pagina_atual = "menu"
        st.rerun()

    st.warning("Sua planilha do Google Sheets deve estar em modo público. Acesse \"Compartilhar\" e ative a opção \"Qualquer pessoa com o link pode visualizar\".")

    # Passo 1: Carregar dados
    with st.expander("Passo 1: Carregar Dados", expanded=('df_original_agrup' not in st.session_state)):
        origem = st.radio("Origem da Planilha", ["Google Sheets", "Upload de Arquivo"], key="agrup_origem")
        if origem == "Google Sheets":
            id_sheet = st.text_input("ID da Planilha", key="agrup_id", help=HELP_TEXT_SHEET_ID)
            gid_sheet = st.text_input("GID da Aba", key="agrup_gid", help=HELP_TEXT_GID)
            if st.button("Carregar Dados"):
                df, erro = carregar_planilha(id_sheet, gid_sheet)
                if erro: st.error(erro)
                else:
                    st.session_state.df_original_agrup = df
                    st.success("Planilha carregada!")
        else:
            upload = st.file_uploader("Selecione um arquivo", type=['csv','xlsx'], key="agrup_upload")
            if upload:
                df, erro = carregar_arquivo_local(upload)
                if erro: st.error(erro)
                else:
                    st.session_state.df_original_agrup = df
                    st.success("Planilha carregada!")

    # Passo 2: Configurar agrupamento
    if 'df_original_agrup' in st.session_state:
        df = st.session_state.df_original_agrup
        st.header("Passo 2: Configurar Agrupamento")

        # --- ALTERAÇÃO PRINCIPAL: REMOVIDO O FILTRO DE COLUNAS ---
        # O usuário agora pode selecionar QUALQUER coluna.
        todas_as_colunas = df.columns.tolist()
        
        cols_agrupar = st.multiselect("Agrupar por (dimensões):", options=todas_as_colunas)
        col_calcular = st.selectbox("Coluna para calcular (métrica):", options=todas_as_colunas) # <-- AQUI ESTÁ A MUDANÇA

        mapa_funcoes = {"Soma": "sum", "Média": "mean", "Contagem": "count", "Valor Máximo": "max", "Valor Mínimo": "min"}
        funcoes = st.multiselect("Cálculos a fazer:", options=list(mapa_funcoes.keys()), default=["Soma"])

        if st.button("Agrupar e Calcular", type="primary"):
            if not cols_agrupar or not col_calcular or not funcoes:
                st.warning("Por favor, preencha todos os campos de configuração.")
            else:
                try:
                    with st.spinner("Calculando..."):
                        df_copia = df.copy()
                        
                        # A conversão para número continua aqui como uma salvaguarda.
                        # Se a coluna não for numérica, ela se tornará uma coluna de 'NaN' (nulos).
                        # As funções de agregação (soma, média) ignoram nulos, resultando em 0 ou NaN,
                        # o que evita que o programa quebre.
                        df_copia[col_calcular] = pd.to_numeric(df_copia[col_calcular], errors='coerce')
                        
                        funcoes_pd = [mapa_funcoes[f] for f in funcoes]
                        
                        df_agrupado = df_copia.groupby(cols_agrupar, as_index=False).agg({
                            col_calcular: funcoes_pd
                        })
                        
                        # Aplaina os nomes das colunas se houver múltiplas agregações
                        df_agrupado.columns = ['_'.join(col).strip() if isinstance(col, tuple) and col[1]!='' else col[0] if isinstance(col, tuple) else col for col in df_agrupado.columns.values]

                        st.session_state.df_agrupado_final = df_agrupado
                        st.success("Agrupamento concluído!")
                except Exception as e:
                    st.error(f"Ocorreu um erro ao agrupar: {e}")

    # Passo 3: Exibir resultados
    if 'df_agrupado_final' in st.session_state:
        st.divider()
        st.header("Resultado do Agrupamento")
        df_resultado = st.session_state.df_agrupado_final
        st.dataframe(df_resultado)
        st.markdown(gerar_download(df_resultado, "dados_agrupados"), unsafe_allow_html=True)

def modulo_limpador():
    st.title("Limpador e Padronizador de Dados")
    st.write("Aplique diversas ações de limpeza para melhorar a qualidade da sua planilha.")
    
    if st.button("← Voltar ao Menu Principal"):
        reset_limpador_state()
        st.session_state.pagina_atual = "menu"
        st.rerun()

    st.warning("Sua planilha do Google Sheets deve estar em modo público. Acesse \"Compartilhar\" e ative a opção \"Qualquer pessoa com o link pode visualizar\".")

    # Passo 1: Carregar dados
    with st.expander("Passo 1: Carregar Planilha", expanded=('df_original_limp' not in st.session_state)):
        # (código de carregamento igual ao módulo agrupador, com chaves 'limp_*')
        origem = st.radio("Origem", ["Google Sheets", "Upload"], key="limp_origem")
        if origem == "Google Sheets":
            id_s = st.text_input("ID Planilha", key="limp_id", help=HELP_TEXT_SHEET_ID)
            gid_s = st.text_input("GID Aba", key="limp_gid", help=HELP_TEXT_GID)
            if st.button("Carregar"):
                df, erro = carregar_planilha(id_s, gid_s)
                if erro: st.error(erro)
                else: 
                    st.session_state.df_original_limp = df
                    st.success("Carregado!"); 
        else:
            up = st.file_uploader("Arquivo", type=['csv','xlsx'], key="limp_upload")
            if up:
                df, erro = carregar_arquivo_local(up)
                if erro: st.error(erro)
                else: 
                    st.session_state.df_original_limp = df
                    st.success("Carregado!"); 

    # Passo 2: Configurar limpeza
    if 'df_original_limp' in st.session_state:
        df = st.session_state.df_original_limp
        st.header("Passo 2: Configurar Ações de Limpeza")

        with st.form("form_limpeza"):
            st.write("Marque as ações que deseja aplicar. Elas serão executadas na ordem abaixo.")
            
            # Ação 1: Remover Duplicatas
            with st.expander("1. Remover Duplicatas"):
                st.session_state.limp_duplicatas = st.checkbox("Ativar remoção de duplicatas")
                st.session_state.limp_duplicatas_cols = st.multiselect(
                    "Verificar duplicatas com base nas colunas:", df.columns, default=list(df.columns),
                    help="Linhas com valores idênticos em todas estas colunas serão removidas."
                )

            # Ação 2: Tratar Vazios
            with st.expander("2. Tratar Valores Vazios"):
                st.session_state.limp_vazios = st.checkbox("Ativar tratamento de valores vazios")
                st.session_state.limp_vazios_acao = st.radio("Ação:", ["Remover linhas com qualquer valor vazio", "Preencher valores vazios"], horizontal=True)
                if st.session_state.limp_vazios_acao == "Preencher valores vazios":
                    st.session_state.limp_vazios_valor = st.text_input("Preencher com o valor:", "0")

            # Ação 3: Padronizar Texto
            with st.expander("3. Padronizar Colunas de Texto"):
                st.session_state.limp_texto = st.checkbox("Ativar padronização de texto")
                colunas_texto = df.select_dtypes(include='object').columns.tolist()
                st.session_state.limp_texto_cols = st.multiselect("Colunas para padronizar:", colunas_texto, default=colunas_texto)
                st.session_state.limp_texto_acoes = st.multiselect("Ações de padronização:", ["Remover espaços extras (início/fim)", "Converter para MAIÚSCULAS", "Converter para minúsculas"])

            submitted = st.form_submit_button("Aplicar Limpeza", type="primary")

            if submitted:
                with st.spinner("Processando limpeza..."):
                    df_processado = df.copy()
                    
                    if st.session_state.get('limp_duplicatas'):
                        subset = st.session_state.get('limp_duplicatas_cols')
                        if subset:
                            df_processado.drop_duplicates(subset=subset, inplace=True)

                    if st.session_state.get('limp_vazios'):
                        if st.session_state.get('limp_vazios_acao') == "Remover linhas com qualquer valor vazio":
                            df_processado.dropna(inplace=True)
                        else:
                            valor_preencher = st.session_state.get('limp_vazios_valor', '')
                            df_processado.fillna(valor_preencher, inplace=True)
                    
                    if st.session_state.get('limp_texto') and st.session_state.get('limp_texto_cols') and st.session_state.get('limp_texto_acoes'):
                        for col in st.session_state.limp_texto_cols:
                            if "Remover espaços extras (início/fim)" in st.session_state.limp_texto_acoes:
                                df_processado[col] = df_processado[col].str.strip()
                            if "Converter para MAIÚSCULAS" in st.session_state.limp_texto_acoes:
                                df_processado[col] = df_processado[col].str.upper()
                            if "Converter para minúsculas" in st.session_state.limp_texto_acoes:
                                df_processado[col] = df_processado[col].str.lower()
                    
                    st.session_state.df_limpo_final = df_processado
                    st.success("Limpeza concluída!")
    
    # Passo 3: Exibir resultado
    if 'df_limpo_final' in st.session_state:
        st.divider()
        st.header("Resultado da Limpeza")
        df_resultado = st.session_state.df_limpo_final
        st.write(f"Linhas antes: {df.shape[0]} | Linhas depois: {df_resultado.shape[0]}")
        st.dataframe(df_resultado)
        st.markdown(gerar_download(df_resultado, "planilha_limpa"), unsafe_allow_html=True)

def modulo_divisor():
    st.title("Divisor de Planilhas")
    st.write("Divida uma planilha grande em vários arquivos menores com base nos valores de uma coluna.")

    if st.button("← Voltar ao Menu Principal"):
        reset_divisor_state()
        st.session_state.pagina_atual = "menu"
        st.rerun()

    st.warning("Sua planilha do Google Sheets deve estar em modo público. Acesse \"Compartilhar\" e ative a opção \"Qualquer pessoa com o link pode visualizar\".")
    
    # Passo 1: Carregar
    with st.expander("Passo 1: Carregar Planilha", expanded=('df_original_div' not in st.session_state)):
        # (código de carregamento igual aos outros, com chaves 'div_*')
        origem = st.radio("Origem", ["Google Sheets", "Upload"], key="div_origem")
        if origem == "Google Sheets":
            id_d = st.text_input("ID Planilha", key="div_id", help=HELP_TEXT_SHEET_ID)
            gid_d = st.text_input("GID Aba", key="div_gid", help=HELP_TEXT_GID)
            if st.button("Carregar"):
                df, erro = carregar_planilha(id_d, gid_d)
                if erro: st.error(erro)
                else: 
                    st.session_state.df_original_div = df
                    st.success("Carregado!"); 
        else:
            up = st.file_uploader("Arquivo", type=['csv','xlsx'], key="div_upload")
            if up:
                df, erro = carregar_arquivo_local(up)
                if erro: st.error(erro)
                else: 
                    st.session_state.df_original_div = df
                    st.success("Carregado!"); 

    # Passo 2: Configurar Divisão
    if 'df_original_div' in st.session_state:
        df = st.session_state.df_original_div
        st.header("Passo 2: Configurar Divisão")
        
        coluna_divisao = st.selectbox("Selecione a coluna para usar como critério de divisão:", df.columns)
        formato_saida = st.radio("Formato dos arquivos de saída:", ["Excel (XLSX)", "CSV"], horizontal=True)
        
        if st.button("Dividir Planilha", type="primary"):
            if not coluna_divisao:
                st.warning("Selecione uma coluna para a divisão.")
            else:
                with st.spinner("Dividindo arquivos..."):
                    valores_unicos = df[coluna_divisao].dropna().unique()
                    dfs_para_zipar = {}
                    
                    for valor in valores_unicos:
                        df_parte = df[df[coluna_divisao] == valor]
                        # Limpa o nome do arquivo para evitar caracteres inválidos
                        nome_arquivo = re.sub(r'[^\w\s-]', '', str(valor)).strip().replace(' ', '_')
                        dfs_para_zipar[nome_arquivo] = df_parte
                    
                    formato = "XLSX" if "Excel" in formato_saida else "CSV"
                    zip_data = gerar_download_zip(dfs_para_zipar, formato)
                    st.session_state.dados_zip_div = zip_data
                    st.session_state.nomes_arquivos_div = list(dfs_para_zipar.keys())
                    st.success(f"{len(dfs_para_zipar)} arquivos gerados com sucesso!")

    # Passo 3: Download do ZIP
    if 'dados_zip_div' in st.session_state:
        st.divider()
        st.header("Passo 3: Baixar Arquivos")
        st.info(f"Arquivos gerados: {', '.join(st.session_state.nomes_arquivos_div)}")
        
        st.download_button(
            label="📥 Baixar todos os arquivos (.zip)",
            data=st.session_state.dados_zip_div,
            file_name="planilhas_divididas.zip",
            mime="application/zip"
        )

# --- Página de Login ---
def pagina_login():
    st.title("Ferramentas de Planilhas - GRB")
    st.subheader("Autenticação")
    
    with st.form("login_form"):
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        submitted = st.form_submit_button("Entrar", type="primary")

        if submitted:
            if usuario in USUARIOS_CADASTRADOS and senha == USUARIOS_CADASTRADOS[usuario]:
                cookie_manager.set(
                    'user_session', 
                    usuario, 
                    expires_at=datetime.datetime.now() + datetime.timedelta(days=1)
                )
                
                st.session_state.autenticado = True
                st.session_state.usuario_logado = usuario
                st.session_state.pagina_atual = "menu"
                st.rerun()
            else:
                st.error("Credenciais inválidas. Tente novamente.")


def menu_principal():
    st.title(f"Ferramentas de Planilhas - GRB")
    st.markdown(f"**Usuário:** `{st.session_state.get('usuario_logado', 'Convidado')}`")
    st.write("### Selecione a ferramenta que deseja utilizar:")
    
    st.write("---")
   
    col1, col2, col3 = st.columns(3)
    with col1:
        with st.container(border=True):
            st.subheader("Unificador")
            st.write("Cruze informações de duas tabelas (PROCV).")
            if st.button("Acessar Unificador", key="btn_unif", use_container_width=True):
                st.session_state.pagina_atual = "unificador"
                st.rerun()
    with col2:
        with st.container(border=True):
            st.subheader("Agrupador")
            st.write("Crie sumarizações e totais (Tabela Dinâmica).")
            if st.button("Acessar Agrupador", key="btn_agrup", use_container_width=True):
                st.session_state.pagina_atual = "agrupador"
                st.rerun()
    with col3:
        with st.container(border=True):
            st.subheader("Limpador")
            st.write("Remova duplicatas, preencha vazios e padronize.")
            if st.button("Acessar Limpador", key="btn_limp", use_container_width=True):
                st.session_state.pagina_atual = "limpador"
                st.rerun()
    
    st.write("---")
   
    col4, col5, col6, col7 = st.columns(4)
    with col4:
        with st.container(border=True):
            st.subheader("Divisor")
            st.write("Divida uma planilha em várias menores.")
            if st.button("Acessar Divisor", key="btn_div", use_container_width=True):
                st.session_state.pagina_atual = "divisor"
                st.rerun()
    with col5:
        with st.container(border=True):
            st.subheader("Comparador")
            st.write("Compare duas tabelas e veja as diferenças.")
            if st.button("Acessar Comparador", key="btn1", use_container_width=True):
                st.session_state.pagina_atual = "comparador"
                st.rerun()
    with col6:
        with st.container(border=True):
            st.subheader("Filtro Avançado")
            st.write("Filtre dados com múltiplos critérios.")
            if st.button("Acessar Filtro", key="btn2", use_container_width=True):
                st.session_state.pagina_atual = "filtro"
                st.rerun()
    with col7:
         with st.container(border=True):
            st.subheader("Renomeador")
            st.write("Renomeie e organize as colunas.")
            if st.button("Acessar Renomeador", key="btn3", use_container_width=True):
                st.session_state.pagina_atual = "renomeador"
                st.rerun()
    
    st.divider()
    if st.button("Sair do Sistema"):
        cookie_manager.set('user_session', '', expires_at=datetime.datetime.now() - datetime.timedelta(days=1))
        st.session_state.clear()
        st.session_state.autenticado = False
        st.session_state.pagina_atual = "login"
        st.rerun()

# --- Bloco Principal de Navegação 
user_from_cookie = cookie_manager.get('user_session')

if not st.session_state.get("autenticado") and user_from_cookie:
    st.session_state.autenticado = True
    st.session_state.usuario_logado = user_from_cookie
    if "pagina_atual" not in st.session_state:
        st.session_state.pagina_atual = "menu"

if not st.session_state.get("autenticado"):
    st.session_state.pagina_atual = "login"

pagina = st.session_state.get("pagina_atual", "login")

if pagina == "login":
    pagina_login()
elif pagina == "menu":
    menu_principal()
elif pagina == "comparador":
    modulo_comparador()
elif pagina == "filtro":
    modulo_filtro()
elif pagina == "renomeador":
    modulo_renomeador()
elif pagina == "unificador":
    modulo_unificador()
elif pagina == "agrupador":
    modulo_agrupador()
elif pagina == "limpador":
    modulo_limpador()
elif pagina == "divisor":
    modulo_divisor()
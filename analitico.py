import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Configuração da página
st.set_page_config(
    page_title="Processador de Planilhas",
    page_icon="📊",
    layout="wide"
)

# Inicializar session state
if 'df_processado' not in st.session_state:
    st.session_state.df_processado = None
if 'descricoes_religacao' not in st.session_state:
    st.session_state.descricoes_religacao = set()
if 'descricoes_fiscalizacao' not in st.session_state:
    st.session_state.descricoes_fiscalizacao = set()

def carregar_cidades():
    """Carrega a lista de cidades do Google Sheets"""
    google_sheet_url = 'https://docs.google.com/spreadsheets/d/1s6KPkKB45R_c6Gc8U6QlgKe3yRvzHBbzbLEf4PXcTQk/export?format=csv'
    try:
        df_cidades = pd.read_csv(google_sheet_url)
        if len(df_cidades.columns) > 0:
            coluna_cidades = df_cidades.columns[0]
            cidades = set(df_cidades[coluna_cidades].dropna().astype(str).str.strip().str.upper())
            cidades = {cidade for cidade in cidades if cidade and cidade != 'NAN'}
            return cidades
        return set()
    except Exception as e:
        st.error(f"Erro ao carregar cidades: {str(e)}")
        return set()

def carregar_equipes():
    """Carrega a lista de equipes do Google Sheets"""
    google_sheet_url = 'https://docs.google.com/spreadsheets/d/1piSbAO3yHdUpEQ7fCQp7UBTMC2apeMNXwEVCPZabmXM/export?format=csv'
    try:
        df_equipes = pd.read_csv(google_sheet_url)
        
        descricoes_religacao = set()
        descricoes_fiscalizacao = set()
        
        # Coluna A (RELIGAÇÃO)
        if 'RELIGAÇÃO' in df_equipes.columns:
            religacao_valores = df_equipes['RELIGAÇÃO'].dropna().tolist()
            descricoes_religacao = {str(desc).strip().upper() for desc in religacao_valores if pd.notna(desc) and str(desc).strip()}
        elif len(df_equipes.columns) > 0:
            religacao_valores = df_equipes.iloc[:, 0].dropna().tolist()
            descricoes_religacao = {str(desc).strip().upper() for desc in religacao_valores if pd.notna(desc) and str(desc).strip()}
        
        # Coluna B (FISCALIZAÇÃO)
        if 'FISCALIZAÇÃO' in df_equipes.columns:
            fiscalizacao_valores = df_equipes['FISCALIZAÇÃO'].dropna().tolist()
            descricoes_fiscalizacao = {str(desc).strip().upper() for desc in fiscalizacao_valores if pd.notna(desc) and str(desc).strip()}
        elif len(df_equipes.columns) > 1:
            fiscalizacao_valores = df_equipes.iloc[:, 1].dropna().tolist()
            descricoes_fiscalizacao = {str(desc).strip().upper() for desc in fiscalizacao_valores if pd.notna(desc) and str(desc).strip()}
        
        st.session_state.descricoes_religacao = descricoes_religacao
        st.session_state.descricoes_fiscalizacao = descricoes_fiscalizacao
        
        return descricoes_religacao.union(descricoes_fiscalizacao)
    except Exception as e:
        st.error(f"Erro ao carregar equipes: {str(e)}")
        return set()

def determinar_tipo_servico(servico_desc):
    """Determina se um serviço é de RELIGAÇÃO ou FISCALIZAÇÃO"""
    servico_upper = str(servico_desc).strip().upper()
    
    if servico_upper in st.session_state.descricoes_religacao:
        return "RELIGAÇÃO"
    elif servico_upper in st.session_state.descricoes_fiscalizacao:
        return "FISCALIZAÇÃO"
    else:
        return "INDEFINIDO"

def processar_planilha(uploaded_file):
    """Processa a planilha uploaded"""
    try:
        # Carregar arquivo
        df_original = pd.read_excel(uploaded_file, engine='openpyxl')
        st.info(f"Arquivo carregado: {len(df_original)} linhas e {len(df_original.columns)} colunas")
        
        # Carregar e filtrar por cidades
        cidades_permitidas = carregar_cidades()
        if cidades_permitidas and len(df_original.columns) > 6:
            coluna_g = df_original.columns[6]
            mask_cidades = df_original[coluna_g].astype(str).str.strip().str.upper().isin(cidades_permitidas)
            df_filtrado = df_original[mask_cidades].copy()
            st.info(f"Após filtro por cidades: {len(df_filtrado)} linhas")
        else:
            df_filtrado = df_original.copy()
        
        # Carregar e filtrar por equipes
        descricoes_equipes = carregar_equipes()
        if descricoes_equipes:
            coluna_descricao = None
            for i, col in enumerate(df_filtrado.columns):
                valores_coluna = df_filtrado[col].astype(str).str.strip().str.upper()
                intersecao = set(valores_coluna.values).intersection(descricoes_equipes)
                if len(intersecao) > 0:
                    coluna_descricao = col
                    st.info(f"Coluna de equipes: {coluna_descricao}")
                    break
            
            if coluna_descricao is not None:
                mask_equipes = df_filtrado[coluna_descricao].astype(str).str.strip().str.upper().isin(descricoes_equipes)
                df_processado = df_filtrado[mask_equipes].copy()
                st.info(f"Após filtro por equipes: {len(df_processado)} linhas")
            else:
                df_processado = df_filtrado.copy()
        else:
            df_processado = df_filtrado.copy()
        
        st.session_state.df_processado = df_processado
        st.success(f"Processamento concluído! {len(df_processado)} linhas disponíveis")
        return True
        
    except Exception as e:
        st.error(f"Erro durante processamento: {str(e)}")
        return False

def aplicar_filtro_status(df, status_selecionados):
    """Aplica filtro por status na coluna V"""
    if df is None or len(df) == 0 or not status_selecionados:
        return None
    
    if len(df.columns) > 21:
        coluna_v = df.columns[21]
        padrao_status = "|".join(status_selecionados)
        mask_status = df[coluna_v].astype(str).str.contains(padrao_status, case=False, na=False, regex=True)
        return df[mask_status].copy()
    return df.copy()

def preparar_dados_visualizacao(df_filtrado):
    """Prepara dados para visualização"""
    if df_filtrado is None or len(df_filtrado) == 0:
        return None
    
    # Selecionar colunas: B, G, H, Z, J, V, X, Y, W
    colunas_desejadas = [1, 6, 7, 25, 9, 21, 23, 24, 22]
    colunas_existentes = [i for i in colunas_desejadas if i < len(df_filtrado.columns)]
    
    nomes_colunas = ["Ordem_Servico", "Cidades", "Matricula", "Data_Limite", 
                     "Serviço", "Situacao", "Parecer_Nao_Execucao", 
                     "Motivo_Nao_Execucao", "Parecer_Solicitante"][:len(colunas_existentes)]
    
    df_selecionado = df_filtrado.iloc[:, colunas_existentes].copy()
    df_selecionado.columns = nomes_colunas
    
    # Formatar Data_Limite
    if "Data_Limite" in df_selecionado.columns:
        df_selecionado['Data_Limite'] = df_selecionado['Data_Limite'].fillna('')
        mask_nao_vazio = df_selecionado['Data_Limite'] != ''
        if mask_nao_vazio.any():
            df_temp = df_selecionado.loc[mask_nao_vazio, 'Data_Limite'].copy()
            df_temp_convertido = pd.to_datetime(df_temp, errors='coerce', dayfirst=True)
            df_temp_formatado = df_temp_convertido.dt.strftime('%d/%m/%Y')
            df_selecionado.loc[mask_nao_vazio, 'Data_Limite'] = df_temp_formatado.fillna('')
    
    # Adicionar tipo de equipe
    if "Serviço" in df_selecionado.columns:
        df_selecionado['Tipo_Equipe'] = df_selecionado['Serviço'].apply(determinar_tipo_servico)
    
    # Limpar campos de texto
    for col in ['Parecer_Nao_Execucao', 'Motivo_Nao_Execucao']:
        if col in df_selecionado.columns:
            df_selecionado[col] = df_selecionado[col].fillna('').astype(str).replace(['nan', 'NaT', 'None'], '')
            if 'Situacao' in df_selecionado.columns:
                mask_nao_postergada = ~df_selecionado['Situacao'].astype(str).str.contains('Postergada', case=False, na=False)
                df_selecionado.loc[mask_nao_postergada, col] = ''
    
    if 'Parecer_Solicitante' in df_selecionado.columns:
        df_selecionado['Parecer_Solicitante'] = df_selecionado['Parecer_Solicitante'].fillna('').astype(str).replace(['nan', 'NaT', 'None'], '')
    
    # Calcular dias de atraso
    if 'Data_Limite' in df_selecionado.columns:
        df_selecionado['Atrasados'] = ''
        data_atual = datetime.now().date()
        
        for idx, row in df_selecionado.iterrows():
            data_limite_str = str(row['Data_Limite']).strip()
            if data_limite_str and data_limite_str not in ['', 'nan', 'NaT', 'None']:
                try:
                    if '/' in data_limite_str:
                        data_limite = datetime.strptime(data_limite_str, '%d/%m/%Y').date()
                        dias_atraso = (data_atual - data_limite).days
                        if dias_atraso > 0:
                            df_selecionado.at[idx, 'Atrasados'] = str(dias_atraso)
                except:
                    pass
    
    return df_selecionado

def converter_df_para_excel(df):
    """Converte DataFrame para bytes Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# Interface Principal
st.title("📊 Processador de Planilhas")
st.markdown("---")

# Upload de arquivo
uploaded_file = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    if st.button("Processar Planilha", type="primary"):
        with st.spinner("Processando..."):
            processar_planilha(uploaded_file)

# Visualização e Filtros
if st.session_state.df_processado is not None:
    st.markdown("---")
    st.subheader("Filtros e Visualização")
    
    # Filtros de status
    col1, col2, col3 = st.columns(3)
    with col1:
        status_pendente = st.checkbox("Pendente", value=True)
    with col2:
        status_postergada = st.checkbox("Postergada", value=True)
    with col3:
        status_programado = st.checkbox("Programado", value=True)
    
    # Aplicar filtros
    status_selecionados = []
    if status_pendente:
        status_selecionados.append("Pendente")
    if status_postergada:
        status_selecionados.append("Postergada")
    if status_programado:
        status_selecionados.append("Programado")
    
    if status_selecionados:
        df_filtrado = aplicar_filtro_status(st.session_state.df_processado, status_selecionados)
        
        if df_filtrado is not None and len(df_filtrado) > 0:
            st.info(f"Total de registros filtrados: {len(df_filtrado)}")
            
            # Preparar dados
            df_visualizacao = preparar_dados_visualizacao(df_filtrado)
            
            if df_visualizacao is not None:
                # Separar por equipe
                tabs = st.tabs(["RELIGAÇÃO", "FISCALIZAÇÃO", "INDEFINIDO", "TODOS"])
                
                for i, tipo_equipe in enumerate(["RELIGAÇÃO", "FISCALIZAÇÃO", "INDEFINIDO"]):
                    with tabs[i]:
                        df_equipe = df_visualizacao[df_visualizacao['Tipo_Equipe'] == tipo_equipe].drop(columns=['Tipo_Equipe'])
                        
                        if len(df_equipe) > 0:
                            st.write(f"**{len(df_equipe)} registros**")
                            
                            # Agrupar por serviço
                            if 'Serviço' in df_equipe.columns:
                                servicos_unicos = df_equipe['Serviço'].unique()
                                
                                for servico in servicos_unicos:
                                    df_servico = df_equipe[df_equipe['Serviço'] == servico]
                                    
                                    with st.expander(f"**{servico}** ({len(df_servico)} registros)"):
                                        st.dataframe(df_servico, use_container_width=True, hide_index=True)
                                        
                                        # Botão para baixar este serviço
                                        excel_data = converter_df_para_excel(df_servico)
                                        st.download_button(
                                            label=f"Baixar {servico}",
                                            data=excel_data,
                                            file_name=f"{servico}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
                            else:
                                st.dataframe(df_equipe, use_container_width=True, hide_index=True)
                            
                            # Botão para baixar toda a equipe
                            excel_equipe = converter_df_para_excel(df_equipe)
                            st.download_button(
                                label=f"Baixar todos de {tipo_equipe}",
                                data=excel_equipe,
                                file_name=f"{tipo_equipe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{tipo_equipe}"
                            )
                        else:
                            st.info(f"Nenhum registro para {tipo_equipe}")
                
                # Aba TODOS
                with tabs[3]:
                    df_todos = df_visualizacao.drop(columns=['Tipo_Equipe'])
                    st.write(f"**{len(df_todos)} registros totais**")
                    st.dataframe(df_todos, use_container_width=True, hide_index=True)
                    
                    excel_todos = converter_df_para_excel(df_todos)
                    st.download_button(
                        label="Baixar Todos os Registros",
                        data=excel_todos,
                        file_name=f"todos_registros_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_todos"
                    )
        else:
            st.warning("Nenhum registro encontrado com os filtros selecionados")
    else:
        st.warning("Selecione pelo menos um status")
    
    # Botão para baixar arquivo processado completo
    st.markdown("---")
    excel_processado = converter_df_para_excel(st.session_state.df_processado)
    st.download_button(
        label="Baixar Arquivo Processado Completo (sem filtro de status)",
        data=excel_processado,
        file_name=f"processado_completo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_completo"
    )

# Instruções
with st.sidebar:
    st.header("Instruções")
    st.markdown("""
    1. **Upload**: Selecione arquivo Excel (.xlsx)
    2. **Processar**: Clique em 'Processar Planilha'
    3. **Filtrar**: Marque os status desejados
    4. **Visualizar**: Navegue pelas abas de equipes
    5. **Baixar**: Use os botões de download
    
    **Recursos:**
    - Filtra por cidades e equipes automaticamente
    - Separa RELIGAÇÃO e FISCALIZAÇÃO
    - Calcula dias de atraso
    - Exporta dados filtrados
    - Funciona em qualquer dispositivo
    """)
    
    st.markdown("---")
    st.caption("Versão Streamlit - Multiplataforma")

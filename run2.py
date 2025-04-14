import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io, os, re, unicodedata
from datetime import datetime
import numpy as np
import math
import openai
from difflib import SequenceMatcher

# ATENÇÃO: Se estiver usando openai>=1.0.0, execute "openai migrate"
# ou, alternativamente, fixe a versão anterior com "pip install openai==0.28.0"

# Configuração da página
st.set_page_config(
    page_title="Dashboard Comercial - Budget vs Logcomex vs iTracker",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Definição de cores básicas
COLORS = {
    'primary': '#1E88E5',
    'secondary': '#26A69A',
    'success': '#66BB6A',
    'warning': '#FFA726',
    'danger': '#EF5350',
    'background': '#F8F9FA',
    'text': '#212529',
    'light_text': '#6C757D',
    'chart1': ['#1E88E5', '#26A69A', '#66BB6A'],
    'chart2': ['#42A5F5', '#7E57C2', '#26A69A', '#EC407A', '#FFA726']
}

# CSS personalizado
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    body, .block-container {{
        font-family: 'Inter', sans-serif;
    }}
    .main {{
        background-color: var(--background-color, {COLORS['background']});
        padding: 20px;
        animation: fadeIn 1s;
    }}
    .section {{
        background-color: var(--secondaryBackgroundColor, #ffffff);
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.08);
        margin-bottom: 20px;
    }}
    .kpi-card {{
        background-color: var(--primary-background-color, #ffffff);
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 3px 6px rgba(0,0,0,0.1);
        text-align: center;
        border-left: 5px solid {COLORS['primary']};
    }}
    .kpi-card:nth-child(2) {{
        border-left-color: {COLORS['warning']};
    }}
    .kpi-card:nth-child(3) {{
        border-left-color: {COLORS['success']};
    }}
    .kpi-card:nth-child(4) {{
        border-left-color: {COLORS['danger']};
    }}
    .kpi-title {{
        font-size: 14px;
        margin-bottom: 5px;
        color: {COLORS['text']};
    }}
    .kpi-value {{
        font-size: 24px;
        font-weight: bold;
        color: {COLORS['text']};
    }}
    .section-title {{
        color: {COLORS['text']};
        font-size: 20px;
        font-weight: 600;
        margin-bottom: 15px;
        border-bottom: 2px solid {COLORS['light_text']};
        padding-bottom: 8px;
    }}
    .sub-title {{
        color: {COLORS['text']};
        font-size: 16px;
        font-weight: 600;
        margin-bottom: 10px;
    }}
    .sidebar .sidebar-content {{
        background-color: var(--secondaryBackgroundColor, #ffffff);
    }}
    .stButton>button {{
        background-color: var(--primaryColor, {COLORS['primary']});
        color: white;
        border-radius: 5px;
        border: none;
        padding: 8px 16px;
        font-weight: 500;
    }}
    .stButton>button:hover {{
        background-color: #1976D2;
    }}
    .divider {{
        margin-top: 20px;
        margin-bottom: 20px;
        border-top: 1px solid #E9ECEF;
    }}
    @keyframes fadeIn {{
      from {{ opacity: 0; }}
      to {{ opacity: 1; }}
    }}
    .custom-footer {{
        text-align: center;
        padding: 10px;
        border-top: 1px solid #E9ECEF;
        font-size: 13px;
        color: gray;
        margin-top: 20px;
    }}
    .custom-footer span {{
        margin: 0 5px;
    }}
</style>
""", unsafe_allow_html=True)

# Novo cabeçalho estilizado
st.markdown("""
    <style>
        .titulo-dashboard-container {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            margin: 0 auto;
            padding: 25px 20px;
            background: linear-gradient(to right, #F37529, rgba(255, 255, 255, 0.8));
            border-radius: 15px;
            box-shadow: 0 6px 10px rgba(0, 0, 0, 0.3);
        }
        .titulo-dashboard {
            font-size: 50px;
            font-weight: bold;
            color: var(--text-color);
            text-transform: uppercase;
            margin: 0;
        }
        .subtitulo-dashboard {
            font-size: 18px;
            color: #555555;
            margin: 10px 0 0 0;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="titulo-dashboard-container">
        <h1 class="titulo-dashboard">Dashboard de análise comercial de clientes</h1>
        <p class="subtitulo-dashboard">Monitoramento em tempo real do desempenho comercial</p>
    </div>
""", unsafe_allow_html=True)

st.markdown("""
    <hr style="border-top: 3px solid #F37529; margin: 20px 0;">
""", unsafe_allow_html=True)

current_date = datetime.now().strftime("%d de %B de %Y")

def styled_container(title):
    st.markdown(f"<div class='section'><h3 class='section-title'>{title}</h3>", unsafe_allow_html=True)
    return st

def format_number(num):
    if num >= 1000000:
        return f"{num/1000000:.2f}M"
    elif num >= 1000:
        return f"{num/1000:.1f}K"
    else:
        return f"{num:.0f}"

def format_percent(value, positive_is_good=True):
    if value is None or pd.isna(value):
        return "N/A"
    if value >= 100:
        color = COLORS['success'] if positive_is_good else COLORS['danger']
        return f"<span style='color:{color};font-weight:bold'>{value:.1f}%</span>"
    elif value >= 70:
        return f"<span style='color:{COLORS['warning']};font-weight:bold'>{value:.1f}%</span>"
    else:
        color = COLORS['danger'] if positive_is_good else COLORS['success']
        return f"<span style='color:{color};font-weight:bold'>{value:.1f}%</span>"

# Função para carregar os dados com cache
@st.cache_data
def load_data():
    df = download_file_from_gdrive()
    if df is None:
        st.error("Não foi possível carregar os dados do Google Sheets.")
        st.stop()
    # Remove registros sem nome de cliente
    df = df[df['Cliente'].notna() & (df['Cliente'] != "undefined")]
    numeric_cols = ['MÊS', 'BUDGET', 'Importação', 'Exportação', 'Cabotagem', 'Quantidade_iTRACKER']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df['Cliente'] = df['Cliente'].str.upper()
    return df

def download_file_from_gdrive():
    try:
        credentials_info = st.secrets["google"]
        credentials = service_account.Credentials.from_service_account_info(
            credentials_info,
            scopes=['https://www.googleapis.com/auth/drive.readonly']
        )
        drive_service = build('drive', 'v3', credentials=credentials)
        file_id = "1Bphi7lChPqh12kAStpupXJmCbwcdImKo"
        st.sidebar.info("Baixando arquivo real do Google Sheets...")
        request = drive_service.files().get_media(fileId=file_id)
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        done = False
        progress_bar = st.sidebar.progress(0)
        status_text = st.sidebar.empty()
        while not done:
            status, done = downloader.next_chunk()
            progress = int(status.progress() * 100)
            progress_bar.progress(progress)
            status_text.text(f"Download: {progress}%")
        status_text.text("Download concluído!")
        progress_bar.empty()
        file.seek(0)
        df = pd.read_excel(file, engine='openpyxl')
        st.sidebar.success("Arquivo carregado com sucesso!")
        return df
    except Exception as e:
        st.sidebar.error(f"Erro ao acessar o Google Drive: {str(e)}")
        return None

df = load_data()

# Sidebar - Filtros de Análise
st.sidebar.markdown("---")
st.sidebar.markdown("### 🔍 Filtros de Análise")
meses_map = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}
meses_disponiveis = sorted(df['MÊS'].unique())
mes_selecionado = st.sidebar.multiselect(
    "Selecione o(s) mês(es):",
    options=meses_disponiveis,
    format_func=lambda x: meses_map.get(x, x),
    default=meses_disponiveis[0] if meses_disponiveis else None
)
clientes_disponiveis = sorted(df['Cliente'].unique())
cliente_selecionado = st.sidebar.multiselect(
    "Selecione o(s) cliente(s):",
    options=clientes_disponiveis
)
if st.sidebar.button("Limpar Filtros"):
    mes_selecionado = []
    cliente_selecionado = []
st.sidebar.markdown("---")
show_detailed_table = st.sidebar.checkbox("Mostrar tabela detalhada", value=True)
chart_height = st.sidebar.slider("Altura dos gráficos", 400, 800, 800, 50)

if mes_selecionado and cliente_selecionado:
    filtered_df = df[(df['MÊS'].isin(mes_selecionado)) & (df['Cliente'].isin(cliente_selecionado))]
elif mes_selecionado:
    filtered_df = df[df['MÊS'].isin(mes_selecionado)]
elif cliente_selecionado:
    filtered_df = df[df['Cliente'].isin(cliente_selecionado)]
else:
    filtered_df = df.copy()

if mes_selecionado or cliente_selecionado:
    filtros_ativos = []
    if mes_selecionado:
        meses_texto = [meses_map.get(m, m) for m in mes_selecionado]
        filtros_ativos.append(f"Meses: {', '.join(map(str, meses_texto))}")
    if cliente_selecionado:
        filtros_ativos.append(f"Clientes: {', '.join(cliente_selecionado)}")
    st.markdown(f"<div style='background-color:#E3F2FD;padding:10px;border-radius:5px;margin-bottom:20px;'><b>Filtros ativos:</b> {' | '.join(filtros_ativos)}</div>", unsafe_allow_html=True)

# =============================================================================
# SEÇÃO DE KPIS COM CARTÕES APRIMORADOS
# =============================================================================
st.markdown("<div class='section'>", unsafe_allow_html=True)
st.markdown("<h3 class='section-title'>VISÃO GERAL</h3>", unsafe_allow_html=True)
col1, col2, col3, col4 = st.columns(4)
total_budget = filtered_df['BUDGET'].sum()
with col1:
    st.markdown(f"""
    <div class='kpi-card'>
        <p class='kpi-title'>💰 TOTAL BUDGET</p>
        <p class='kpi-value'>{format_number(total_budget)}</p>
    </div>
    """, unsafe_allow_html=True)
total_oportunidades = filtered_df['Importação'].sum() + filtered_df['Exportação'].sum() + filtered_df['Cabotagem'].sum()
with col2:
    st.markdown(f"""
    <div class='kpi-card'>
        <p class='kpi-title'>🧭 TOTAL OPORTUNIDADES</p>
        <p class='kpi-value'>{format_number(total_oportunidades)}</p>
    </div>
    """, unsafe_allow_html=True)
total_itracker = filtered_df['Quantidade_iTRACKER'].sum()
with col3:
    st.markdown(f"""
    <div class='kpi-card'>
        <p class='kpi-title'>🚚 TOTAL REALIZADO (SYSTRACKER)</p>
        <p class='kpi-value'>{format_number(total_itracker)}</p>
    </div>
    """, unsafe_allow_html=True)
with col4:
    st.markdown(f"""
    <div class='kpi-card'>
        <p class='kpi-title'>🎯 PERFORMANCE VS BUDGET (Até Hoje)</p>
        <p class='kpi-value'>{format_percent((total_itracker / total_budget) * 100 if total_budget > 0 else 0)}</p>
    </div>
    """, unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# Configuração da API do OpenAI para utilizar o modelo GPT-4
openai.api_key = st.secrets["openai"]["api_key"]

st.markdown("<div class='section'><h3 class='section-title'>🤖 ASSISTENTE COMERCIAL</h3></div>", unsafe_allow_html=True)

# Função para normalizar e remover acentuação de strings
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = text.lower().strip()
    return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')

# FUNÇÃO APRIMORADA PARA EXTRAIR O NOME DO CLIENTE
def extrair_nome_cliente(pergunta_input, lista_clientes):
    """
    Identifica o cliente mais provável a partir da pergunta digitada pelo usuário,
    mesmo com nomes incompletos ou parciais (ex: 'dart' → 'DART BRASIL LIMITADA').
    """
    pergunta_norm = normalize_text(pergunta_input)
    palavras_pergunta = set(pergunta_norm.split())

    # Candidatos com pelo menos uma palavra da pergunta contida no nome do cliente
    candidatos = []
    for cliente in lista_clientes:
        cliente_norm = normalize_text(cliente)
        palavras_cliente = set(cliente_norm.split())

        if palavras_pergunta & palavras_cliente:
            candidatos.append(cliente)

    # Se encontrou candidatos, retorna o mais semelhante
    if candidatos:
        melhor_candidato = None
        melhor_ratio = 0.0
        for candidato in candidatos:
            ratio = SequenceMatcher(None, pergunta_norm, normalize_text(candidato)).ratio()
            if ratio > melhor_ratio:
                melhor_ratio = ratio
                melhor_candidato = candidato
        return melhor_candidato

    # Se nada for encontrado, fallback: usar o mais semelhante da lista inteira
    melhor_candidato = None
    melhor_ratio = 0.0
    for cliente in lista_clientes:
        ratio = SequenceMatcher(None, pergunta_norm, normalize_text(cliente)).ratio()
        if ratio > melhor_ratio:
            melhor_ratio = ratio
            melhor_candidato = cliente
    if melhor_ratio >= 0.5:
        return melhor_candidato

    return None

# Função para extrair o mês mencionado na pergunta (por extenso)
def extrair_mes_da_pergunta(pergunta):
    meses = {
        "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
        "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
        "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
    }
    pergunta_norm = normalize_text(pergunta)
    for mes_nome, mes_num in meses.items():
        if mes_nome in pergunta_norm:
            return mes_num
    return None

# Obter a lista de clientes
clientes_lista = df["Cliente"].unique()

# Campo para inserção da pergunta
pergunta_input = st.text_input(
    "Faça uma pergunta sobre desempenho comercial, oportunidades ou status do cliente:",
    placeholder="Ex: Como está o desempenho da empresa OLAM em abril?"
)

# Extração do nome do cliente e do mês (se mencionado) a partir da pergunta
cliente_mencionado = extrair_nome_cliente(pergunta_input, clientes_lista) if pergunta_input else None
mes_extraido = extrair_mes_da_pergunta(pergunta_input) if pergunta_input else None
mes_corrente = mes_extraido if mes_extraido else datetime.now().month
mes_nome = meses_map.get(mes_corrente, f"Mês {mes_corrente}")

# Exibir o cliente identificado (para debug, pode ser removido)
if cliente_mencionado:
    st.info(f"Cliente identificado: {cliente_mencionado}")

# Preparação do contexto para a resposta
if cliente_mencionado:
    df_cliente = df[(df['Cliente'] == cliente_mencionado) & (df['MÊS'] == mes_corrente)]
    if not df_cliente.empty:
        budget_cliente = df_cliente['BUDGET'].sum()
        realizado_cliente = df_cliente['Quantidade_iTRACKER'].sum()
        oportunidades_cliente = df_cliente[['Importação', 'Exportação', 'Cabotagem']].sum().sum()
        performance_cliente = (realizado_cliente / budget_cliente) * 100 if budget_cliente > 0 else 0
        gap_cliente = df_cliente['Gap de Realização'].sum()
        contexto_cliente = (
            f"Dados do cliente {cliente_mencionado} ({mes_nome}):\n"
            f"- Budget: {int(budget_cliente)}\n"
            f"- Realizado: {int(realizado_cliente)}\n"
            f"- Oportunidades: {int(oportunidades_cliente)}\n"
            f"- Performance: {performance_cliente:.1f}%\n"
            f"- Gap: {gap_cliente:.0f}\n"
        )
    else:
        contexto_cliente = f"Não foram encontrados dados para o cliente {cliente_mencionado} no {mes_nome}."
else:
    total_budget = filtered_df['BUDGET'].sum()
    total_itracker = filtered_df['Quantidade_iTRACKER'].sum()
    contexto_cliente = (
        f"Dados Gerais ({mes_nome}):\n"
        f"- Budget Total: {format_number(total_budget)}\n"
        f"- Realizado Total: {format_number(total_itracker)}\n"
        f"- Performance Geral: {round((total_itracker/total_budget)*100,1) if total_budget > 0 else 0}%\n"
    )

# Botão para submeter a pergunta utilizando os dados reais ou GPT-4
if st.button("Enviar Pergunta"):
    if pergunta_input.strip() != "":
        with st.spinner("Buscando informações..."):
            if cliente_mencionado and 'df_cliente' in locals() and not df_cliente.empty:
                status_performance = (
                    "excelente" if performance_cliente >= 100 else
                    "satisfatória" if performance_cliente >= 80 else
                    "média" if performance_cliente >= 70 else
                    "abaixo do esperado" if performance_cliente >= 50 else
                    "crítica"
                )
                
                resposta_texto = (
                    f"📊 **Análise para {cliente_mencionado} ({mes_nome}):**\n\n"
                    f"A performance está **{status_performance}** com **{performance_cliente:.1f}%** do Budget atingido "
                    f"({int(realizado_cliente)} de {int(budget_cliente)} containers).\n\n"
                )
                if oportunidades_cliente > 0:
                    aproveitamento = (realizado_cliente / oportunidades_cliente) * 100
                    resposta_texto += (
                        f"**Aproveitamento de Oportunidades:** {aproveitamento:.1f}% "
                        f"({int(realizado_cliente)} de {int(oportunidades_cliente)} potenciais).\n\n"
                    )
                if performance_cliente < 70:
                    resposta_texto += "**Recomendação:** Necessita atenção prioritária para reduzir o gap de atendimento."
                elif performance_cliente < 100:
                    resposta_texto += "**Recomendação:** Um follow-up adicional pode ajudar a atingir a meta."
                else:
                    resposta_texto += "**Recomendação:** Manter a estratégia atual que tem mostrado eficácia."
            else:
                sistema_prompt = (
                    "Você é um assistente comercial especialista em análise de dados de logística. "
                    "Responda de forma objetiva e resumida (até 3 parágrafos), utilizando exclusivamente os dados fornecidos. "
                    "Inclua números e percentuais exatos e, se aplicável, recomendações claras. "
                    "Se os dados disponíveis não forem suficientes, indique claramente que não foi possível extrair informações precisas.\n\n"
                    f"DADOS DISPONÍVEIS:\n{contexto_cliente}"
                )
                try:
                    resposta = openai.ChatCompletion.create(
                        model="gpt-4",
                        messages=[
                            {"role": "system", "content": sistema_prompt},
                            {"role": "user", "content": pergunta_input}
                        ],
                        temperature=0.3,
                        max_tokens=500
                    )
                    resposta_texto = resposta.choices[0].message['content']
                except Exception as e:
                    resposta_texto = f"⚠️ Erro ao processar sua pergunta: {str(e)}"
        st.markdown(f"**Resposta:**\n\n{resposta_texto}")
        if 'chat_history' not in st.session_state:
            st.session_state.chat_history = []
        st.session_state.chat_history.append({"pergunta": pergunta_input, "resposta": resposta_texto})
        
        if len(st.session_state.chat_history) > 1:
            with st.expander("Ver histórico de perguntas", expanded=False):
                for i, item in enumerate(st.session_state.chat_history[:-1]):
                    st.markdown(f"**Pergunta {i+1}:** {item['pergunta']}")
                    st.markdown(f"**Resposta {i+1}:** {item['resposta']}")
                    st.markdown("---")
    else:
        st.warning("Por favor, insira uma pergunta antes de enviar.")

# =============================================================================
# Gráfico Principal: Clientes com Maior Gap vs Target Acumulado (Mês Corrente)
# =============================================================================
def custom_round(x):
    frac = x - int(x)
    return int(x) + 1 if frac > 0.5 else int(x)

if not filtered_df.empty:
    current_month = datetime.now().month
    df_current = filtered_df[filtered_df['MÊS'] == current_month]
    if not df_current.empty:
        df_gap = df_current.groupby("Cliente", as_index=False).agg({
            "Target Acumulado": "sum",
            "Quantidade_iTRACKER": "sum",
            "Gap de Realização": "sum"
        })
        df_gap['Gap de Realização'] = df_gap['Gap de Realização'].apply(custom_round)
        df_gap = df_gap.sort_values("Gap de Realização", ascending=False)
        df_gap_top = df_gap.head(15)
        
        fig_gap = px.bar(
            df_gap_top,
            x="Gap de Realização",
            y="Cliente",
            orientation="h",
            text="Gap de Realização",
            color="Gap de Realização",
            color_continuous_scale=px.colors.sequential.Reds,
            labels={"Gap de Realização": "Gap de Atendimento"},
            title="CLIENTES COM MAIOR GAP DE ATENDIMENTO (MÊS CORRENTE)"
        )
        fig_gap.update_layout(
            yaxis=dict(autorange="reversed"),
            height=chart_height,
            margin=dict(l=60, r=60, t=40, b=80),
            legend=dict(orientation='h', y=-0.25, x=0.5, xanchor='center'),
            plot_bgcolor="white"
        )
        fig_gap.update_traces(texttemplate='%{text}', textposition='outside')
        
        st.markdown("<div class='section'><h3 class='section-title'>🚨 CLIENTES COM MAIOR GAP VS TARGET ACUMULADO</h3></div>", unsafe_allow_html=True)
        st.plotly_chart(fig_gap, use_container_width=True)
        
        with st.expander("VER DETALHES DO CÁLCULO DESTE GRÁFICO"):
            st.markdown("""
            **DETALHAMENTO DO CÁLCULO:**
            - **FILTRAGEM:** Seleciona os dados referentes ao MÊS CORRENTE.
            - **AGRUPAMENTO:** Agrupa os dados por CLIENTE somando os valores de TARGET ACUMULADO, REALIZADO SYSTRACKER e GAP DE REALIZAÇÃO.
            - **CÁLCULO DO GAP:** O GAP é arredondado usando a função `custom_round`.
            - **ORDENAÇÃO:** Os CLIENTES são ordenados em ordem decrescente com base no GAP DE REALIZAÇÃO.
            """)
    else:
        st.info("NÃO EXISTEM DADOS PARA O MÊS CORRENTE PARA ANÁLISE DE GAP.")

# =============================================================================
# GRÁFICO 1: PERFORMANCE VS BUDGET POR CLIENTE
# =============================================================================
if not filtered_df.empty:
    budget_df = filtered_df[filtered_df['BUDGET'] > 0].copy()
    if not budget_df.empty:
        st.markdown("<h4 class='sub-title'>PERFORMANCE VS BUDGET POR CLIENTE</h4>", unsafe_allow_html=True)
        df_graph3 = budget_df.groupby('Cliente', as_index=False).agg({
            'BUDGET': 'sum',
            'Quantidade_iTRACKER': 'sum'
        })
        df_graph3['Performance'] = (df_graph3['Quantidade_iTRACKER'] / df_graph3['BUDGET']) * 100
        df_graph3 = df_graph3.sort_values('Performance', ascending=False)
        if len(df_graph3) > 15:
            df_graph3 = df_graph3.head(15)
        df_graph3['Color'] = df_graph3['Performance'].apply(
            lambda x: COLORS['success'] if x >= 100 else (COLORS['warning'] if x >= 70 else COLORS['danger'])
        )
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            x=df_graph3['Performance'],
            y=df_graph3['Cliente'],
            orientation='h',
            marker_color=df_graph3['Color'],
            text=df_graph3['Performance'].apply(lambda x: f'{x:.1f}%'),
            hovertemplate='<b>%{y}</b><br>Performance: %{x:.1f}%<br>Budget: %{customdata[0]:,.0f}<br>Realizado: %{customdata[1]:,.0f}<extra></extra>',
            customdata=np.stack((df_graph3['BUDGET'], df_graph3['Quantidade_iTRACKER']), axis=-1)
        ))
        fig3.add_shape(
            type="line",
            x0=100,
            y0=-0.5,
            x1=100,
            y1=len(df_graph3)-0.5,
            line=dict(color="black", width=2, dash="dash")
        )
        fig3.add_shape(
            type="rect",
            x0=0,
            y0=-0.5,
            x1=70,
            y1=len(df_graph3)-0.5,
            line=dict(width=0),
            fillcolor="rgba(239, 83, 80, 0.1)",
            layer="below"
        )
        fig3.add_shape(
            type="rect",
            x0=70,
            y0=-0.5,
            x1=100,
            y1=len(df_graph3)-0.5,
            line=dict(width=0),
            fillcolor="rgba(255, 167, 38, 0.1)",
            layer="below"
        )
        fig3.add_shape(
            type="rect",
            x0=100,
            y0=-0.5,
            x1=df_graph3['Performance'].max() * 1.1,
            y1=len(df_graph3)-0.5,
            line=dict(width=0),
            fillcolor="rgba(102, 187, 106, 0.1)",
            layer="below"
        )
        fig3.add_annotation(
            x=35,
            y=len(df_graph3)-1,
            text="CRÍTICO (<70%)",
            showarrow=False,
            font=dict(color=COLORS['danger']),
            xanchor="center",
            yanchor="top"
        )
        fig3.add_annotation(
            x=85,
            y=len(df_graph3)-1,
            text="ATENÇÃO (70-100%)",
            showarrow=False,
            font=dict(color=COLORS['warning']),
            xanchor="center",
            yanchor="top"
        )
        fig3.add_annotation(
            x=min(150, df_graph3['Performance'].max() * 0.9),
            y=len(df_graph3)-1,
            text="META ATINGIDA (>100%)",
            showarrow=False,
            font=dict(color=COLORS['success']),
            xanchor="center",
            yanchor="top"
        )
        fig3.update_traces(textposition='inside')
        fig3.update_layout(
            xaxis_title='PERFORMANCE (%)',
            yaxis_title='CLIENTE',
            height=chart_height,
            template="plotly",
            margin=dict(l=60, r=30, t=30, b=40),
            xaxis=dict(range=[0, max(200, df_graph3['Performance'].max() * 1.1)])
        )
        st.plotly_chart(fig3, use_container_width=True)
        
        with st.expander("VER DETALHES DO CÁLCULO DESTE GRÁFICO"):
            st.markdown("""
            **DETALHAMENTO DO CÁLCULO:**
            - **FILTRAGEM:** Considera apenas os CLIENTES com BUDGET > 0.
            - **AGRUPAMENTO:** Soma dos valores de BUDGET e REALIZADO SYSTRACKER por CLIENTE.
            - **PERFORMANCE:** Calculada como (REALIZADO SYSTRACKER / BUDGET) * 100.
            - **CORES:** Definidas com base em thresholds (≥100% = SUCESSO, 70-99% = ATENÇÃO, <70% = CRÍTICO).
            """)
        
        total_clientes = len(df_graph3)
        clientes_acima_meta = len(df_graph3[df_graph3['Performance'] >= 100])
        clientes_atencao = len(df_graph3[(df_graph3['Performance'] < 100) & (df_graph3['Performance'] >= 70)])
        clientes_critico = len(df_graph3[df_graph3['Performance'] < 70])
        
        st.markdown(f"""
        <div style='background-color:{COLORS['background']}; padding:10px; border-radius:5px; margin-top:10px;'>
            <h5 style='margin-top:0'>📊 INSIGHTS - PERFORMANCE</h5>
            <ul>
                <li>Dos {total_clientes} CLIENTES analisados:</li>
                <li><span style='color:{COLORS["success"]};'>✓ {clientes_acima_meta} CLIENTES ({clientes_acima_meta/total_clientes*100:.1f}%) ATINGIRAM OU SUPERARAM A META</span></li>
                <li><span style='color:{COLORS["warning"]};'>⚠️ {clientes_atencao} CLIENTES ({clientes_atencao/total_clientes*100:.1f}%) ESTÃO EM ZONA DE ATENÇÃO (70-99%)</span></li>
                <li><span style='color:{COLORS["danger"]};'>❌ {clientes_critico} CLIENTES ({clientes_critico/total_clientes*100:.1f}%) ESTÃO EM SITUAÇÃO CRÍTICA (<70%)</span></li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("SEM DADOS DE BUDGET DISPONÍVEIS PARA OS FILTROS SELECIONADOS.")
st.markdown("</div>", unsafe_allow_html=True)

# =============================================================================
# GRÁFICO 2: COMPARATIVO BUDGET VS REALIZADO POR CATEGORIA (AGRUPADO)
# =============================================================================
if not filtered_df.empty:
    st.markdown("<h4 class='sub-title'>COMPARATIVO BUDGET VS REALIZADO POR CATEGORIA</h4>", unsafe_allow_html=True)
    clientes_top = filtered_df.groupby('Cliente', as_index=False)['BUDGET'].sum()\
                    .sort_values('BUDGET', ascending=False)['Cliente'].head(15)
    df_top = filtered_df[filtered_df['Cliente'].isin(clientes_top)]
    df_grouped = df_top.groupby('Cliente', as_index=False).agg({
        'BUDGET': 'sum',
        'Importação': 'sum',
        'Exportação': 'sum',
        'Cabotagem': 'sum'
    })
    df_grouped['Total'] = df_grouped[['BUDGET', 'Importação', 'Exportação', 'Cabotagem']].sum(axis=1)
    df_grouped = df_grouped.sort_values('Total', ascending=False)
    df_melted = df_grouped.melt(
        id_vars='Cliente',
        value_vars=['BUDGET', 'Importação', 'Exportação', 'Cabotagem'],
        var_name='Categoria',
        value_name='Quantidade'
    )
    df_melted = df_melted[df_melted['Quantidade'] > 0]
    df_melted['Categoria_Label'] = df_melted['Categoria']
    
    fig = px.bar(
        df_melted,
        x='Cliente',
        y='Quantidade',
        color='Categoria',
        barmode='group',
        height=chart_height,
        color_discrete_map={
            'BUDGET': '#0D47A1',
            'Importação': '#00897B',
            'Exportação': '#F4511E',
            'Cabotagem': '#FFB300'
        },
        labels={'Quantidade': 'QTD. DE CONTAINERS'},
        custom_data=['Categoria_Label']
    )
    fig.update_traces(
        texttemplate='%{y:.0f}',
        textposition='outside',
        hovertemplate='<b>CLIENTE:</b> %{x}<br><b>CATEGORIA:</b> %{customdata[0]}<br><b>QTD.:</b> %{y:.0f}<extra></extra>'
    )
    fig.update_layout(
        xaxis=dict(title='CLIENTE', tickangle=-30),
        yaxis=dict(title='QTD. DE CONTAINERS', range=[0, df_melted['Quantidade'].max() * 1.1]),
        legend=dict(orientation='h', y=-0.25, x=0.5, xanchor='center'),
        margin=dict(l=60, r=40, t=20, b=100),
        template='plotly_white',
        bargap=0.25,
        title_text="",
        plot_bgcolor="white"
    )
    st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("VER DETALHES DO CÁLCULO DESTE GRÁFICO"):
        st.markdown("""
        **DETALHAMENTO DO CÁLCULO:**
        - **AGRUPAMENTO:** Dados agrupados por CLIENTE com soma dos valores de:
            - BUDGET
            - IMPORTAÇÃO, EXPORTAÇÃO E CABOTAGEM
        - **CÁLCULO DO TOTAL:** Soma de todas as categorias (BUDGET + IMPORTAÇÃO + EXPORTAÇÃO + CABOTAGEM)
        - O gráfico compara o BUDGET versus o REALIZADO SYSTRACKER por categoria.
        """)
else:
    st.info("SEM DADOS DISPONÍVEIS PARA O GRÁFICO DE COMPARATIVO APÓS APLICAÇÃO DOS FILTROS.")

# =============================================================================
# GRÁFICO 3: APROVEITAMENTO DE OPORTUNIDADES POR CLIENTE
# =============================================================================
if not filtered_df.empty:
    opp_df = filtered_df[(filtered_df['Importação']+filtered_df['Exportação']+filtered_df['Cabotagem']) > 0].copy()
    if not opp_df.empty:
        st.markdown("<h4 class='sub-title'>APROVEITAMENTO DE OPORTUNIDADES POR CLIENTE</h4>", unsafe_allow_html=True)
        df_graph2 = opp_df.groupby('Cliente', as_index=False).agg({
            'Importação': 'sum',
            'Exportação': 'sum',
            'Cabotagem': 'sum',
            'Quantidade_iTRACKER': 'sum'
        })
        df_graph2['Total_Oportunidades'] = df_graph2[['Importação', 'Exportação', 'Cabotagem']].sum(axis=1)
        df_graph2['Aproveitamento'] = (df_graph2['Quantidade_iTRACKER'] / df_graph2['Total_Oportunidades']) * 100
        df_graph2 = df_graph2.sort_values('Aproveitamento', ascending=False)
        if len(df_graph2) > 15:
            df_graph2 = df_graph2.head(15)
        fig2 = px.bar(
            df_graph2,
            x='Cliente',
            y='Aproveitamento',
            color='Aproveitamento',
            color_continuous_scale=px.colors.sequential.Blues,
            text_auto='.1f',
            labels={'Aproveitamento': 'TAXA DE APROVEITAMENTO (%)'},
            custom_data=['Total_Oportunidades', 'Quantidade_iTRACKER']
        )
        fig2.update_traces(
            texttemplate='%{y:.1f}%',
            textposition='outside',
            hovertemplate=(
                '<b>CLIENTE:</b> %{x}<br>'
                '<b>TAXA DE APROVEITAMENTO:</b> %{y:.1f}%<br>'
                '<b>TOTAL OPORTUNIDADES:</b> %{customdata[0]:,.0f}<br>'
                '<b>REALIZADO:</b> %{customdata[1]:,.0f}<extra></extra>'
            )
        )
        fig2.update_layout(
            xaxis_title='CLIENTE',
            yaxis_title='TAXA DE APROVEITAMENTO (%)',
            coloraxis_colorbar=dict(title='APROVEITAMENTO (%)'),
            height=chart_height,
            template="plotly",
            margin=dict(l=60, r=60, t=30, b=60),
            xaxis=dict(tickangle=-45),
            yaxis=dict(range=[0, min(150, df_graph2['Aproveitamento'].max() * 1.1)])
        )
        st.plotly_chart(fig2, use_container_width=True)
        
        with st.expander("VER DETALHES DO CÁLCULO DESTE GRÁFICO"):
            st.markdown("""
            **DETALHAMENTO DO CÁLCULO:**
            - **FILTRAGEM:** Considera apenas os CLIENTES com oportunidades (IMPORTAÇÃO + EXPORTAÇÃO + CABOTAGEM > 0).
            - **AGRUPAMENTO:** Soma dos valores de IMPORTAÇÃO, EXPORTAÇÃO, CABOTAGEM e REALIZADO SYSTRACKER por CLIENTE.
            - **TOTAL DE OPORTUNIDADES:** Soma das três categorias (IMPORTAÇÃO + EXPORTAÇÃO + CABOTAGEM).
            - **APROVEITAMENTO:** Calculado como (REALIZADO SYSTRACKER / TOTAL DE OPORTUNIDADES) * 100.
            """)
        
        media_aproveitamento = df_graph2['Aproveitamento'].mean()
        melhor_cliente = df_graph2.iloc[0]['Cliente']
        melhor_aproveitamento = df_graph2.iloc[0]['Aproveitamento']
        
        st.markdown(f"""
        <div style='background-color:{COLORS['background']}; padding:10px; border-radius:5px; margin-top:10px;'>
            <h5 style='margin-top:0'>📊 INSIGHTS - APROVEITAMENTO</h5>
            <ul>
                <li>A TAXA MÉDIA DE APROVEITAMENTO DE OPORTUNIDADES É DE {media_aproveitamento:.1f}%</li>
                <li>O CLIENTE COM MELHOR APROVEITAMENTO É <b>{melhor_cliente}</b> COM {melhor_aproveitamento:.1f}%</li>
                <li>{"A MAIORIA DOS CLIENTES ESTÁ ABAIXO DA META MÍNIMA DE 50%" if media_aproveitamento < 50 else "A MAIORIA DOS CLIENTES ATINGE PELO MENOS A META MÍNIMA DE 50%"} </li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("SEM DADOS DE OPORTUNIDADES DISPONÍVEIS PARA OS FILTROS SELECIONADOS.")
        
# =============================================================================
# Tabela de Dados Detalhados
# =============================================================================
if show_detailed_table and not filtered_df.empty:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.markdown("<h3 class='section-title'>DADOS DETALHADOS</h3>", unsafe_allow_html=True)
    if 'MÊS' in filtered_df.columns:
        detailed_df = filtered_df.sort_values(['Cliente', 'MÊS'])
    else:
        detailed_df = filtered_df.sort_values(['Cliente'])
    detailed_df['MÊS_NOME'] = detailed_df['MÊS'].map(meses_map)
    detailed_df = detailed_df[[ 
        'Cliente', 'MÊS', 'MÊS_NOME', 'BUDGET', 'Importação', 'Exportação', 'Cabotagem',
        'Target Acumulado', 'Quantidade_iTRACKER', 'Gap de Realização'
    ]]
    detailed_df.columns = [
        'CLIENTE', 'MÊS (NÚM)', 'MÊS', 'BUDGET', 'IMPORTAÇÃO', 'EXPORTAÇÃO',
        'CABOTAGEM', 'TARGET ACUMULADO', 'REALIZADO (SYSTRACKER)', 'GAP DE REALIZAÇÃO'
    ]
    detailed_df = detailed_df.sort_values(['CLIENTE', 'MÊS (NÚM)'])
    cols = st.columns([3, 1])
    with cols[0]:
        search_term = st.text_input("BUSCAR CLIENTE", "")
    with cols[1]:
        sort_by = st.selectbox(
            "ORDENAR POR",
            options=["CLIENTE", "MÊS", "BUDGET", "REALIZADO (SYSTRACKER)", "GAP DE REALIZAÇÃO"],
            index=0
        )
    if search_term:
        detailed_df = detailed_df[detailed_df['CLIENTE'].str.contains(search_term.upper(), case=False)]
    if sort_by == "CLIENTE":
        detailed_df = detailed_df.sort_values(['CLIENTE', 'MÊS (NÚM)'])
    elif sort_by == "MÊS":
        detailed_df = detailed_df.sort_values(['MÊS (NÚM)', 'CLIENTE'])
    elif sort_by == "BUDGET":
        detailed_df = detailed_df.sort_values('BUDGET', ascending=False)
    elif sort_by == "REALIZADO (SYSTRACKER)":
        detailed_df = detailed_df.sort_values('REALIZADO (SYSTRACKER)', ascending=False)
    elif sort_by == "GAP DE REALIZAÇÃO":
        detailed_df = detailed_df.sort_values('GAP DE REALIZAÇÃO', ascending=False)
    detailed_df['REALIZADO (SYSTRACKER)'] = detailed_df['REALIZADO (SYSTRACKER)'].apply(lambda x: f'{x:.0f}')
    detailed_df['GAP DE REALIZAÇÃO'] = detailed_df['GAP DE REALIZAÇÃO'].apply(lambda x: f'{x:.1f}')
    st.dataframe(
        detailed_df,
        column_config={
            "CLIENTE": st.column_config.TextColumn("CLIENTE"),
            "MÊS": st.column_config.TextColumn("MÊS"),
            "BUDGET": st.column_config.NumberColumn("BUDGET", format="%d"),
            "IMPORTAÇÃO": st.column_config.NumberColumn("IMPORTAÇÃO", format="%d"),
            "EXPORTAÇÃO": st.column_config.NumberColumn("EXPORTAÇÃO", format="%d"),
            "CABOTAGEM": st.column_config.NumberColumn("CABOTAGEM", format="%d"),
            "TARGET ACUMULADO": st.column_config.NumberColumn("TARGET ACUMULADO", format="%d"),
            "REALIZADO (SYSTRACKER)": st.column_config.TextColumn("REALIZADO (SYSTRACKER)"),
            "GAP DE REALIZAÇÃO": st.column_config.TextColumn("GAP DE REALIZAÇÃO"),
        },
        use_container_width=True,
        height=500,
        hide_index=True
    )
    csv = detailed_df.to_csv(index=False)
    excel_buffer = io.BytesIO()
    detailed_df.to_excel(excel_buffer, index=False, engine='openpyxl')
    excel_data = excel_buffer.getvalue()
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            "📥 BAIXAR CSV",
            csv,
            "dados_detalhados.csv",
            "text/csv",
            key='download-csv'
        )
    with col_dl2:
        st.download_button(
            "📥 BAIXAR EXCEL",
            excel_data,
            "dados_detalhados.xlsx",
            "application/vnd.ms-excel",
            key='download-excel'
        )
    st.markdown("</div>", unsafe_allow_html=True)

# =============================================================================
# CONCLUSÕES E RECOMENDAÇÕES AUTOMÁTICAS
# =============================================================================
if not filtered_df.empty:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.markdown("<h3 class='section-title'>CONCLUSÕES E RECOMENDAÇÕES</h3>", unsafe_allow_html=True)
    total_budget = filtered_df['BUDGET'].sum()
    total_realizado = filtered_df['Quantidade_iTRACKER'].sum()
    performance_geral = (total_realizado / total_budget) * 100 if total_budget > 0 else 0
    total_oportunidades = filtered_df['Importação'].sum() + filtered_df['Exportação'].sum() + filtered_df['Cabotagem'].sum()
    aproveitamento_geral = (total_realizado / total_oportunidades) * 100 if total_oportunidades > 0 else 0
    clientes_prioritarios = filtered_df.groupby('Cliente').agg({
        'BUDGET': 'sum',
        'Quantidade_iTRACKER': 'sum'
    })
    clientes_prioritarios['Performance'] = (clientes_prioritarios['Quantidade_iTRACKER'] / clientes_prioritarios['BUDGET']) * 100
    clientes_prioritarios = clientes_prioritarios.sort_values(['BUDGET', 'Performance'])
    clientes_prioritarios = clientes_prioritarios[
        (clientes_prioritarios['BUDGET'] > clientes_prioritarios['BUDGET'].median()) &
        (clientes_prioritarios['Performance'] < 70) &
        (clientes_prioritarios['Performance'] > 0)
    ]
    top_prioritarios = clientes_prioritarios.head(3)
    st.markdown(f"""
    <div style='background-color:{COLORS['background']}; padding:15px; border-radius:8px;'>
        <h4 style='margin-top:0'>📈 ANÁLISE DE PERFORMANCE</h4>
        <p>Com base nos dados analisados, a performance geral está em <b>{format_percent(performance_geral)}</b> do budget projetado, com um aproveitamento de oportunidades de <b>{format_percent(aproveitamento_geral)}</b>.</p>
        <h4>🎯 RECOMENDAÇÕES</h4>
        <ol>
    """, unsafe_allow_html=True)
    if performance_geral < 70:
        st.markdown("""
            <li><b>ATENÇÃO IMEDIATA:</b> A performance geral está abaixo da meta aceitável de 70%. É recomendado revisar a estratégia comercial para aumentar o número de CONTAINERS movimentados.</li>
        """, unsafe_allow_html=True)
    elif performance_geral < 100:
        st.markdown("""
            <li><b>OPORTUNIDADE DE MELHORIA:</b> A performance está na zona intermediária. Há espaço para otimização das operações comerciais para atingir plenamente as metas do BUDGET.</li>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
            <li><b>MANTER ESTRATÉGIA:</b> A performance geral está atingindo ou superando o BUDGET. Recomenda-se manter a estratégia atual e possivelmente rever o BUDGET para metas mais ambiciosas.</li>
        """, unsafe_allow_html=True)
    if aproveitamento_geral < 50:
        st.markdown("""
            <li><b>MELHORAR APROVEITAMENTO:</b> A taxa de aproveitamento de oportunidades está baixa. Reveja os processos de prospecção e conversão.</li>
        """, unsafe_allow_html=True)
    elif aproveitamento_geral < 70:
        st.markdown("""
            <li><b>APRIMORAR CONVERSÃO:</b> A taxa de aproveitamento está em nível intermediário. Considere treinamentos e estratégias de follow-up.</li>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
            <li><b>EXCELENTE APROVEITAMENTO:</b> A taxa de aproveitamento está alta. Mantenha os processos atuais e explore novas oportunidades de mercado.</li>
        """, unsafe_allow_html=True)
    if not top_prioritarios.empty:
        st.markdown("<h4>CLIENTES PRIORITÁRIOS PARA AÇÕES</h4>", unsafe_allow_html=True)
        st.markdown("<ul>", unsafe_allow_html=True)
        for idx, row in top_prioritarios.iterrows():
            st.markdown(f"<li><b>{idx}</b>: Performance de {row['Performance']:.1f}% com budget de {format_number(row['BUDGET'])}</li>", unsafe_allow_html=True)
        st.markdown("</ul>", unsafe_allow_html=True)
    st.markdown("</ol></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# =============================================================================
# FOOTER PERSONALIZADO
# =============================================================================
st.markdown(f"""
<div class="custom-footer">
    <span>📅 ATUALIZADO EM: {current_date}</span> | 
    <span>📧 EMAIL: COMERCIAL@EMPRESA.COM</span> | 
    <span>📞 TELEFONE: (21) 99999-9999</span>
</div>
""", unsafe_allow_html=True)

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io, os
from datetime import datetime
import numpy as np

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

# CSS personalizado (com fonte Inter, animação, melhorias no layout, KPI uniformizados e footer com ícones)
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
    .main-header {{
        background-color: var(--secondaryBackgroundColor, #ffffff);
        padding: 20px 30px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }}
    /* Cabeçalho aprimorado: conteúdo centralizado mesmo com sidebar recolhida */
    .header-content {{
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 20px;
    }}
    .header-title {{
        text-align: center;
        flex: 1;
    }}
    .header-title h1 {{
        font-size: 42px;
        font-weight: 700;
        margin: 0;
    }}
    .header-title p {{
        font-size: 18px;
        color: gray;
        margin-top: 5px;
    }}
    .section {{
        background-color: var(--secondaryBackgroundColor, #ffffff);
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.08);
        margin-bottom: 20px;
    }}
    /* KPI uniformizados com maior padding e borda lateral para destaque */
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
    /* Animação de fadeIn */
    @keyframes fadeIn {{
      from {{ opacity: 0; }}
      to {{ opacity: 1; }}
    }}
    /* Footer aprimorado com ícones */
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

# Funções utilitárias
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

def download_file_from_gdrive():
    try:
        credentials_file = "C:/Users/leonardo.fragoso/Desktop/Projetos/dash-burgetXLogComexXComercial/service_account.json"
        if not os.path.exists(credentials_file):
            st.sidebar.error(f"Arquivo de credenciais não encontrado em: {credentials_file}")
            return None
        credentials = service_account.Credentials.from_service_account_file(
            credentials_file, 
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

# Carregar os dados do Google Sheets
df = download_file_from_gdrive()
if df is None:
    st.error("Não foi possível carregar os dados do Google Sheets.")
    st.stop()

# Remover registros sem nome de cliente para evitar "undefined"
df = df[df['Cliente'].notna() & (df['Cliente'] != "undefined")]

# Converter colunas numéricas
numeric_cols = ['MÊS', 'BUDGET', 'Importação', 'Exportação', 'Cabotagem', 'Quantidade_iTRACKER']
for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# Adicionar colunas calculadas
df['Total_Oportunidades'] = df['Importação'] + df['Exportação'] + df['Cabotagem']
df['Aproveitamento_Oportunidades'] = (df['Quantidade_iTRACKER'] / df['Total_Oportunidades']) * 100
df['Aproveitamento_Oportunidades'] = df['Aproveitamento_Oportunidades'].fillna(0)
df['Performance_Budget'] = (df['Quantidade_iTRACKER'] / df['BUDGET']) * 100
df['Performance_Budget'] = df['Performance_Budget'].fillna(0)
df['Gap_Budget'] = df['Quantidade_iTRACKER'] - df['BUDGET']

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
# Removemos a opção de "Análise Mensal" (não será exibida)
chart_height = st.sidebar.slider("Altura dos gráficos", 400, 800, 500, 50)

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
    st.markdown(f"<div style='background-color:#E3F2FD;padding:10px;border-radius:5px;margin-bottom:20px;'>"
                f"<b>Filtros ativos:</b> {' | '.join(filtros_ativos)}</div>", unsafe_allow_html=True)

# Cabeçalho Principal com ícone e título centralizado
st.markdown("<div class='main-header'>", unsafe_allow_html=True)
st.markdown("""
    <div class="header-content">
        <img src="https://cdn-icons-png.flaticon.com/512/2620/2620253.png" width="80"/>
        <div class="header-title">
            <h1>Dashboard Comercial</h1>
            <p>Análise Comparativa: <b>Budget</b> vs <b>Logcomex</b> vs <b>Desempenho Comercial</b></p>
        </div>
    </div>
""", unsafe_allow_html=True)
current_date = datetime.now().strftime("%d de %B de %Y")
st.markdown(f"<p style='text-align:right; color:{COLORS['text']}; font-weight:bold;'>Atualizado em: {current_date}</p>", unsafe_allow_html=True)
st.markdown("<hr style='margin-top:20px; margin-bottom:20px; border:none; border-bottom:2px solid #f0f0f0;'>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# Seção de KPIs com cartões aprimorados e ícones
st.markdown("<div class='section'>", unsafe_allow_html=True)
st.markdown("<h3 class='section-title'>Visão Geral</h3>", unsafe_allow_html=True)
col1, col2, col3, col4 = st.columns(4)
total_budget = filtered_df['BUDGET'].sum()
with col1:
    st.markdown(f"""
    <div class='kpi-card'>
        <p class='kpi-title'>💰 TOTAL BUDGET</p>
        <p class='kpi-value'>{format_number(total_budget)}</p>
    </div>
    """, unsafe_allow_html=True)
total_oportunidades = filtered_df['Total_Oportunidades'].sum()
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
        <p class='kpi-title'>🚚 TOTAL REALIZADO</p>
        <p class='kpi-value'>{format_number(total_itracker)}</p>
    </div>
    """, unsafe_allow_html=True)
with col4:
    st.markdown(f"""
    <div class='kpi-card'>
        <p class='kpi-title'>🎯 PERFORMANCE VS BUDGET</p>
        <p class='kpi-value'>{format_percent((total_itracker / total_budget) * 100 if total_budget > 0 else 0)}</p>
    </div>
    """, unsafe_allow_html=True)
if not filtered_df.empty:
    opp_aproveitadas = (filtered_df['Quantidade_iTRACKER'].sum() / filtered_df['Total_Oportunidades'].sum() * 100
                         if filtered_df['Total_Oportunidades'].sum() > 0 else 0)
    if 'MÊS' in filtered_df.columns:
        meses = filtered_df['MÊS'].unique()
        if len(meses) > 1:
            trend_df = filtered_df.groupby('MÊS').agg({
                'Quantidade_iTRACKER': 'sum',
                'BUDGET': 'sum'
            }).reset_index()
            trend_df['Performance'] = (trend_df['Quantidade_iTRACKER'] / trend_df['BUDGET']) * 100
            trend_df = trend_df.sort_values('MÊS')
            if len(trend_df) >= 2:
                first_perf = trend_df['Performance'].iloc[0]
                last_perf = trend_df['Performance'].iloc[-1]
                perf_trend = last_perf - first_perf
                trend_icon = "↗️" if perf_trend > 0 else "↘️" if perf_trend < 0 else "➡️"
                trend_color = COLORS['success'] if perf_trend > 0 else COLORS['danger'] if perf_trend < 0 else COLORS['light_text']
                col_trend1, col_trend2 = st.columns(2)
                with col_trend1:
                    st.markdown(f"<div style='padding:10px; border-radius:5px; background-color:{COLORS['background']};'>"
                                f"<b>Taxa de Aproveitamento:</b> {format_percent(opp_aproveitadas)}</div>",
                                unsafe_allow_html=True)
                with col_trend2:
                    st.markdown(f"<div style='padding:10px; border-radius:5px; background-color:{COLORS['background']};'>"
                                f"<b>Tendência de Performance:</b> <span style='color:{trend_color}'>{trend_icon} {perf_trend:.1f}%</span></div>",
                                unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# =============================================================================
# Gráfico 1: Comparativo Budget vs Realizado por Categoria (Agrupado)
# =============================================================================
if not filtered_df.empty:
    st.markdown("<h4 class='sub-title'>Comparativo Budget vs Realizado por Categoria</h4>", unsafe_allow_html=True)
    # Filtrar para remover eventuais "undefined" já está feito ao carregar os dados
    clientes_top = filtered_df.groupby('Cliente', as_index=False)['BUDGET'].sum()\
                    .sort_values('BUDGET', ascending=False)['Cliente'].head(15)
    df_top = filtered_df[filtered_df['Cliente'].isin(clientes_top)]
    df_grouped = df_top.groupby('Cliente', as_index=False).agg({
        'BUDGET': 'sum',
        'Importação': 'sum',
        'Exportação': 'sum',
        'Cabotagem': 'sum'
    })
    # Ordenar os clientes pelo volume total
    df_grouped['Total'] = df_grouped[['BUDGET', 'Importação', 'Exportação', 'Cabotagem']].sum(axis=1)
    df_grouped = df_grouped.sort_values('Total', ascending=False)
    df_melted = df_grouped.melt(
        id_vars='Cliente',
        value_vars=['BUDGET', 'Importação', 'Exportação', 'Cabotagem'],
        var_name='Categoria',
        value_name='Quantidade'
    )
    # Adicionar coluna auxiliar para a legenda (tooltip)
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
        labels={'Quantidade': 'Qtd. de Containers'},
        custom_data=['Categoria_Label']
    )
    fig.update_traces(
        texttemplate='%{y:.0f}',
        textposition='outside',
        hovertemplate='<b>Cliente:</b> %{x}<br><b>Categoria:</b> %{customdata[0]}<br><b>Qtd:</b> %{y:.0f}<extra></extra>'
    )
    fig.update_layout(
        xaxis=dict(title='Cliente', tickangle=-30),
        yaxis=dict(title='Quantidade (Containers)', range=[0, df_melted['Quantidade'].max() * 1.1]),
        legend=dict(orientation='h', y=1.02, x=1, xanchor='right'),
        margin=dict(l=60, r=40, t=20, b=80),
        template='plotly_white',
        bargap=0.2,
        title_text=""
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Sem dados disponíveis para o gráfico de comparativo após aplicação dos filtros.")

# =============================================================================
# Gráfico 2: Aproveitamento de Oportunidades por Cliente
# =============================================================================
if not filtered_df.empty:
    opp_df = filtered_df[filtered_df['Total_Oportunidades'] > 0].copy()
    if not opp_df.empty:
        st.markdown("<h4 class='sub-title'>Aproveitamento de Oportunidades por Cliente</h4>", unsafe_allow_html=True)
        df_graph2 = opp_df.groupby('Cliente', as_index=False).agg({
            'Total_Oportunidades': 'sum',
            'Quantidade_iTRACKER': 'sum'
        })
        df_graph2['Aproveitamento'] = (df_graph2['Quantidade_iTRACKER'] / df_graph2['Total_Oportunidades']) * 100
        df_graph2 = df_graph2.sort_values('Aproveitamento', ascending=False)
        if len(df_graph2) > 15:
            df_graph2 = df_graph2.head(15)
        # Atualização do hover para utilizar nomes amigáveis
        fig2 = px.bar(
            df_graph2,
            x='Cliente',
            y='Aproveitamento',
            color='Aproveitamento',
            color_continuous_scale=px.colors.sequential.Blues,
            text_auto='.1f',
            labels={'Aproveitamento': 'Taxa de Aproveitamento (%)'},
            custom_data=['Total_Oportunidades', 'Quantidade_iTRACKER']
        )
        fig2.update_traces(
            texttemplate='%{y:.1f}%',
            textposition='outside',
            hovertemplate=(
                '<b>Cliente:</b> %{x}<br>'
                '<b>Taxa de Aproveitamento:</b> %{y:.1f}%<br>'
                '<b>Total Oportunidades:</b> %{customdata[0]:,.0f}<br>'
                '<b>Realizado:</b> %{customdata[1]:,.0f}<extra></extra>'
            )
        )
        # Limitar o eixo Y para melhor visualização (até 150%)
        fig2.update_layout(
            xaxis_title='Cliente',
            yaxis_title='Taxa de Aproveitamento (%)',
            coloraxis_colorbar=dict(title='Aproveitamento (%)'),
            height=chart_height,
            template="plotly",
            margin=dict(l=60, r=60, t=30, b=60),
            xaxis=dict(tickangle=-45),
            yaxis=dict(range=[0, min(150, df_graph2['Aproveitamento'].max() * 1.1)])
        )
        st.plotly_chart(fig2, use_container_width=True)
        
        media_aproveitamento = df_graph2['Aproveitamento'].mean()
        melhor_cliente = df_graph2.iloc[0]['Cliente']
        melhor_aproveitamento = df_graph2.iloc[0]['Aproveitamento']
        
        st.markdown(f"""
        <div style='background-color:{COLORS['background']}; padding:10px; border-radius:5px; margin-top:10px;'>
            <h5 style='margin-top:0'>📊 Insights - Aproveitamento</h5>
            <ul>
                <li>A taxa média de aproveitamento de oportunidades é de {media_aproveitamento:.1f}%</li>
                <li>O cliente com melhor aproveitamento é <b>{melhor_cliente}</b> com {melhor_aproveitamento:.1f}%</li>
                <li>{"A maioria dos clientes está abaixo da meta mínima de 50%" if media_aproveitamento < 50 else "A maioria dos clientes atinge pelo menos a meta mínima de 50%"}</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("Sem dados de oportunidades disponíveis para os filtros selecionados.")

# =============================================================================
# Gráfico 3: Performance vs Budget por Cliente
# =============================================================================
if not filtered_df.empty:
    budget_df = filtered_df[filtered_df['BUDGET'] > 0].copy()
    if not budget_df.empty:
        st.markdown("<h4 class='sub-title'>Performance vs Budget por Cliente</h4>", unsafe_allow_html=True)
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
            text="Crítico (<70%)",
            showarrow=False,
            font=dict(color=COLORS['danger']),
            xanchor="center",
            yanchor="top"
        )
        fig3.add_annotation(
            x=85,
            y=len(df_graph3)-1,
            text="Atenção (70-100%)",
            showarrow=False,
            font=dict(color=COLORS['warning']),
            xanchor="center",
            yanchor="top"
        )
        fig3.add_annotation(
            x=min(150, df_graph3['Performance'].max() * 0.9),
            y=len(df_graph3)-1,
            text="Meta Atingida (>100%)",
            showarrow=False,
            font=dict(color=COLORS['success']),
            xanchor="center",
            yanchor="top"
        )
        fig3.update_traces(textposition='inside')
        fig3.update_layout(
            xaxis_title='Performance (%)',
            yaxis_title='Cliente',
            height=chart_height,
            template="plotly",
            margin=dict(l=60, r=30, t=30, b=40),
            xaxis=dict(range=[0, max(200, df_graph3['Performance'].max() * 1.1)])
        )
        st.plotly_chart(fig3, use_container_width=True)
        
        total_clientes = len(df_graph3)
        clientes_acima_meta = len(df_graph3[df_graph3['Performance'] >= 100])
        clientes_atencao = len(df_graph3[(df_graph3['Performance'] < 100) & (df_graph3['Performance'] >= 70)])
        clientes_critico = len(df_graph3[df_graph3['Performance'] < 70])
        
        st.markdown(f"""
        <div style='background-color:{COLORS['background']}; padding:10px; border-radius:5px; margin-top:10px;'>
            <h5 style='margin-top:0'>📊 Insights - Performance</h5>
            <ul>
                <li>Dos {total_clientes} clientes analisados:</li>
                <li><span style='color:{COLORS["success"]};'>✓ {clientes_acima_meta} clientes ({clientes_acima_meta/total_clientes*100:.1f}%) atingiram ou superaram a meta</span></li>
                <li><span style='color:{COLORS["warning"]};'>⚠️ {clientes_atencao} clientes ({clientes_atencao/total_clientes*100:.1f}%) estão em zona de atenção (70-99%)</span></li>
                <li><span style='color:{COLORS["danger"]};'>❌ {clientes_critico} clientes ({clientes_critico/total_clientes*100:.1f}%) estão em situação crítica (<70%)</span></li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("Sem dados de budget disponíveis para os filtros selecionados.")
st.markdown("</div>", unsafe_allow_html=True)

# =============================================================================
# Tabela de Dados Detalhados
# =============================================================================
if show_detailed_table and not filtered_df.empty:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.markdown("<h3 class='section-title'>Dados Detalhados</h3>", unsafe_allow_html=True)
    detailed_df = filtered_df.sort_values(['Cliente', 'MÊS'])
    detailed_df['Mês_Nome'] = detailed_df['MÊS'].map(meses_map)
    detailed_df = detailed_df[[ 'Cliente', 'MÊS', 'Mês_Nome', 'BUDGET', 'Importação', 'Exportação', 'Cabotagem',
                                'Total_Oportunidades', 'Quantidade_iTRACKER', 'Performance_Budget',
                                'Aproveitamento_Oportunidades', 'Gap_Budget' ]]
    detailed_df.columns = [ 'Cliente', 'Mês (Núm)', 'Mês', 'Budget', 'Importação', 'Exportação',
                              'Cabotagem', 'Oportunidades Totais', 'Realizado (iTracker)',
                              'Performance (%)', 'Aproveitamento (%)', 'Gap Budget' ]
    cols = st.columns([3, 1])
    with cols[0]:
        search_term = st.text_input("Buscar cliente", "")
    with cols[1]:
        sort_by = st.selectbox(
            "Ordenar por",
            options=["Cliente", "Mês", "Budget", "Performance (%)", "Aproveitamento (%)"],
            index=0
        )
    if search_term:
        detailed_df = detailed_df[detailed_df['Cliente'].str.contains(search_term, case=False)]
    if sort_by == "Cliente":
        detailed_df = detailed_df.sort_values(['Cliente', 'Mês (Núm)'])
    elif sort_by == "Mês":
        detailed_df = detailed_df.sort_values(['Mês (Núm)', 'Cliente'])
    elif sort_by == "Budget":
        detailed_df = detailed_df.sort_values('Budget', ascending=False)
    elif sort_by == "Performance (%)":
        detailed_df = detailed_df.sort_values('Performance (%)', ascending=False)
    elif sort_by == "Aproveitamento (%)":
        detailed_df = detailed_df.sort_values('Aproveitamento (%)', ascending=False)
    detailed_df['Performance (%)'] = detailed_df['Performance (%)'].apply(lambda x: f'{x:.1f}%')
    detailed_df['Aproveitamento (%)'] = detailed_df['Aproveitamento (%)'].apply(lambda x: f'{x:.1f}%')
    st.dataframe(
        detailed_df,
        column_config={
            "Cliente": st.column_config.TextColumn("Cliente"),
            "Mês": st.column_config.TextColumn("Mês"),
            "Budget": st.column_config.NumberColumn("Budget", format="%d"),
            "Importação": st.column_config.NumberColumn("Importação", format="%d"),
            "Exportação": st.column_config.NumberColumn("Exportação", format="%d"),
            "Cabotagem": st.column_config.NumberColumn("Cabotagem", format="%d"),
            "Oportunidades Totais": st.column_config.NumberColumn("Oportunidades Totais", format="%d"),
            "Realizado (iTracker)": st.column_config.NumberColumn("Realizado (iTracker)", format="%d"),
            "Performance (%)": st.column_config.TextColumn("Performance (%)"),
            "Aproveitamento (%)": st.column_config.TextColumn("Aproveitamento (%)"),
            "Gap Budget": st.column_config.NumberColumn("Gap Budget", format="%d"),
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
            "📥 Baixar CSV",
            csv,
            "dados_detalhados.csv",
            "text/csv",
            key='download-csv'
        )
    with col_dl2:
        st.download_button(
            "📥 Baixar Excel",
            excel_data,
            "dados_detalhados.xlsx",
            "application/vnd.ms-excel",
            key='download-excel'
        )
    st.markdown("</div>", unsafe_allow_html=True)

# =============================================================================
# Conclusões e Recomendações Automáticas
# =============================================================================
if not filtered_df.empty:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.markdown("<h3 class='section-title'>Conclusões e Recomendações</h3>", unsafe_allow_html=True)
    total_budget = filtered_df['BUDGET'].sum()
    total_realizado = filtered_df['Quantidade_iTRACKER'].sum()
    performance_geral = (total_realizado / total_budget) * 100 if total_budget > 0 else 0
    total_oportunidades = filtered_df['Total_Oportunidades'].sum()
    aproveitamento_geral = (total_realizado / total_oportunidades) * 100 if total_oportunidades > 0 else 0
    clientes_prioritarios = filtered_df.groupby('Cliente').agg({
        'BUDGET': 'sum',
        'Quantidade_iTRACKER': 'sum'
    })
    clientes_prioritarios['Performance'] = (clientes_prioritarios['Quantidade_iTRACKER'] / clientes_prioritarios['BUDGET']) * 100
    clientes_prioritarios = clientes_prioritarios.sort_values(['BUDGET', 'Performance'])
    clientes_prioritarios = clientes_prioritarios[(clientes_prioritarios['BUDGET'] > clientes_prioritarios['BUDGET'].median()) &
                                                 (clientes_prioritarios['Performance'] < 70)]
    top_prioritarios = clientes_prioritarios.head(3)
    st.markdown(f"""
    <div style='background-color:{COLORS['background']}; padding:15px; border-radius:8px;'>
        <h4 style='margin-top:0'>📈 Análise de Performance</h4>
        <p>Com base nos dados analisados, a performance geral está em <b>{format_percent(performance_geral)}</b> do budget projetado, com um aproveitamento de oportunidades de <b>{format_percent(aproveitamento_geral)}</b>.</p>
        <h4>🎯 Recomendações</h4>
        <ol>
    """, unsafe_allow_html=True)
    if performance_geral < 70:
        st.markdown("""
            <li><b>Atenção Imediata:</b> A performance geral está abaixo da meta aceitável de 70%. É recomendado revisar a estratégia comercial para aumentar o número de containers movimentados.</li>
        """, unsafe_allow_html=True)
    elif performance_geral < 100:
        st.markdown("""
            <li><b>Oportunidade de Melhoria:</b> A performance está na zona intermediária. Há espaço para otimização das operações comerciais para atingir plenamente as metas do budget.</li>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
            <li><b>Manter Estratégia:</b> A performance geral está atingindo ou superando o budget. Recomenda-se manter a estratégia atual e possivelmente rever o budget para metas mais ambiciosas.</li>
        """, unsafe_allow_html=True)
    if aproveitamento_geral < 50:
        st.markdown("""
            <li><b>Melhorar Aproveitamento:</b> A taxa de aproveitamento de oportunidades está baixa. Revise os processos de prospecção e conversão.</li>
        """, unsafe_allow_html=True)
    elif aproveitamento_geral < 70:
        st.markdown("""
            <li><b>Aprimorar Conversão:</b> A taxa de aproveitamento está em nível intermediário. Considere treinamentos e estratégias de follow-up.</li>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
            <li><b>Excelente Aproveitamento:</b> A taxa de aproveitamento está alta. Mantenha os processos atuais e explore novas oportunidades de mercado.</li>
        """, unsafe_allow_html=True)
    if not top_prioritarios.empty:
        st.markdown("<h4>Clientes Prioritários para Ações</h4>", unsafe_allow_html=True)
        st.markdown("<ul>", unsafe_allow_html=True)
        for idx, row in top_prioritarios.iterrows():
            st.markdown(f"<li><b>{idx}</b>: Performance de {row['Performance']:.1f}% com budget de {format_number(row['BUDGET'])}</li>", unsafe_allow_html=True)
        st.markdown("</ul>", unsafe_allow_html=True)
    st.markdown("</ol></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# =============================================================================
# Footer Personalizado
# =============================================================================
st.markdown(f"""
<div class="custom-footer">
    <span>📅 Atualizado em: {current_date}</span> | 
    <span>📧 Email: comercial@empresa.com</span> | 
    <span>📞 Telefone: (21) 99999-9999</span>
</div>
""", unsafe_allow_html=True)

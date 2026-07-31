import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os, glob

# ---------------------------------------------------------
# Page Configuration & Custom CSS Theme
# ---------------------------------------------------------
st.set_page_config(
    page_title="SUDAMERICANA - Dashboard de Margen Operativo",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom High-End Dashboard Styling
st.markdown("""
    <style>
    /* Dark Modern Theme Palette */
    .stApp {
        background-color: #0b0f19;
        color: #f1f5f9;
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
    }
    
    /* Metric Cards */
    .metric-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 14px;
        padding: 16px 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
        transition: transform 0.2s, border-color 0.2s;
    }
    .metric-card:hover {
        border-color: #38bdf8;
        transform: translateY(-2px);
    }
    .metric-title {
        color: #94a3b8;
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }
    .metric-value {
        color: #38bdf8;
        font-size: 1.65rem;
        font-weight: 700;
        margin-top: 6px;
    }
    .metric-sub {
        color: #64748b;
        font-size: 0.76rem;
        margin-top: 2px;
    }
    
    .filter-bar-title {
        font-size: 0.9rem;
        font-weight: 600;
        color: #38bdf8;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
    }

    /* Summary Chips Badge Styling */
    .summary-chip-container {
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
        margin-top: 16px;
        margin-bottom: 16px;
    }
    
    .chip {
        background-color: #0f172a;
        border-radius: 24px;
        padding: 10px 20px;
        font-size: 0.98rem;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 8px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.3);
        border-width: 1.5px !important;
    }

    .chip-blue {
        border: 1px solid #38bdf8;
        color: #38bdf8;
    }

    .chip-red {
        border: 1px solid #f43f5e;
        color: #f43f5e;
    }

    .chip-green {
        border: 1px solid #10b981;
        color: #10b981;
    }

    .chip-purple {
        border: 1px solid #a855f7;
        color: #c084fc;
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #0f172a;
        padding: 6px;
        border-radius: 12px;
        border: 1px solid #334155;
    }

    .stTabs [data-baseweb="tab"] {
        height: 42px;
        border-radius: 8px;
        color: #94a3b8;
        font-weight: 500;
        font-size: 0.9rem;
        padding: 0 16px;
    }

    .stTabs [aria-selected="true"] {
        background-color: #38bdf8 !important;
        color: #0f172a !important;
        font-weight: 700 !important;
    }

    /* Input Controls Custom Styling */
    div[data-baseweb="select"] > div {
        background-color: #0f172a !important;
        border-color: #334155 !important;
        border-radius: 8px !important;
    }
    
    .stMultiSelect div[data-baseweb="select"] span {
        background-color: #334155 !important;
        color: #f8fafc !important;
        border-radius: 6px !important;
    }
    </style>
""", unsafe_allow_html=True)

# Month and Day dictionaries in Spanish
MESES_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Setiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

MESES_ABR_ES = {
    1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
    7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Oct', 11: 'Nov', 12: 'Dic'
}

DIAS_ES = {
    'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
    'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
}

# Helper to load and parse any Sistrans Excel layout (Python 3.14 Linux Compatible)
def load_sistrans_file(file_source):
    df_raw = pd.read_excel(file_source, header=None)
    
    # Locate header row containing PLACA or RUTA or GASTOS or CHOFER
    h_idx = 7
    for idx, row in df_raw.iloc[:15].iterrows():
        row_str = ' '.join([str(val) for val in row.values if pd.notnull(val) and str(val) != 'nan']).upper()
        if 'PLACA' in row_str or 'GASTOS' in row_str or 'CHOFER' in row_str:
            h_idx = idx
            break
            
    row1 = [str(v).strip() if pd.notnull(v) and str(v) != 'nan' else '' for v in df_raw.iloc[h_idx].values]
    row2 = [str(v).strip() if pd.notnull(v) and str(v) != 'nan' else '' for v in df_raw.iloc[h_idx+1].values] if h_idx+1 < len(df_raw) else ['']*len(row1)
    
    col_names = []
    for idx_c, (h1, h2) in enumerate(zip(row1, row2)):
        name = f'{h1} {h2}'.strip().upper()
        name = ' '.join(name.split())
        if not name:
            name = f'UNNAMED_{idx_c}'
        col_names.append(name)
        
    ser_col_idx = 0
    for c in range(min(5, len(col_names))):
        vals = pd.to_numeric(df_raw[c], errors='coerce')
        if (vals > 1000).sum() >= 3:
            ser_col_idx = c
            break
            
    ser_vals = pd.to_numeric(df_raw[ser_col_idx], errors='coerce')
    df_data = df_raw[ser_vals > 1000].copy()
    df_data.columns = col_names
    return df_data

# ---------------------------------------------------------
# Dynamic Data Loader (Direct Database Loading)
# ---------------------------------------------------------
@st.cache_data
def process_excel_files(file_mrg_path, file_gen_path):
    # 1. Load Margen File (Financial base)
    df_mrg_raw = load_sistrans_file(file_mrg_path)
    
    df_mrg = pd.DataFrame()
    for col in df_mrg_raw.columns:
        c_upper = col.upper()
        if 'SERVICIO IDA' in c_upper or 'NO. SERVICIO' in c_upper:
            df_mrg['servicio_ida'] = df_mrg_raw[col]
        elif 'FECHA SALIDA' in c_upper or ('FECHA' in c_upper and 'fecha_ida' not in df_mrg.columns):
            df_mrg['fecha_ida'] = df_mrg_raw[col]
        elif 'TIPO SERVICIO' in c_upper:
            df_mrg['tipo_servicio'] = df_mrg_raw[col]
        elif 'RUTA' in c_upper:
            df_mrg['ruta'] = df_mrg_raw[col]
        elif 'PLACA' in c_upper:
            df_mrg['placa'] = df_mrg_raw[col]
        elif 'OPERADOR' in c_upper:
            df_mrg['operador'] = df_mrg_raw[col]
        elif 'SERVICIO RETORNO' in c_upper:
            df_mrg['servicio_retorno'] = df_mrg_raw[col]
        elif 'FLETE IDA' in c_upper:
            df_mrg['flete_ida'] = df_mrg_raw[col]
        elif 'FLETE RETORNO' in c_upper:
            df_mrg['flete_retorno'] = df_mrg_raw[col]
        elif 'GASTOS IDA' in c_upper:
            df_mrg['gastos_ida'] = df_mrg_raw[col]
        elif 'GASTOS RETORNO' in c_upper:
            df_mrg['gastos_retorno'] = df_mrg_raw[col]
        elif 'MARGEN BRUTO' in c_upper or 'MARGEN' in c_upper:
            df_mrg['margen_bruto'] = df_mrg_raw[col]

    df_mrg['servicio_ida'] = pd.to_numeric(df_mrg['servicio_ida'], errors='coerce')
    df_mrg['servicio_retorno'] = pd.to_numeric(df_mrg['servicio_retorno'], errors='coerce').fillna(0)
    df_mrg['fecha_ida'] = pd.to_datetime(df_mrg['fecha_ida'], errors='coerce')

    num_cols = ['flete_ida', 'flete_retorno', 'gastos_ida', 'gastos_retorno', 'margen_bruto']
    for col in num_cols:
        if col in df_mrg.columns:
            df_mrg[col] = pd.to_numeric(df_mrg[col], errors='coerce').fillna(0)
        else:
            df_mrg[col] = 0.0

    # 2. Load General File for Client, Driver, and Real Departure Date
    df_gen_raw = load_sistrans_file(file_gen_path)
    
    df_gen = pd.DataFrame()
    for col in df_gen_raw.columns:
        c_upper = col.upper()
        if 'NO. SERVICIO' in c_upper or 'SERVICIO IDA' in c_upper:
            df_gen['id_servicio'] = df_gen_raw[col]
        elif 'FECHA SALIDA' in c_upper or ('FECHA' in c_upper and 'fecha_gen' not in df_gen.columns):
            df_gen['fecha_gen'] = df_gen_raw[col]
        elif 'OFICINA' in c_upper:
            df_gen['oficina'] = df_gen_raw[col]
        elif 'TIPO SERVICIO' in c_upper:
            df_gen['tipo_servicio'] = df_gen_raw[col]
        elif 'RUTA' in c_upper:
            df_gen['ruta'] = df_gen_raw[col]
        elif 'PLACA' in c_upper:
            df_gen['placa'] = df_gen_raw[col]
        elif 'OPERADOR' in c_upper:
            df_gen['operador'] = df_gen_raw[col]
        elif 'CHOFER' in c_upper:
            df_gen['chofer'] = df_gen_raw[col]
        elif 'CLIENTE' in c_upper:
            df_gen['cliente'] = df_gen_raw[col]
        elif 'PRODUCTO' in c_upper:
            df_gen['producto'] = df_gen_raw[col]
        elif 'FECHA FIN' in c_upper:
            df_gen['fecha_fin'] = df_gen_raw[col]
        elif 'ESTADO' in c_upper:
            df_gen['estado'] = df_gen_raw[col]
        elif 'IMPORTE FLETE' in c_upper or 'FLETE' in c_upper:
            df_gen['flete_gen'] = df_gen_raw[col]
        elif 'OBS' in c_upper:
            df_gen['obs'] = df_gen_raw[col]

    df_gen['id_servicio'] = pd.to_numeric(df_gen['id_servicio'], errors='coerce')
    df_gen = df_gen.dropna(subset=['id_servicio'])
    df_gen['fecha_gen'] = pd.to_datetime(df_gen['fecha_gen'], errors='coerce')
    df_gen['flete_gen'] = pd.to_numeric(df_gen['flete_gen'], errors='coerce').fillna(0)

    df_gen['cliente'] = df_gen['cliente'].astype(str).str.strip().str.upper()
    df_gen['chofer'] = df_gen['chofer'].astype(str).str.strip().str.upper()

    # 3. Unroll IDA and RETORNO legs to count each service by its REAL departure date!
    leg_ida = df_mrg[['servicio_ida', 'fecha_ida', 'tipo_servicio', 'ruta', 'placa', 'operador', 
                       'flete_ida', 'gastos_ida']].copy()
    leg_ida.columns = ['id_servicio', 'fecha_mrg', 'tipo_servicio', 'ruta', 'placa', 'operador', 
                       'flete_mrg', 'gastos_total']
    leg_ida['tramo'] = 'IDA'

    leg_ret = df_mrg[df_mrg['servicio_retorno'] > 0][['servicio_retorno', 'fecha_ida', 'tipo_servicio', 'ruta', 'placa', 'operador', 
                                                       'flete_retorno', 'gastos_retorno']].copy()
    leg_ret.columns = ['id_servicio', 'fecha_mrg', 'tipo_servicio', 'ruta', 'placa', 'operador', 
                       'flete_mrg', 'gastos_total']
    leg_ret['tramo'] = 'RETORNO'

    df_services = pd.concat([leg_ida, leg_ret], ignore_index=True)

    # Merge operational details from General file by id_servicio
    df_services = pd.merge(df_services, df_gen[['id_servicio', 'fecha_gen', 'cliente', 'chofer', 'producto', 'estado', 'flete_gen']], 
                           on='id_servicio', how='left')

    # Effective departure date: prefer real operational date from General list, fallback to Margen date
    df_services['fecha_salida'] = df_services['fecha_gen'].fillna(pd.to_datetime(df_services['fecha_mrg']))
    df_services['cliente'] = df_services['cliente'].fillna('SIN CLIENTE')
    df_services['chofer'] = df_services['chofer'].fillna('SIN CHOFER')

    df_services['cliente_ida'] = df_services['cliente']
    df_services['cliente_retorno'] = np.where(df_services['tramo'] == 'RETORNO', df_services['cliente'], 'N/A')
    df_services['chofer_ida'] = df_services['chofer']
    df_services['servicio_ida'] = df_services['id_servicio']
    df_services['servicio_retorno'] = np.where(df_services['tramo'] == 'RETORNO', df_services['id_servicio'], 0)
    df_services['flete_ida'] = np.where(df_services['tramo'] == 'IDA', df_services['flete_mrg'], 0)
    df_services['flete_retorno'] = np.where(df_services['tramo'] == 'RETORNO', df_services['flete_mrg'], 0)
    df_services['gastos_ida'] = np.where(df_services['tramo'] == 'IDA', df_services['gastos_total'], 0)
    df_services['gastos_retorno'] = np.where(df_services['tramo'] == 'RETORNO', df_services['gastos_total'], 0)

    df_services['flete_total'] = df_services['flete_mrg']
    df_services['flete_gen_total'] = df_services['flete_gen']
    df_services['margen_calculado'] = df_services['flete_total'] - df_services['gastos_total']
    df_services['pct_margen'] = np.where(df_services['flete_total'] > 0, (df_services['margen_calculado'] / df_services['flete_total']) * 100, 0)
    df_services['cant_servicios'] = 1

    # Clean text columns
    df_services['tipo_servicio'] = df_services['tipo_servicio'].astype(str).str.strip().str.upper()
    df_services['ruta'] = df_services['ruta'].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
    df_services['placa'] = df_services['placa'].astype(str).str.strip().str.upper()
    df_services['operador'] = df_services['operador'].astype(str).str.strip().str.upper()

    # Spanish dates and weeks formatting
    df_services['mes_num'] = df_services['fecha_salida'].dt.month
    df_services['mes_nombre'] = df_services['fecha_salida'].dt.month.map(MESES_ES)
    df_services['mes_año'] = df_services['fecha_salida'].apply(lambda d: f"{MESES_ES[d.month]} {d.year}" if pd.notnull(d) else "N/A")
    df_services['dia_semana'] = df_services['fecha_salida'].dt.day_name().map(DIAS_ES)

    # ISO Week with friendly Spanish date range
    df_services['semana_num'] = df_services['fecha_salida'].dt.isocalendar().week
    df_services['semana_inicio'] = df_services['fecha_salida'].apply(lambda d: d - pd.Timedelta(days=d.weekday()) if pd.notnull(d) else d)
    df_services['semana_fin'] = df_services['semana_inicio'] + pd.Timedelta(days=6)

    def make_week_label(r):
        if pd.isnull(r['fecha_salida']):
            return 'N/A'
        sem = r['semana_num']
        ini_d = r['semana_inicio'].day
        ini_m = MESES_ABR_ES[r['semana_inicio'].month]
        fin_d = r['semana_fin'].day
        fin_m = MESES_ABR_ES[r['semana_fin'].month]
        return f"Semana {sem:02d} ({ini_d:02d} {ini_m} - {fin_d:02d} {fin_m})"

    df_services['semana_label'] = df_services.apply(make_week_label, axis=1)

    return df_services

# ---------------------------------------------------------
# Absolute Path Resolution & Multi-Week Dataset Discovery
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_available_weeks_map():
    data_dir = os.path.join(BASE_DIR, "data")
    weeks_map = {}
    
    if os.path.exists(data_dir):
        subdirs = sorted([d for d in os.listdir(data_dir) if os.path.isdir(os.path.join(data_dir, d))])
        for sdir in subdirs:
            folder_path = os.path.join(data_dir, sdir)
            xls_files = [f for f in os.listdir(folder_path) if f.endswith(('.xls', '.xlsx'))]
            if len(xls_files) >= 2:
                # Find margen and general files
                mrg_f = [f for f in xls_files if 'margen' in f.lower()]
                gen_f = [f for f in xls_files if 'margen' not in f.lower()]
                
                m_path = os.path.join(folder_path, mrg_f[0]) if mrg_f else os.path.join(folder_path, xls_files[0])
                g_path = os.path.join(folder_path, gen_f[0]) if gen_f else os.path.join(folder_path, xls_files[-1])
                
                label = sdir.replace('_', ' ').title()
                weeks_map[label] = (m_path, g_path)
                
    # Fallback to root files if data/ is empty
    if not weeks_map:
        root_mrg = os.path.join(BASE_DIR, "rptlistadoserviciosmargen_nuevo.xls")
        root_gen = os.path.join(BASE_DIR, "rptlistadoservicios_nuevo.xls")
        if os.path.exists(root_mrg) and os.path.exists(root_gen):
            weeks_map["Base Inicial (Corte Semana 30)"] = (root_mrg, root_gen)
            
    return weeks_map

weeks_map = get_available_weeks_map()
week_labels = list(weeks_map.keys())

# Sidebar Data Version Selection
st.sidebar.markdown("### 📁 Selección de Corte / Semana")
selected_week_label = st.sidebar.selectbox(
    "Seleccionar Versión de Datos:",
    options=week_labels,
    index=len(week_labels) - 1,
    help="Permite a la gerencia seleccionar la versión de datos a visualizar."
)

current_mrg_path, current_gen_path = weeks_map[selected_week_label]

# Process Current Selected Dataset
df_all = process_excel_files(current_mrg_path, current_gen_path)

# Determine Baseline Dataset for Comparison if previous week exists
baseline_mrg_path, baseline_gen_path = None, None
if len(week_labels) > 1:
    curr_idx = week_labels.index(selected_week_label)
    prev_idx = max(0, curr_idx - 1)
    if prev_idx != curr_idx:
        baseline_mrg_path, baseline_gen_path = weeks_map[week_labels[prev_idx]]

min_date = df_all['fecha_salida'].min().date()
max_date = df_all['fecha_salida'].max().date()

# ---------------------------------------------------------
# Top Main Header
# ---------------------------------------------------------
st.markdown(f"""
    <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px;">
        <div>
            <h1 style="font-size: 1.9rem; margin: 0; font-weight: 700; color: #f8fafc;">🚚 SUDAMERICANA - DASHBOARD DE MARGEN OPERATIVO</h1>
            <p style="color: #94a3b8; font-size: 0.9rem; margin-top: 4px;">Sistema Sistrans | Empresa: <b>Inversiones Comerciales Sudamericana S.R.L.</b> | <b>Corte: {selected_week_label}</b></p>
        </div>
    </div>
""", unsafe_allow_html=True)

# Controls bar
gcol1, gcol2 = st.columns([2, 2])
with gcol1:
    start_date, end_date = st.date_input(
        "📅 Rango de Fechas Global",
        value=[min_date, max_date],
        min_value=min_date,
        max_value=max_date
    )
with gcol2:
    flete_source = st.radio(
        "💵 Criterio de Flete para Clientes:",
        options=["Flete Liquidado (Reporte Margen)", "Flete Inicial (Listado General)"],
        horizontal=True,
        help="El Flete Liquidado muestra el ingreso cobrado real. El Flete Inicial muestra el registro operativo preliminar."
    )

# Global Filter Mask
df = df_all[
    (df_all['fecha_salida'].dt.date >= start_date) &
    (df_all['fecha_salida'].dt.date <= end_date)
].copy()

# ---------------------------------------------------------
# Top Executive KPI Cards
# ---------------------------------------------------------
c1, c2, c3, c4, c5 = st.columns(5)

total_flete = df['flete_total'].sum()
total_gastos = df['gastos_total'].sum()
total_margen = df['margen_calculado'].sum()
pct_margen_glob = (total_margen / total_flete * 100) if total_flete > 0 else 0
total_servicios = len(df)

with c1:
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">💰 Facturación (Flete)</div>
            <div class="metric-value">S/ {total_flete:,.2f}</div>
            <div class="metric-sub">Ingreso Total Bruto</div>
        </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">💸 Gastos Operativos</div>
            <div class="metric-value" style="color: #f43f5e;">S/ {total_gastos:,.2f}</div>
            <div class="metric-sub">Costo de Transporte</div>
        </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📈 Margen Bruto</div>
            <div class="metric-value" style="color: #10b981;">S/ {total_margen:,.2f}</div>
            <div class="metric-sub">Margen Operativo Bruto</div>
        </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">🎯 Margen %</div>
            <div class="metric-value" style="color: #f59e0b;">{pct_margen_glob:.1f}%</div>
            <div class="metric-sub">Retorno sobre Ingreso</div>
        </div>
    """, unsafe_allow_html=True)

with c5:
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📦 Total Servicios</div>
            <div class="metric-value" style="color: #a855f7;">{total_servicios}</div>
            <div class="metric-sub">Servicios Reales Ejecutados</div>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ---------------------------------------------------------
# Interactive Contextual Tabs
# ---------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Balance Mensual y Semanal", 
    "👥 Top de Clientes", 
    "🚛 Top de Placas (Vehículos)", 
    "🗺️ Rutas y Tipos de Servicio",
    "📋 Registro Detallado",
    "🔄 Auditoría de Regularizaciones"
])

# ---------------------------------------------------------
# TAB 1: Balance Mensual y Semanal
# ---------------------------------------------------------
with tab1:
    col_m1, col_m2 = st.columns(2)
    
    with col_m1:
        st.markdown('<div class="filter-bar-title">📅 1. Balance Mensual de Ingresos vs Gastos</div>', unsafe_allow_html=True)
        
        meses_opt = df.sort_values('fecha_salida')['mes_año'].unique().tolist()
        sel_meses = st.multiselect("🔍 Filtrar Meses en este gráfico:", options=meses_opt, default=meses_opt, key="tab1_meses")
        
        df_mes_filtered = df[df['mes_año'].isin(sel_meses)]
        
        df_mes = df_mes_filtered.groupby('mes_año', sort=False).agg(
            Flete_Total=('flete_total', 'sum'),
            Gastos_Total=('gastos_total', 'sum'),
            Margen_Bruto=('margen_calculado', 'sum'),
            Servicios=('cant_servicios', 'sum')
        ).reset_index()
        
        fig_mes = go.Figure()
        fig_mes.add_trace(go.Bar(
            x=df_mes['mes_año'], 
            y=df_mes['Flete_Total'], 
            name='Flete (Ingresos)', 
            marker_color='#38bdf8',
            hovertemplate="<b>%{x}</b><br>💰 Flete: <b>S/ %{y:,.2f}</b><extra></extra>"
        ))
        fig_mes.add_trace(go.Bar(
            x=df_mes['mes_año'], 
            y=df_mes['Gastos_Total'], 
            name='Gastos Operativos', 
            marker_color='#f43f5e',
            hovertemplate="<b>%{x}</b><br>💸 Gastos: <b>S/ %{y:,.2f}</b><extra></extra>"
        ))
        fig_mes.add_trace(go.Scatter(
            x=df_mes['mes_año'], 
            y=df_mes['Margen_Bruto'], 
            name='Margen Bruto', 
            mode='lines+markers+text',
            text=[f"S/ {v:,.0f}" for v in df_mes['Margen_Bruto']], 
            textposition="top center",
            line=dict(color='#10b981', width=3),
            hovertemplate="<b>%{x}</b><br>📈 Margen Bruto: <b>S/ %{y:,.2f}</b><extra></extra>"
        ))
        
        fig_mes.update_layout(
            barmode='group', 
            template='plotly_dark', 
            height=380, 
            margin=dict(l=10, r=10, t=30, b=10),
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_mes, use_container_width=True)

    with col_m2:
        st.markdown('<div class="filter-bar-title">📆 2. Evolución Semanal de Servicios & Margen</div>', unsafe_allow_html=True)
        
        semanas_df = df.sort_values('semana_num')[['semana_num', 'semana_label']].drop_duplicates()
        semanas_opt = semanas_df['semana_label'].tolist()
        sel_semanas = st.multiselect("🔍 Filtrar Semanas en este gráfico:", options=semanas_opt, default=semanas_opt, key="tab1_semanas")
        
        df_sem_filtered = df[df['semana_label'].isin(sel_semanas)]
        
        df_sem = df_sem_filtered.groupby(['semana_num', 'semana_label'], sort=True).agg(
            Servicios=('cant_servicios', 'sum'),
            Margen_Bruto=('margen_calculado', 'sum'),
            Flete_Total=('flete_total', 'sum')
        ).reset_index()

        fig_sem = go.Figure()
        
        fig_sem.add_trace(go.Bar(
            x=df_sem['semana_label'], 
            y=df_sem['Servicios'], 
            name='Cant. Servicios', 
            yaxis='y2', 
            marker_color='#a855f7', 
            opacity=0.75,
            customdata=np.stack((df_sem['Flete_Total'], df_sem['Margen_Bruto']), axis=-1),
            hovertemplate="<b>%{x}</b><br>📦 Servicios Realizados: <b>%{y} servicios</b><br>💰 Flete Total: <b>S/ %{customdata[0]:,.2f}</b><br>📈 Margen Bruto: <b>S/ %{customdata[1]:,.2f}</b><extra></extra>"
        ))
        
        fig_sem.add_trace(go.Scatter(
            x=df_sem['semana_label'], 
            y=df_sem['Margen_Bruto'], 
            name='Margen Bruto (S/)', 
            mode='lines+markers', 
            line=dict(color='#10b981', width=3),
            customdata=np.stack((df_sem['Flete_Total'], df_sem['Servicios']), axis=-1),
            hovertemplate="<b>%{x}</b><br>📈 Margen Bruto: <b>S/ %{y:,.2f}</b><br>💰 Flete Total: <b>S/ %{customdata[0]:,.2f}</b><br>📦 Servicios: <b>%{customdata[1]} servicios</b><extra></extra>"
        ))

        fig_sem.update_layout(
            template='plotly_dark', 
            height=380, 
            margin=dict(l=10, r=10, t=30, b=10),
            hovermode="x unified",
            yaxis=dict(title='Margen Bruto (S/)', showgrid=False, tickprefix="S/ "),
            yaxis2=dict(title='Cant. Servicios', overlaying='y', side='right', showgrid=False),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_sem, use_container_width=True)

    st.markdown("---")
    
    st.markdown("##### 📌 Resumen KPI de la Semana Seleccionada")
    if len(df_sem_filtered) > 0:
        latest_week_label = df_sem_filtered.sort_values('fecha_salida')['semana_label'].iloc[-1]
        df_curr_week = df_sem_filtered[df_sem_filtered['semana_label'] == latest_week_label]
        
        wc1, wc2, wc3, wc4 = st.columns(4)
        wc1.metric("Semana Activa", latest_week_label)
        wc2.metric("Servicios Realizados", f"{len(df_curr_week)} servicios")
        wc3.metric("Flete esta Semana", f"S/ {df_curr_week['flete_total'].sum():,.2f}")
        wc4.metric("Margen esta Semana", f"S/ {df_curr_week['margen_calculado'].sum():,.2f}")

# ---------------------------------------------------------
# TAB 2: Top de Clientes
# ---------------------------------------------------------
with tab2:
    use_liquidado = (flete_source == "Flete Liquidado (Reporte Margen)")

    all_legs = df.copy()
    all_legs['flete_usado'] = all_legs['flete_total'] if use_liquidado else all_legs['flete_gen_total']

    st.markdown('<div class="filter-bar-title">👥 Controles Directos para Análisis de Clientes</div>', unsafe_allow_html=True)
    fc1, fc2 = st.columns([3, 1])
    with fc1:
        all_clients_opt = sorted(all_legs['cliente'].unique().tolist())
        sel_tab2_clients = st.multiselect("🔍 Seleccionar/Excluir Clientes:", options=all_clients_opt, default=all_clients_opt, key="tab2_clients")
    with fc2:
        top_n_cli = st.selectbox("📊 Mostrar Top:", options=[5, 10, 15, 20, "Todos"], index=1, key="tab2_topn")

    all_legs_filtered = all_legs[all_legs['cliente'].isin(sel_tab2_clients)].copy()

    df_cli = all_legs_filtered.groupby('cliente').agg(
        Servicios=('id_servicio', 'count'),
        Flete_Total=('flete_usado', 'sum'),
        Flete_Liquidado=('flete_total', 'sum'),
        Flete_Registrado_General=('flete_gen_total', 'sum'),
        Gastos_Total=('gastos_total', 'sum'),
        Margen_Bruto=('margen_calculado', 'sum'),
        Flete_Promedio=('flete_usado', 'mean')
    ).reset_index()

    df_cli['Pct_Margen'] = np.where(df_cli['Flete_Liquidado'] > 0, (df_cli['Margen_Bruto'] / df_cli['Flete_Liquidado']) * 100, 0)

    top_flete_cli = df_cli.sort_values('Flete_Total', ascending=False).iloc[0] if len(df_cli) > 0 else None
    top_margen_cli = df_cli.sort_values('Margen_Bruto', ascending=False).iloc[0] if len(df_cli) > 0 else None

    ck1, ck2, ck3 = st.columns(3)
    with ck1:
        if top_flete_cli is not None:
            st.metric("🥇 Cliente con Mayor Flete", f"{top_flete_cli['cliente']}", f"S/ {top_flete_cli['Flete_Total']:,.2f}")
    with ck2:
        if top_margen_cli is not None:
            st.metric("🏆 Cliente con Mayor Margen Bruto", f"{top_margen_cli['cliente']}", f"S/ {top_margen_cli['Margen_Bruto']:,.2f}")
    with ck3:
        st.metric("👥 Total Clientes en Vista", f"{len(df_cli)} clientes")

    st.markdown("---")

    col_c1, col_c2 = st.columns(2)

    limit_cli = len(df_cli) if top_n_cli == "Todos" else int(top_n_cli)

    with col_c1:
        st.markdown(f"#### 💰 Top Clientes por Mayor Flete ({'Liquidado' if use_liquidado else 'Registrado'})")
        df_top_flete = df_cli.sort_values('Flete_Total', ascending=False).head(limit_cli)
        fig_cli_flete = px.bar(
            df_top_flete,
            x='Flete_Total',
            y='cliente',
            orientation='h',
            text_auto=',.0f',
            color='Flete_Total',
            color_continuous_scale='Blues',
            labels={'Flete_Total': 'Flete Total (S/)', 'cliente': 'Cliente'}
        )
        fig_cli_flete.update_traces(hovertemplate="<b>%{y}</b><br>💰 Flete: <b>S/ %{x:,.2f}</b><extra></extra>")
        fig_cli_flete.update_layout(template='plotly_dark', height=390, margin=dict(l=10, r=10, t=20, b=10), yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_cli_flete, use_container_width=True)

    with col_c2:
        st.markdown("#### 📈 Top Clientes por Mayor Margen Bruto")
        df_top_margen = df_cli.sort_values('Margen_Bruto', ascending=False).head(limit_cli)
        fig_cli_margen = px.bar(
            df_top_margen,
            x='Margen_Bruto',
            y='cliente',
            orientation='h',
            text_auto=',.0f',
            color='Margen_Bruto',
            color_continuous_scale='Greens',
            labels={'Margen_Bruto': 'Margen Bruto (S/)', 'cliente': 'Cliente'}
        )
        fig_cli_margen.update_traces(hovertemplate="<b>%{y}</b><br>📈 Margen Bruto: <b>S/ %{x:,.2f}</b><extra></extra>")
        fig_cli_margen.update_layout(template='plotly_dark', height=390, margin=dict(l=10, r=10, t=20, b=10), yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_cli_margen, use_container_width=True)

    st.markdown("#### 📊 Tabla Comparativa General de Clientes")
    df_cli_sort = df_cli.sort_values('Flete_Total', ascending=False)
    st.dataframe(
        df_cli_sort[['cliente', 'Servicios', 'Flete_Liquidado', 'Flete_Registrado_General', 'Gastos_Total', 'Margen_Bruto', 'Pct_Margen']].style.format({
            'Flete_Liquidado': 'S/ {:,.2f}',
            'Flete_Registrado_General': 'S/ {:,.2f}',
            'Gastos_Total': 'S/ {:,.2f}',
            'Margen_Bruto': 'S/ {:,.2f}',
            'Pct_Margen': '{:.1f}%'
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# TAB 3: Top de Placas
# ---------------------------------------------------------
with tab3:
    st.markdown('<div class="filter-bar-title">🚛 Filtros Directos de Vehículos / Placas</div>', unsafe_allow_html=True)
    fp1, fp2 = st.columns([3, 1])
    with fp1:
        placas_all_opt = sorted(df['placa'].dropna().unique())
        sel_tab3_placas = st.multiselect("🔍 Seleccionar Placas:", options=placas_all_opt, default=placas_all_opt, key="tab3_placas")
    with fp2:
        top_n_placas = st.selectbox("📊 Mostrar Top Placas:", options=[5, 10, 15, "Todas"], index=1, key="tab3_topn")

    df_p_filtered = df[df['placa'].isin(sel_tab3_placas)].copy()

    df_placas = df_p_filtered.groupby('placa').agg(
        Servicios=('cant_servicios', 'sum'),
        Flete_Total=('flete_total', 'sum'),
        Gastos_Total=('gastos_total', 'sum'),
        Margen_Bruto=('margen_calculado', 'sum'),
        Margen_Prom_Viaje=('margen_calculado', 'mean')
    ).reset_index()

    df_placas['Pct_Margen'] = np.where(df_placas['Flete_Total'] > 0, (df_placas['Margen_Bruto'] / df_placas['Flete_Total']) * 100, 0)
    df_placas = df_placas.sort_values('Margen_Bruto', ascending=False)

    col_p1, col_p2 = st.columns([3, 2])

    limit_p = len(df_placas) if top_n_placas == "Todas" else int(top_n_placas)

    with col_p1:
        st.markdown("#### 🏆 Ranking de Placas por Margen Bruto Generado (S/)")
        fig_placa_bar = px.bar(
            df_placas.head(limit_p),
            x='Margen_Bruto',
            y='placa',
            orientation='h',
            text_auto=',.0f',
            color='Margen_Bruto',
            color_continuous_scale='Viridis',
            labels={'Margen_Bruto': 'Margen Bruto (S/)', 'placa': 'Placa'}
        )
        fig_placa_bar.update_traces(hovertemplate="<b>Placa: %{y}</b><br>📈 Margen Bruto: <b>S/ %{x:,.2f}</b><extra></extra>")
        fig_placa_bar.update_layout(template='plotly_dark', height=390, margin=dict(l=10, r=10, t=20, b=10), yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_placa_bar, use_container_width=True)

    with col_p2:
        st.markdown("#### 📊 Participación en Cantidad de Servicios")
        fig_placa_pie = px.pie(
            df_placas.head(limit_p),
            names='placa',
            values='Servicios',
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_placa_pie.update_traces(hovertemplate="<b>Placa: %{label}</b><br>📦 Servicios: <b>%{value}</b> (%{percent})<extra></extra>")
        fig_placa_pie.update_layout(template='plotly_dark', height=390, margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig_placa_pie, use_container_width=True)

    st.markdown("#### 📑 Detalle Comparativo de Vehículos")
    st.dataframe(
        df_placas.style.format({
            'Flete_Total': 'S/ {:,.2f}',
            'Gastos_Total': 'S/ {:,.2f}',
            'Margen_Bruto': 'S/ {:,.2f}',
            'Margen_Prom_Viaje': 'S/ {:,.2f}',
            'Pct_Margen': '{:.1f}%'
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# TAB 4: Rutas y Tipos de Servicio
# ---------------------------------------------------------
with tab4:
    col_r1, col_r2 = st.columns(2)

    with col_r1:
        st.markdown('<div class="filter-bar-title">📍 1. Análisis por Rutas de Transporte</div>', unsafe_allow_html=True)
        rutas_opt = sorted(df['ruta'].dropna().unique())
        sel_rutas = st.multiselect("🔍 Filtrar Rutas:", options=rutas_opt, default=rutas_opt, key="tab4_rutas")
        
        df_r_filtered = df[df['ruta'].isin(sel_rutas)]
        
        df_rutas = df_r_filtered.groupby('ruta').agg(
            Servicios=('cant_servicios', 'sum'),
            Margen_Bruto=('margen_calculado', 'sum')
        ).reset_index().sort_values('Servicios', ascending=False).head(10)

        fig_rutas = px.bar(
            df_rutas,
            x='Servicios',
            y='ruta',
            orientation='h',
            color='Margen_Bruto',
            text_auto=True,
            color_continuous_scale='Blues',
            labels={'ruta': 'Ruta', 'Servicios': 'N° de Servicios'}
        )
        fig_rutas.update_traces(hovertemplate="<b>Ruta: %{y}</b><br>📦 Servicios: <b>%{x}</b><extra></extra>")
        fig_rutas.update_layout(template='plotly_dark', height=340, margin=dict(l=10, r=10, t=20, b=10), yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_rutas, use_container_width=True)

        r_flete = df_r_filtered['flete_total'].sum()
        r_gastos = df_r_filtered['gastos_total'].sum()
        r_margen = df_r_filtered['margen_calculado'].sum()
        r_pct = (r_margen / r_flete * 100) if r_flete > 0 else 0
        r_serv = int(df_r_filtered['cant_servicios'].sum())

        st.markdown(f"""
            <div class="summary-chip-container">
                <div class="chip chip-purple">📦 <b>{r_serv}</b> servicios</div>
                <div class="chip chip-blue">💰 Flete: <b>S/ {r_flete:,.2f}</b></div>
                <div class="chip chip-red">💸 Gastos: <b>S/ {r_gastos:,.2f}</b></div>
                <div class="chip chip-green">📈 Margen: <b>S/ {r_margen:,.2f}</b> ({r_pct:.1f}%)</div>
            </div>
        """, unsafe_allow_html=True)

    with col_r2:
        st.markdown('<div class="filter-bar-title">📦 2. Análisis por Tipo de Servicio</div>', unsafe_allow_html=True)
        tipos_opt = sorted(df['tipo_servicio'].dropna().unique())
        sel_tipos = st.multiselect("🔍 Filtrar Tipos de Servicio:", options=tipos_opt, default=tipos_opt, key="tab4_tipos")
        
        df_t_filtered = df[df['tipo_servicio'].isin(sel_tipos)]

        df_tipos = df_t_filtered.groupby('tipo_servicio').agg(
            Servicios=('cant_servicios', 'sum'),
            Flete_Total=('flete_total', 'sum'),
            Gastos_Total=('gastos_total', 'sum'),
            Margen_Bruto=('margen_calculado', 'sum')
        ).reset_index().sort_values('Margen_Bruto', ascending=False)

        fig_tipos = px.bar(
            df_tipos,
            x='tipo_servicio',
            y=['Flete_Total', 'Gastos_Total', 'Margen_Bruto'],
            barmode='group',
            labels={'value': 'Monto (S/)', 'tipo_servicio': 'Tipo de Servicio', 'variable': 'Métrica'},
            color_discrete_sequence=['#38bdf8', '#f43f5e', '#10b981']
        )
        fig_tipos.update_traces(hovertemplate="<b>%{x} (%{variable})</b><br>Monto: <b>S/ %{y:,.2f}</b><extra></extra>")
        fig_tipos.update_layout(template='plotly_dark', height=340, margin=dict(l=10, r=10, t=20, b=10), legend=dict(title=""))
        st.plotly_chart(fig_tipos, use_container_width=True)

        t_flete = df_t_filtered['flete_total'].sum()
        t_gastos = df_t_filtered['gastos_total'].sum()
        t_margen = df_t_filtered['margen_calculado'].sum()
        t_pct = (t_margen / t_flete * 100) if t_flete > 0 else 0
        t_serv = int(df_t_filtered['cant_servicios'].sum())

        st.markdown(f"""
            <div class="summary-chip-container">
                <div class="chip chip-purple">📦 <b>{t_serv}</b> servicios</div>
                <div class="chip chip-blue">💰 Flete Total: <b>S/ {t_flete:,.2f}</b></div>
                <div class="chip chip-red">💸 Gastos: <b>S/ {t_gastos:,.2f}</b></div>
                <div class="chip chip-green">📈 Margen: <b>S/ {t_margen:,.2f}</b> ({t_pct:.1f}%)</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 🏢 Operadores de Flota (Propios vs Tercerizados)")
    df_oper = df.groupby('operador').agg(
        Servicios=('cant_servicios', 'sum'),
        Flete_Total=('flete_total', 'sum'),
        Gastos_Total=('gastos_total', 'sum'),
        Margen_Bruto=('margen_calculado', 'sum')
    ).reset_index().sort_values('Margen_Bruto', ascending=False)
    
    st.dataframe(
        df_oper.style.format({
            'Flete_Total': 'S/ {:,.2f}',
            'Gastos_Total': 'S/ {:,.2f}',
            'Margen_Bruto': 'S/ {:,.2f}'
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# TAB 5: Registro Detallado y Exportación
# ---------------------------------------------------------
with tab5:
    st.markdown('<div class="filter-bar-title">📋 Explorador y Buscador Avanzado de Registros</div>', unsafe_allow_html=True)
    
    bcol1, bcol2 = st.columns([3, 1])
    with bcol1:
        search_term = st.text_input("🔍 Búsqueda rápida por texto (Cliente, Chofer, Placa, Ruta, Servicio):", "", key="tab5_search")
    with bcol2:
        sel_tab5_tipo = st.selectbox("Filtrar por Tipo:", options=["Todos", "EXCLUSIVO", "CONSOLIDADO", "REPARTO", "PAQUETERIA"], index=0, key="tab5_tipo")

    df_show = df.copy()
    if sel_tab5_tipo != "Todos":
        df_show = df_show[df_show['tipo_servicio'] == sel_tab5_tipo]

    if search_term:
        df_show = df_show[
            df_show['cliente'].str.contains(search_term.upper(), na=False) |
            df_show['chofer'].str.contains(search_term.upper(), na=False) |
            df_show['placa'].str.contains(search_term.upper(), na=False) |
            df_show['ruta'].str.contains(search_term.upper(), na=False) |
            df_show['id_servicio'].astype(str).str.contains(search_term, na=False)
        ]
        
    display_cols = [
        'id_servicio', 'tramo', 'fecha_salida', 'semana_label', 'cliente', 'chofer', 'placa', 
        'ruta', 'tipo_servicio', 'flete_total', 'gastos_total', 'margen_calculado'
    ]
    
    st.dataframe(
        df_show[display_cols].style.format({
            'flete_total': 'S/ {:,.2f}',
            'gastos_total': 'S/ {:,.2f}',
            'margen_calculado': 'S/ {:,.2f}'
        }),
        use_container_width=True,
        height=450
    )

    csv = df_show.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Descargar Registros Filtrados en CSV",
        data=csv,
        file_name="reporte_sistrans_filtrado.csv",
        mime="text/csv"
    )

# ---------------------------------------------------------
# TAB 6: Auditoría y Comparativo de Regularizaciones
# ---------------------------------------------------------
with tab6:
    st.markdown('<div class="filter-bar-title">🔄 Auditoría de Regularizaciones y Ajustes de Montos</div>', unsafe_allow_html=True)
    
    if baseline_mrg_path is None or baseline_gen_path is None:
        st.info("ℹ️ Para activar la auditoría comparativa, añade una segunda carpeta de semana en `data/` (por ejemplo `data/semana_31/`). El sistema comparará automáticamente la versión nueva contra la anterior.")
        
        # Display Baseline Audit Summary
        st.markdown("#### 📌 Resumen General del Corte Actual")
        st.markdown(f"""
            - **Corte Activo**: `{selected_week_label}`
            - **Servicios Evaluados**: `{len(df_all)} servicios`
            - **Facturación Total**: `S/ {df_all['flete_total'].sum():,.2f}`
            - **Gastos Operativos**: `S/ {df_all['gastos_total'].sum():,.2f}`
            - **Margen Bruto**: `S/ {df_all['margen_calculado'].sum():,.2f}`
        """)
    else:
        # Load Baseline Dataset
        df_base_all = process_excel_files(baseline_mrg_path, baseline_gen_path)
        
        # Merge datasets to detect variations
        merged_audit = pd.merge(
            df_all[['id_servicio', 'tramo', 'fecha_salida', 'cliente', 'ruta', 'placa', 'flete_total', 'gastos_total', 'margen_calculado']],
            df_base_all[['id_servicio', 'flete_total', 'gastos_total', 'margen_calculado']],
            on='id_servicio', how='outer', suffixes=('_nuevo', '_base')
        )
        
        merged_audit['flete_base'] = merged_audit['flete_total_base'].fillna(0)
        merged_audit['flete_nuevo'] = merged_audit['flete_total_nuevo'].fillna(0)
        merged_audit['diff_flete'] = merged_audit['flete_nuevo'] - merged_audit['flete_base']
        
        merged_audit['gastos_base'] = merged_audit['gastos_total_base'].fillna(0)
        merged_audit['gastos_nuevo'] = merged_audit['gastos_total_nuevo'].fillna(0)
        merged_audit['diff_gastos'] = merged_audit['gastos_nuevo'] - merged_audit['gastos_base']
        
        merged_audit['margen_base'] = merged_audit['margen_calculado_base'].fillna(0)
        merged_audit['margen_nuevo'] = merged_audit['margen_calculado_nuevo'].fillna(0)
        merged_audit['diff_margen'] = merged_audit['margen_nuevo'] - merged_audit['margen_base']
        
        def classify_status(r):
            if pd.isnull(r['flete_total_base']):
                return '🟢 Nuevo Registro'
            elif pd.isnull(r['flete_total_nuevo']):
                return '🔴 Eliminado / Anulado'
            elif abs(r['diff_flete']) > 0.01 or abs(r['diff_gastos']) > 0.01:
                return '🟡 Modificado / Ajustado'
            else:
                return '⚪ Sin Cambios'
                
        merged_audit['estado_cambio'] = merged_audit.apply(classify_status, axis=1)

        # Audit Executive Cards
        ak1, ak2, ak3, ak4, ak5 = st.columns(5)
        
        delta_flete = merged_audit['diff_flete'].sum()
        delta_gastos = merged_audit['diff_gastos'].sum()
        delta_margen = merged_audit['diff_margen'].sum()
        n_mod = (merged_audit['estado_cambio'] == '🟡 Modificado / Ajustado').sum()
        n_new = (merged_audit['estado_cambio'] == '🟢 Nuevo Registro').sum()

        with ak1:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">💰 Δ Facturación (Flete)</div>
                    <div class="metric-value" style="color: {'#10b981' if delta_flete >= 0 else '#f43f5e'};">S/ {delta_flete:+,.2f}</div>
                    <div class="metric-sub">Variación en Flete Total</div>
                </div>
            """, unsafe_allow_html=True)
            
        with ak2:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">💸 Δ Gastos Operativos</div>
                    <div class="metric-value" style="color: {'#f43f5e' if delta_gastos > 0 else '#10b981'};">S/ {delta_gastos:+,.2f}</div>
                    <div class="metric-sub">Variación en Costos</div>
                </div>
            """, unsafe_allow_html=True)

        with ak3:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">📈 Δ Margen Bruto</div>
                    <div class="metric-value" style="color: {'#10b981' if delta_margen >= 0 else '#f43f5e'};">S/ {delta_margen:+,.2f}</div>
                    <div class="metric-sub">Impacto Neto en Soles</div>
                </div>
            """, unsafe_allow_html=True)

        with ak4:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">🟡 Modificados</div>
                    <div class="metric-value" style="color: #f59e0b;">{n_mod}</div>
                    <div class="metric-sub">Servicios Ajustados</div>
                </div>
            """, unsafe_allow_html=True)

        with ak5:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">🟢 Nuevos Servicios</div>
                    <div class="metric-value" style="color: #10b981;">{n_new}</div>
                    <div class="metric-sub">Incorporados</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        
        st.markdown("#### 📋 Matriz Detallada de Auditoría por Servicio")
        
        audit_filter = st.radio(
            "🔍 Filtrar vista de auditoría:",
            options=["Todos", "Solo Modificados / Ajustados", "Solo Nuevos Registros"],
            horizontal=True
        )
        
        df_audit_show = merged_audit.copy()
        if audit_filter == "Solo Modificados / Ajustados":
            df_audit_show = df_audit_show[df_audit_show['estado_cambio'] == '🟡 Modificado / Ajustado']
        elif audit_filter == "Solo Nuevos Registros":
            df_audit_show = df_audit_show[df_audit_show['estado_cambio'] == '🟢 Nuevo Registro']
            
        audit_cols = [
            'id_servicio', 'cliente', 'ruta', 'placa', 'estado_cambio',
            'flete_base', 'flete_nuevo', 'diff_flete',
            'gastos_base', 'gastos_nuevo', 'diff_gastos',
            'margen_base', 'margen_nuevo', 'diff_margen'
        ]
        
        st.dataframe(
            df_audit_show[audit_cols].style.format({
                'flete_base': 'S/ {:,.2f}',
                'flete_nuevo': 'S/ {:,.2f}',
                'diff_flete': 'S/ {:+,.2f}',
                'gastos_base': 'S/ {:,.2f}',
                'gastos_nuevo': 'S/ {:,.2f}',
                'diff_gastos': 'S/ {:+,.2f}',
                'margen_base': 'S/ {:,.2f}',
                'margen_nuevo': 'S/ {:,.2f}',
                'diff_margen': 'S/ {:+,.2f}'
            }),
            use_container_width=True,
            height=450
        )
        
        csv_audit = df_audit_show[audit_cols].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar Reporte Completo de Auditoría en CSV",
            data=csv_audit,
            file_name="reporte_auditoria_regularizaciones.csv",
            mime="text/csv"
        )

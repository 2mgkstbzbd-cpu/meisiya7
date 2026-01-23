import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import time
import os
import socket
import hashlib
import re
from datetime import datetime
import html as _html
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
def _is_nan(x):
    try:
        return x != x
    except Exception:
        return False

def fmt_num(x, na="—"):
    if x is None or _is_nan(x):
        return na
    try:
        s = f"{float(x):,.2f}"
    except Exception:
        return str(x)
    s = s.rstrip("0").rstrip(".")
    return s

def fmt_pct_ratio(r, na="—", decimals=1):
    if r is None or _is_nan(r):
        return na
    v = float(r) * 100.0
    s = f"{v:.{decimals}f}".rstrip("0").rstrip(".")
    return f"{s}%"

def fmt_pct_value(p, na="—", decimals=1):
    if p is None or _is_nan(p):
        return na
    v = float(p)
    sign = "+" if v > 0 else ("-" if v < 0 else "")
    s = f"{abs(v):.{decimals}f}".rstrip("0").rstrip(".")
    return f"{sign}{s}%"

_COORD_NUM_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")

def _parse_lon_lat(v):
    if v is None:
        return None, None
    s = str(v).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None, None
    nums = _COORD_NUM_RE.findall(s)
    if len(nums) < 2:
        return None, None
    a = float(nums[0])
    b = float(nums[1])

    def _is_lon(x): return 70 <= x <= 140
    def _is_lat(x): return 0 <= x <= 60

    if _is_lon(a) and _is_lat(b):
        lon, lat = a, b
    elif _is_lon(b) and _is_lat(a):
        lon, lat = b, a
    else:
        lon, lat = (a, b) if abs(a) >= abs(b) else (b, a)
    if not _is_lon(lon) or not _is_lat(lat):
        return None, None
    return lon, lat

def get_host_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

# -----------------------------------------------------------------------------
# 1. Page Config
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="美思雅数据分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)
st.markdown('<meta name="google" content="notranslate" />', unsafe_allow_html=True)
st.markdown("""
<script>
  document.documentElement.setAttribute("translate", "no");
  document.querySelector('html').setAttribute("translate", "no");
</script>
""", unsafe_allow_html=True)

_required_password = os.getenv("DASHBOARD_PASSWORD", "").strip()
if _required_password:
    if not st.session_state.get("_authed", False):
        st.markdown("### 🔒 访问验证")
        _pwd = st.text_input("请输入访问密码", type="password")
        if st.button("验证", type="primary"):
            if _pwd == _required_password:
                st.session_state["_authed"] = True
                st.rerun()
            else:
                st.error("密码错误")
        st.stop()

if 'drill_level' not in st.session_state:
    st.session_state.drill_level = 1
if 'selected_prov' not in st.session_state:
    st.session_state.selected_prov = None
if 'selected_dist' not in st.session_state:
    st.session_state.selected_dist = None
if 'perf_time_mode' not in st.session_state:
    st.session_state.perf_time_mode = '近12个月'
if 'perf_provs' not in st.session_state:
    st.session_state.perf_provs = []
if 'perf_cats' not in st.session_state:
    st.session_state.perf_cats = []

# -----------------------------------------------------------------------------
# 2. Custom CSS
# -----------------------------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
    :root {
        --bg-1: #F5F5F7;
        --bg-2: #ECECEC;
        --bg-3: #E0E0E0;
        --panel: #FFFFFF;
        --panel-2: rgba(255, 255, 255, 0.85);
        --stroke: #E0E0E0;
        --stroke-strong: #D1D1D1;
        --text: #1B1530;
        --text-muted: rgba(27, 21, 48, 0.7);
        --primary: #5B2EA6;
        --primary-2: #6A3AD0;
        --accent: #FFC400;
        --accent-2: #FFB000;
        --danger: #E5484D;
        --success: #2FBF71;
        --shadow: 0 10px 26px rgba(0, 0, 0, 0.08);
        --shadow-soft: 0 4px 12px rgba(0, 0, 0, 0.05);
        --radius: 12px;
        --radius-sm: 10px;
        --transition: 240ms cubic-bezier(.2,.8,.2,1);
        --focus: 0 0 0 3px rgba(91, 46, 166, 0.2);
        --tbl-header-bg: #4285F4;
        --tbl-header-bg-hover: #2F76E4;
        --tbl-header-border: #2B63C4;
        --tbl-header-fg: #FFFFFF;
        --tbl-header-icon: rgba(255, 255, 255, 0.92);
        --tbl-header-shadow: 0 6px 16px rgba(0, 0, 0, 0.16);
        --tbl-header-font-size: 15px;
        --tbl-header-font-weight: 800;
        --tbl-cell-font-size: 13px;
    }
    
    @media (prefers-color-scheme: dark) {
        :root {
            --tbl-header-bg: #2B66D9;
            --tbl-header-bg-hover: #2358C2;
            --tbl-header-border: #1B46A0;
            --tbl-header-fg: #FFFFFF;
            --tbl-header-icon: rgba(255, 255, 255, 0.95);
            --tbl-header-shadow: 0 10px 22px rgba(0, 0, 0, 0.32);
        }
    }

    html, body, [class*="css"] {
        font-family: 'Inter', 'Microsoft YaHei', sans-serif;
        color: var(--text);
    }

    .stApp {
        background: #F5F5F7;
    }

    [data-testid="stSidebar"], [data-testid="collapsedControl"] {
        display: none !important;
    }

    div[data-testid="stDataFrame"] thead tr th,
    div[data-testid="stTable"] thead tr th {
        background: var(--tbl-header-bg) !important;
        color: var(--tbl-header-fg) !important;
        font-weight: var(--tbl-header-font-weight) !important;
        font-size: var(--tbl-header-font-size) !important;
        border-bottom: 1px solid var(--tbl-header-border) !important;
    }

    div[data-testid="stDataFrame"] thead tr th:hover,
    div[data-testid="stTable"] thead tr th:hover {
        background: var(--tbl-header-bg-hover) !important;
    }

    div[data-testid="stDataFrame"] thead tr th:active,
    div[data-testid="stTable"] thead tr th:active {
        box-shadow: var(--tbl-header-shadow) !important;
    }

    .out-kpi-card {
        background: linear-gradient(180deg, rgba(66,133,244,0.08) 0%, rgba(255,255,255,0.92) 60%, #FFFFFF 100%);
        border-radius: 14px;
        padding: 16px 16px 14px;
        border: 1px solid rgba(66,133,244,0.22);
        box-shadow: 0 10px 26px rgba(0,0,0,0.06);
        margin-bottom: 10px;
        position: relative;
        overflow: hidden;
    }
    .out-kpi-bar {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 3px;
        background: linear-gradient(90deg, var(--tbl-header-bg) 0%, var(--success) 60%, var(--accent) 100%);
        opacity: 0.9;
    }
    .out-kpi-head { display:flex; align-items:center; gap:10px; margin-bottom: 10px; }
    .out-kpi-ico {
        width: 34px;
        height: 34px;
        border-radius: 10px;
        display:flex;
        justify-content:center;
        align-items:center;
        background: rgba(66,133,244,0.16);
        border: 1px solid rgba(66,133,244,0.28);
        color: var(--tbl-header-bg);
        font-weight: 900;
        font-size: 18px;
    }
    .out-kpi-title { font-size: 15px; color: rgba(27,21,48,0.78); font-weight: 800; letter-spacing: 0.2px; }
    .out-kpi-val { font-size: 26px; font-weight: 900; color: #1B1530; margin-bottom: 4px; }
    .out-kpi-sub { font-size: 13px; display:flex; justify-content:space-between; align-items:center; color: rgba(27,21,48,0.72); }
    .out-kpi-sub2 { font-size: 12px; display:flex; justify-content:space-between; align-items:center; color: rgba(27,21,48,0.62); margin-top: 4px; }
    .out-kpi-progress { background: rgba(27,21,48,0.10); border-radius: 999px; height: 6px; width: 100%; overflow: hidden; }
    .out-kpi-progress-bar { height: 100%; border-radius: 999px; }
    .trend-up { color: var(--success); font-weight: 800; }
    .trend-down { color: var(--danger); font-weight: 800; }
    .trend-neutral { color: rgba(27,21,48,0.72); font-weight: 800; }
    @media (max-width: 768px) {
        .out-kpi-card { padding: 14px 14px 12px; }
        .out-kpi-val { font-size: 22px; }
        .out-kpi-title { font-size: 14px; }
    }

    h1, h2, h3, h4, h5, h6 {
        color: var(--text);
        letter-spacing: 0.2px;
        text-shadow: none;
    }

    [data-testid="stAppViewContainer"] {
        color: var(--text);
    }

    /* Reset global text visibility */
    .stMarkdown, .stText, p, li, span, label {
        color: var(--text) !important;
        text-shadow: none;
    }
    
    /* Caption specific */
    .stCaption {
        color: var(--text-muted) !important;
    }

    /* --- SIDEBAR STYLING REMOVED TO RESTORE VISIBILITY --- */
    
    /* Only keep global metric styling that doesn't affect visibility */
    div[data-testid="stMetric"] {
        background: var(--panel);
        border: 1px solid var(--stroke);
        border-radius: var(--radius);
        box-shadow: var(--shadow-soft);
        padding: 18px;
    }

    div[data-testid="stMetric"] * {
        color: var(--text) !important;
    }
    
    div[data-testid="stMetric"] label {
        color: var(--text-muted) !important;
    }

    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: var(--primary) !important;
    }

    div[data-testid="stMetric"] div[data-testid="stMetricDelta"] {
        color: var(--accent-2) !important;
    }
    
    /* Buttons */
    div.stButton > button {
        border-radius: var(--radius-sm);
    }
    
    /* Analysis Button Customization */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #FFC400 0%, #FFB000 100%) !important;
        border: 1px solid rgba(255, 176, 0, 0.4) !important;
        color: #5B2EA6 !important;
        font-weight: 700 !important;
        text-shadow: none !important;
        box-shadow: 0 4px 12px rgba(255, 196, 0, 0.25) !important;
        transition: all 0.2s ease !important;
    }

    div.stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #FFD54F 0%, #FFC107 100%) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 16px rgba(255, 196, 0, 0.35) !important;
        border-color: rgba(255, 176, 0, 0.6) !important;
    }
    
    div.stButton > button[kind="primary"]:active {
        transform: translateY(1px) !important;
        box-shadow: 0 2px 8px rgba(255, 196, 0, 0.2) !important;
    }
    
    /* Tabs styling kept simple */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }

    /* Outbound subtabs (radio styled as tabs) */
    div[data-testid="stRadio"] .out-subtab-hint {display:none;}
    div[data-testid="stRadio"] [data-baseweb="radio"] > div {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    div[data-testid="stRadio"] [data-baseweb="radio"] input {
        position: absolute !important;
        opacity: 0 !important;
        pointer-events: none !important;
    }
    div[data-testid="stRadio"] [data-baseweb="radio"] div[role="radio"] {
        display: none !important;
    }
    div[data-testid="stRadio"] [data-baseweb="radio"] span {
        font-weight: 600 !important;
        color: rgba(27, 21, 48, 0.75) !important;
    }
    div[data-testid="stRadio"] [data-baseweb="radio"] input:checked ~ div span {
        color: rgba(27, 21, 48, 0.95) !important;
    }
    div[data-testid="stRadio"] [data-testid="stRadio"] > div[role="radiogroup"] {
        border-bottom: 1px solid rgba(0, 0, 0, 0.08) !important;
        padding-bottom: 6px !important;
        gap: 10px !important;
    }
    div[data-testid="stRadio"] [data-baseweb="radio"] {
        position: relative !important;
        padding: 8px 0 10px 0 !important;
        margin-right: 14px !important;
    }
    div[data-testid="stRadio"] [data-baseweb="radio"] input:checked ~ div::after {
        content: "" !important;
        position: absolute !important;
        left: 0 !important;
        right: 0 !important;
        bottom: -7px !important;
        height: 2px !important;
        background: #E5484D !important;
        border-radius: 2px !important;
        transition: all 0.2s ease !important;
    }

    .out-subtab-content {
        animation: outFadeUp 240ms cubic-bezier(.2,.8,.2,1);
    }
    @keyframes outFadeUp {
        from { opacity: 0; transform: translateY(8px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    /* Ensure DataFrame styling is applied even if internal structure varies */
    [data-testid="stDataFrame"] {
        background: var(--panel);
        border: 1px solid var(--stroke);
        border-radius: var(--radius);
        box-shadow: var(--shadow-soft);
        overflow: hidden;
    }
    
    /* Target all possible table cells within the dataframe container */
    [data-testid="stDataFrame"] td, 
    [data-testid="stDataFrame"] th,
    [data-testid="stDataFrame"] [role="gridcell"],
    [data-testid="stDataFrame"] [role="columnheader"],
    [data-testid="stDataFrame"] div[data-testid="stDataFrameResizable"] {
        text-align: center !important;
        vertical-align: middle !important;
        color: var(--text) !important;
        display: flex;
        justify-content: center;
        align-items: center;
    }
    
    /* Force header content center */
    [data-testid="stDataFrame"] [role="columnheader"] > div {
        justify-content: center !important;
        text-align: center !important;
        width: 100%;
        display: flex;
    }
    
    /* Force cell content center */
    [data-testid="stDataFrame"] [role="gridcell"] > div {
        justify-content: center !important;
        text-align: center !important;
        width: 100%;
        display: flex;
    }

    /* Essential Visibility Controls */
    button[kind="header"], [data-testid="collapsedControl"] {
        visibility: visible !important;
        z-index: 999999 !important;
    }

    header {visibility: visible !important;}
    [data-testid="stSidebarNav"] {display: block !important;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display:none;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    [data-testid="stToolbar"] {display: none !important;}
    [data-testid="stHeader"] {background: transparent !important;}
    [data-testid="stHeader"] a {display:none !important;}
    [data-testid="stViewerBadge"] {display:none !important;}
    [data-testid="stGitHubLink"] {display:none !important;}

    /* Sidebar explicit frosted light scheme to ensure readability */
    [data-testid="stSidebar"] {
        background: rgba(255, 255, 255, 0.88) !important;
        backdrop-filter: blur(10px) !important;
        border-right: 1px solid rgba(0, 0, 0, 0.08) !important;
    }
    [data-testid="stSidebar"] * {
        color: #333333 !important;
    }
    [data-testid="stSidebar"] summary svg,
    [data-testid="stSidebar"] svg {
        fill: #333333 !important;
    }
    [data-testid="stSidebar"] details[data-testid="stExpander"] > summary:hover {
        background: rgba(0,0,0,0.05) !important;
    }

    /* Sidebar inputs and selects */
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] div[data-baseweb="select"] > div {
        background: rgba(255,255,255,0.95) !important;
        border: 1px solid rgba(0,0,0,0.15) !important;
        color: #333333 !important;
    }
    [data-testid="stSidebar"] input::placeholder {
        color: rgba(0,0,0,0.55) !important;
    }
    [data-testid="stSidebar"] div[data-baseweb="select"] > div:hover,
    [data-testid="stSidebar"] input:hover {
        border-color: rgba(91,46,166,0.6) !important;
    }

    /* File uploader dropzone */
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
        background: rgba(255,255,255,0.92) !important;
        border: 1px dashed rgba(0,0,0,0.18) !important;
        color: #333333 !important;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] span {
        color: #333333 !important;
    }
    [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {
        background: rgba(91,46,166,0.10) !important;
        color: #333333 !important;
        border: 1px solid rgba(91,46,166,0.35) !important;
    }

    /* Collapsed control arrow visibility and contrast */
    [data-testid="collapsedControl"] {
        color: #333333 !important;
        background: rgba(255,255,255,0.55) !important;
        border: 1px solid rgba(0,0,0,0.12) !important;
        border-radius: 6px !important;
        top: 56px !important;
    }

    @media (max-width: 768px) {
        div[data-testid="stMetric"] {
            padding: 14px;
        }
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
    div[data-testid="stDataFrame"] div[role="gridcell"] { display: flex; align-items: center; }
    div[data-testid="stDataFrame"] div[role="columnheader"] { display: flex; align-items: center; justify-content: center; }
    .msy-table-wrap {
        width: 100%;
        overflow-x: auto;
        border-radius: 12px;
        border: 1px solid rgba(0,0,0,0.08);
        background: rgba(255,255,255,0.9);
        box-shadow: 0 6px 18px rgba(0,0,0,0.06);
    }
    table.msy-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: auto;
        font-size: 14px;
        line-height: 1.45;
    }
    table.msy-table thead th {
        position: sticky;
        top: 0;
        background: #1F2937;
        color: #F9FAFB;
        font-weight: 700;
        padding: 10px 12px;
        border: 1px solid rgba(255,255,255,0.12);
        text-align: center;
        vertical-align: middle;
        white-space: nowrap;
    }
    table.msy-table tbody td {
        padding: 10px 12px;
        border: 1px solid rgba(0,0,0,0.08);
        text-align: center;
        vertical-align: middle;
        font-variant-numeric: tabular-nums;
        white-space: nowrap;
    }
    table.msy-table tbody tr:nth-child(even) td {
        background: #F8FAFC;
    }
    table.msy-table tbody tr:hover td {
        background: #EEF2FF;
    }
</style>
""", unsafe_allow_html=True)

def _format_cell(v):
    if v is None or pd.isna(v):
        return ""
    if isinstance(v, (int, float, np.integer, np.floating)):
        return fmt_num(v, na="")
    return str(v)


# -----------------------------------------------------------------------------
# AgGrid Helper
# -----------------------------------------------------------------------------
JS_COLOR_CONDITIONAL = JsCode("""
function(params) {
    if (params.value > 0) {
        return {'color': '#28A745', 'textAlign': 'center', 'fontWeight': 'bold'};
    } else if (params.value < 0) {
        return {'color': '#DC3545', 'textAlign': 'center', 'fontWeight': 'bold'};
    }
    return {'textAlign': 'center'};
}
""")

JS_CENTER = JsCode("""
function(params) {
    return {'textAlign': 'center'};
}
""")

JS_FMT_NUM = JsCode("""
function(params) {
    const v = params.value;
    if (v === null || v === undefined) return '';
    const n = Number(v);
    if (!isFinite(n)) return '';
    return n.toLocaleString('zh-CN', {minimumFractionDigits: 1, maximumFractionDigits: 1});
}
""")

JS_FMT_PCT_RATIO = JsCode("""
function(params) {
    const v = params.value;
    if (v === null || v === undefined) return '';
    const n = Number(v);
    if (!isFinite(n)) return '';
    const p = n * 100;
    return p.toLocaleString('zh-CN', {minimumFractionDigits: 1, maximumFractionDigits: 1}) + '%';
}
""")

# Custom Cell Renderer for Progress Bar (Mockup using HTML)
JS_PROGRESS_BAR = JsCode("""
class ProgressBarRenderer {
    init(params) {
        this.eGui = document.createElement('div');
        this.eGui.style.width = '100%';
        this.eGui.style.height = '100%';
        this.eGui.style.display = 'flex';
        this.eGui.style.alignItems = 'center';
        
        const fmt1 = (v) => {
            if (v === null || v === undefined) return '';
            const n = Number(v);
            if (!isFinite(n)) return '';
            if (Math.abs(n - Math.round(n)) < 1e-9) return Math.round(n).toLocaleString('zh-CN');
            return n.toLocaleString('zh-CN', {minimumFractionDigits: 1, maximumFractionDigits: 1});
        };

        let value = params.value;
        if (value === null || value === undefined) value = 0;

        let maxValue = 0;
        if (params.colDef && params.colDef.cellRendererParams && params.colDef.cellRendererParams.maxValue !== undefined) {
            maxValue = Number(params.colDef.cellRendererParams.maxValue) || 0;
        }

        const percent = maxValue > 0 ? Math.min(Math.max((Number(value) / maxValue) * 100, 0), 100) : 0;

        let color = '#007bff';
        if (percent >= 100) color = '#28a745';
        else if (percent < 60) color = '#dc3545';
        else color = '#ffc107';
        
        this.eGui.innerHTML = `
            <div style="width: 100%; height: 20px; background-color: #e9ecef; border-radius: 3px; position: relative;">
                <div style="width: ${percent}%; height: 100%; background-color: ${color}; border-radius: 3px; transition: width 0.5s;"></div>
                <span style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; text-align: center; line-height: 20px; font-size: 12px; color: #000;">${fmt1(value)}</span>
            </div>
        `;
    }
    getGui() {
        return this.eGui;
    }
}
""")

# Custom Cell Renderer for Count (No %)
JS_PROGRESS_BAR_COUNT = JsCode("""
class ProgressBarCountRenderer {
    init(params) {
        this.eGui = document.createElement('div');
        this.eGui.style.width = '100%';
        this.eGui.style.height = '100%';
        this.eGui.style.display = 'flex';
        this.eGui.style.alignItems = 'center';
        
        const fmt1 = (v) => {
            if (v === null || v === undefined) return '';
            const n = Number(v);
            if (!isFinite(n)) return '';
            if (Math.abs(n - Math.round(n)) < 1e-9) return Math.round(n).toLocaleString('zh-CN');
            return n.toLocaleString('zh-CN', {minimumFractionDigits: 1, maximumFractionDigits: 1});
        };

        let value = params.value;
        if (value === null || value === undefined) value = 0;

        let maxValue = 0;
        if (params.colDef && params.colDef.cellRendererParams && params.colDef.cellRendererParams.maxValue !== undefined) {
            maxValue = Number(params.colDef.cellRendererParams.maxValue) || 0;
        }

        const percent = maxValue > 0 ? Math.min(Math.max((Number(value) / maxValue) * 100, 0), 100) : 0;

        let color = '#28a745';
        if (percent > 0) color = '#ffc107';
        if (percent >= 60) color = '#dc3545';
        
        this.eGui.innerHTML = `
            <div style="width: 100%; height: 20px; background-color: #e9ecef; border-radius: 3px; position: relative;">
                <div style="width: ${percent}%; height: 100%; background-color: ${color}; border-radius: 3px; transition: width 0.5s;"></div>
                <span style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; text-align: center; line-height: 20px; font-size: 12px; color: #000;">${fmt1(value)}</span>
            </div>
        `;
    }
    getGui() {
        return this.eGui;
    }
}
""")

def show_aggrid_table(df, height=None, key=None, on_row_selected=None, 
                      columns_props=None, 
                      column_defs=None,
                      grid_options_overrides=None,
                      auto_height_limit=2000):
    """
    Standardized AgGrid Table
    :param df: DataFrame to display
    :param height: Fixed height (optional)
    :param key: Unique key
    :param on_row_selected: 'single' or 'multiple' or None
    :param columns_props: Dict of col_name -> {type: 'percent'|'money'|'growth'|'bar', ...}
    :param auto_height_limit: Max height for auto calculation
    """
    if df is None or df.empty:
        # Custom Empty State
        st.markdown("""
            <div style="text-align: center; padding: 40px; background: #f8f9fa; border-radius: 8px; border: 1px dashed #d9d9d9;">
                <div style="font-size: 24px; margin-bottom: 10px;">📭</div>
                <div style="color: #666; font-size: 14px;">暂无数据</div>
            </div>
        """, unsafe_allow_html=True)
        return None

    # Inject CSS for Custom AgGrid Styling
    st.markdown("""
        <style>
        /* --- 1. Header Styling --- */
        .ag-header {
            background-color: var(--tbl-header-bg) !important;
            border-bottom: 1px solid var(--tbl-header-border) !important;
        }
        .ag-header-row,
        .ag-header-group-cell,
        .ag-header-cell {
            background-color: var(--tbl-header-bg) !important;
        }
        .ag-header-group-cell:hover,
        .ag-header-cell:hover {
            background-color: var(--tbl-header-bg-hover) !important;
        }
        .ag-header-group-cell:active,
        .ag-header-cell:active {
            box-shadow: var(--tbl-header-shadow) !important;
        }
        .ag-header-cell {
            color: var(--tbl-header-fg) !important;
            font-family: 'Inter', 'Microsoft YaHei', sans-serif !important;
            font-size: var(--tbl-header-font-size) !important;
            font-weight: var(--tbl-header-font-weight) !important;
            padding: 0 12px !important;
        }
        .ag-header-group-cell {
            color: var(--tbl-header-fg) !important;
            font-family: 'Inter', 'Microsoft YaHei', sans-serif !important;
            font-size: var(--tbl-header-font-size) !important;
            font-weight: var(--tbl-header-font-weight) !important;
        }
        .ag-header-cell .ag-icon,
        .ag-header-group-cell .ag-icon,
        .ag-sort-indicator-icon,
        .ag-icon-asc,
        .ag-icon-desc,
        .ag-icon-menu {
            color: var(--tbl-header-icon) !important;
            fill: var(--tbl-header-icon) !important;
            opacity: 1 !important;
        }
        /* Strict Centering for Header */
        .ag-header-cell-label {
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            text-align: center !important;
            width: 100% !important;
        }
        .ag-header-cell-label, .ag-header-cell-text {
            white-space: normal !important;
            overflow: visible !important;
            text-overflow: clip !important;
            line-height: 1.2 !important;
        }

        /* Strict Centering for Cells */
        .ag-cell, .ag-cell-value {
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            text-align: center !important;
        }
        
        /* Remove default separator bars in header */
        .ag-header-cell::after, .ag-header-group-cell::after {
            display: none !important;
        }

        /* --- 2. Row & Cell Styling --- */
        .ag-row {
            font-family: 'Inter', 'Microsoft YaHei', sans-serif !important;
            font-size: var(--tbl-cell-font-size) !important;
            color: #333333 !important;
            border-bottom-color: #f0f0f0 !important;
        }
        .ag-row-odd {
            background-color: #f8f9fa !important;
        }
        .ag-row-even {
            background-color: #ffffff !important;
        }
        .ag-row-hover {
            background-color: #f0f7ff !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important;
            z-index: 5;
        }
        .ag-row-selected {
            background-color: #e6f7ff !important;
            border-left: 2px solid #4096ff !important; /* Left highlight */
        }
        
        /* Removed duplicate .ag-cell rule, handled above */

        /* Selected Row Text */
        .ag-row-selected .ag-cell {
            font-weight: 500 !important;
        }

        .ag-row.ag-row-pinned,
        .ag-row.ag-row-pinned-bottom {
            background-color: var(--tbl-header-bg) !important;
        }
        .ag-row-pinned .ag-cell,
        .ag-row-pinned-bottom .ag-cell {
            color: var(--tbl-header-fg) !important;
            font-weight: 900 !important;
            border-top: 1px solid var(--tbl-header-border) !important;
        }
        .ag-row-pinned .ag-cell .ag-cell-value,
        .ag-row-pinned-bottom .ag-cell .ag-cell-value {
            color: var(--tbl-header-fg) !important;
            font-weight: 900 !important;
        }
        .ag-row-pinned .ag-cell .ag-icon,
        .ag-row-pinned-bottom .ag-cell .ag-icon {
            color: var(--tbl-header-icon) !important;
            fill: var(--tbl-header-icon) !important;
        }
        
        /* --- 3. Container & Borders --- */
        .ag-root-wrapper {
            border: 1px solid #e5e6eb !important;
            border-radius: 4px !important;
            overflow: hidden !important; /* For radius */
        }
        
        /* --- 4. Scrollbars (Optional, for better look) --- */
        .ag-body-viewport::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        .ag-body-viewport::-webkit-scrollbar-thumb {
            background: #ccc;
            border-radius: 4px;
        }
        .ag-body-viewport::-webkit-scrollbar-track {
            background: #f1f1f1;
        }

        /* --- 5. Mobile Optimization --- */
        @media (max-width: 768px) {
            .ag-header-cell {
                font-size: 13px !important;
                padding: 0 4px !important;
            }
            .ag-header-group-cell {
                font-size: 13px !important;
            }
            .ag-cell {
                font-size: 12px !important;
                padding: 0 4px !important;
            }
            .ag-header-group-cell:active,
            .ag-header-cell:active {
                box-shadow: none !important;
            }
        }
        </style>
    """, unsafe_allow_html=True)

    gb = GridOptionsBuilder.from_dataframe(df)

    percent_cols = set()
    if columns_props:
        for col, props in columns_props.items():
            c_type = (props or {}).get('type')
            if c_type in ['percent', 'growth']:
                percent_cols.add(col)
    for col in df.columns:
        if ('同比' in str(col)) or ('增长' in str(col)) or ('达成率' in str(col)) or (str(col).endswith('率')):
            percent_cols.add(col)
    
    total_row = {c: None for c in df.columns}
    if len(df.columns) > 0:
        total_row[df.columns[0]] = '合计'
    yoy_cols = [c for c in df.columns if ('同比' in str(c)) or (str(c) == '同比增长')]

    for c in df.columns:
        if c == df.columns[0]:
            continue
        if c in percent_cols:
            continue
        s = pd.to_numeric(df[c], errors='coerce')
        if s.notna().sum() == 0:
            continue
        total_row[c] = float(s.fillna(0).sum())

    def _infer_yoy_pair(yoy_col: str):
        if yoy_col not in df.columns:
            return None

        for old, new in [
            ('同比(箱)', '箱数'),
            ('同比（箱）', '箱数'),
            ('同比(门店)', '门店数'),
            ('同比（门店）', '门店数'),
        ]:
            if old in str(yoy_col):
                cur = str(yoy_col).replace(old, new)
                last = str(yoy_col).replace(old, '同期(箱数)' if '箱' in old else '同期(门店数)')
                if cur in df.columns and last in df.columns:
                    return cur, last

        if str(yoy_col) == '同比增长':
            for cur, last in [
                ('本月', '同期'),
                ('本月业绩', '同期业绩'),
                ('本月(万)', '同期(万)'),
                ('本月业绩(万)', '同期业绩(万)'),
                ('实际', '同期'),
            ]:
                if cur in df.columns and last in df.columns:
                    return cur, last

        base = (
            str(yoy_col)
            .replace('同比增长', '')
            .replace('同比', '')
            .replace('增长', '')
            .strip()
        )
        if not base:
            return None
        last_candidates = [c for c in df.columns if ('同期' in str(c) or '去年' in str(c)) and base in str(c)]
        cur_candidates = [c for c in df.columns if ('同期' not in str(c) and '去年' not in str(c) and '同比' not in str(c) and '增长' not in str(c)) and base in str(c)]
        if len(cur_candidates) == 1 and len(last_candidates) == 1:
            return cur_candidates[0], last_candidates[0]
        return None

    for c in yoy_cols:
        pair = _infer_yoy_pair(c)
        if not pair:
            continue
        cur_col, last_col = pair
        try:
            cur_sum = float(pd.to_numeric(df[cur_col], errors='coerce').fillna(0).sum())
            last_sum = float(pd.to_numeric(df[last_col], errors='coerce').fillna(0).sum())
            total_row[c] = (cur_sum - last_sum) / last_sum if last_sum > 0 else None
        except Exception:
            total_row[c] = None
    
    # Configure General Options
    gb.configure_grid_options(
        rowHeight=40, # increased for padding
        headerHeight=60,
        animateRows=True,
        suppressCellFocus=True, # remove blue outline on click
        enableCellTextSelection=True,
        suppressDragLeaveHidesColumns=True,
        sideBar={
            "toolPanels": [
                {
                    "id": "columns",
                    "labelDefault": "列",
                    "iconKey": "columns",
                    "toolPanel": "agColumnsToolPanel",
                    "toolPanelParams": {
                        "suppressRowGroups": True,
                        "suppressValues": True,
                        "suppressPivots": True,
                        "suppressPivotMode": True
                    }
                }
            ],
            "defaultToolPanel": None
        }
    )
    
    # Default Config: Centered, Resizable, Sortable, Filterable
    gb.configure_default_column(
        resizable=True,
        filterable=True,
        sortable=True,
        cellStyle=JS_CENTER,
        headerClass='ag-header-center',
        headerStyle={'textAlign': 'center', 'justifyContent': 'center'},
        wrapHeaderText=True,
        autoHeaderHeight=True,
        minWidth=70,
        flex=1
    )
    
    configured_cols = set()

    # Apply Column Specific Props
    if columns_props:
        for col, props in columns_props.items():
            if col not in df.columns:
                continue
            
            c_type = props.get('type')
            max_value = None
            if c_type in ("bar", "bar_count"):
                s = pd.to_numeric(df[col], errors='coerce')
                max_value = float(s.max()) if len(s) and pd.notna(s.max()) else 0.0
            
            if c_type == 'growth':
                gb.configure_column(col, 
                                    cellStyle=JS_COLOR_CONDITIONAL, 
                                    type=["numericColumn", "numberColumnFilter"], 
                                    valueFormatter=JS_FMT_PCT_RATIO,
                                    minWidth=70,
                                    flex=1)
                configured_cols.add(col)
            elif c_type == 'percent':
                 gb.configure_column(col, 
                                    type=["numericColumn", "numberColumnFilter"], 
                                    valueFormatter=JS_FMT_PCT_RATIO,
                                    minWidth=70,
                                    flex=1)
                 configured_cols.add(col)
            elif c_type == 'money':
                gb.configure_column(col, 
                                    type=["numericColumn", "numberColumnFilter"], 
                                    valueFormatter=JS_FMT_NUM,
                                    minWidth=70,
                                    flex=1)
                configured_cols.add(col)
            elif c_type == 'bar':
                # Use custom renderer
                gb.configure_column(col, 
                                    cellRenderer=JS_PROGRESS_BAR,
                                    cellRendererParams={'maxValue': max_value},
                                    type=["numericColumn", "numberColumnFilter"],
                                    valueFormatter=JS_FMT_NUM,
                                    minWidth=70,
                                    flex=1)
                configured_cols.add(col)
            elif c_type == 'bar_count':
                # Use custom renderer for count
                gb.configure_column(col, 
                                    cellRenderer=JS_PROGRESS_BAR_COUNT,
                                    cellRendererParams={'maxValue': max_value},
                                    type=["numericColumn", "numberColumnFilter"],
                                    valueFormatter=JS_FMT_NUM,
                                    minWidth=70,
                                    flex=1)
                configured_cols.add(col)
                
    # Generic Auto-Type Logic (Fallbacks)
    for col in df.columns:
        if col in configured_cols:
            continue
        
        # Check if column has 'growth' or '同比' -> Growth Color
        if '同比' in col or '增长' in col:
            gb.configure_column(col, 
                                cellStyle=JS_COLOR_CONDITIONAL, 
                                type=["numericColumn", "numberColumnFilter"], 
                                valueFormatter=JS_FMT_PCT_RATIO,
                                minWidth=70,
                                flex=1)
        
        # Check if '达成率' or '率' -> Percent
        elif '达成率' in col or '占比' in col or str(col).endswith('率'):
            gb.configure_column(col, 
                                type=["numericColumn", "numberColumnFilter"], 
                                valueFormatter=JS_FMT_PCT_RATIO,
                                minWidth=70,
                                flex=1)
            
            # Optional: Add Data Bar style for '达成率' if requested
            if '达成率' in col:
                 gb.configure_column(col,
                    cellStyle=JsCode("""
                        function(params) {
                            let ratio = params.value;
                            if (ratio === null || isNaN(ratio)) return {'textAlign': 'center'};
                            let percent = ratio * 100;
                             let color = '#28a745'; // Green
                             if (percent < 100) color = '#ffc107'; // Yellow
                             if (percent < 60) color = '#dc3545'; // Red
                             return {
                                 'textAlign': 'center', 
                                 'background': `linear-gradient(90deg, ${color} ${Math.min(percent, 100)}%, transparent ${Math.min(percent, 100)}%)`
                             };
                        }
                    """),
                    valueFormatter=JS_FMT_PCT_RATIO,
                    minWidth=70,
                    flex=1
                 )

        # Money/Number
        elif pd.api.types.is_numeric_dtype(df[col]):
            gb.configure_column(col, 
                                type=["numericColumn", "numberColumnFilter"], 
                                valueFormatter=JS_FMT_NUM,
                                minWidth=70,
                                flex=1)
        else:
            if col == df.columns[0]:
                gb.configure_column(col, minWidth=95, flex=1.2, tooltipField=col)
            else:
                gb.configure_column(col, minWidth=100, flex=1.2, tooltipField=col)

    # Selection
    if on_row_selected:
        gb.configure_selection('single', use_checkbox=False)
        
    gridOptions = gb.build()
    gridOptions['pinnedBottomRowData'] = [total_row]
    if column_defs:
        gridOptions['columnDefs'] = column_defs
        gridOptions['groupHeaderHeight'] = 40
        gridOptions['headerHeight'] = 46
    if grid_options_overrides:
        gridOptions.update(grid_options_overrides)
    
    # --- Auto Height & Pagination Logic ---
    # 1. Calculate ideal height for all rows
    n_rows = len(df)
    row_h = 40  # consistent with configure_grid_options rowHeight
    header_h = 60 # consistent with configure_grid_options headerHeight
    padding = 20
    
    calc_full_height = header_h + (n_rows * row_h) + padding + 40 # +40 buffer for potential horizontal scrollbar/total row
    
    # 2. Thresholds
    MAX_HEIGHT_NO_SCROLL = 600  # If content < 600px, show full height (no scroll/pagination)
    PAGE_SIZE = 20              # If content > 600px, use pagination with 20 rows/page
    
    # 3. Determine Mode
    # If explicit height provided, use it (and scroll if needed)
    # Else, apply auto-logic
    if height:
        final_height = height
        # If explicitly short height, maybe enable pagination? No, trust caller or AgGrid default scroll.
    else:
        if calc_full_height <= MAX_HEIGHT_NO_SCROLL:
            final_height = max(150, calc_full_height) # At least 150px
            # No pagination needed
            gridOptions['pagination'] = False
        else:
            # Content too long -> Use Pagination
            gridOptions['pagination'] = True
            gridOptions['paginationPageSize'] = PAGE_SIZE
            # Height fits PageSize rows + Header + PaginationPanel
            # PageSize * RowHeight + Header + PagerPanel(~50px)
            final_height = (PAGE_SIZE * row_h) + header_h + 50 + padding
    
    # Enable SideBar for Columns Tool Panel (Optional, user asked for "Drop down menu for each column")
    # AgGrid default filter menu is on column header. 
    
    # --- Responsive & Horizontal Scroll Logic ---
    # If too many columns, disable 'fit_columns_on_grid_load' to allow horizontal scroll
    # Heuristic: > 8-10 columns or if we suspect wide content
    should_fit_columns = True
    if len(df.columns) > 10:
        should_fit_columns = False
    
    return AgGrid(
        df,
        gridOptions=gridOptions,
        height=final_height,
        width='100%',
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.SELECTION_CHANGED | GridUpdateMode.VALUE_CHANGED,
        fit_columns_on_grid_load=should_fit_columns,
        allow_unsafe_jscode=True, 
        theme='streamlit', 
        key=key
    )

# -----------------------------------------------------------------------------
# 3. Data Logic
# -----------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_data_v2(file_bytes: bytes, file_name: str):
    debug_logs = []
    try:
        file_name_lower = (file_name or "").lower()
        bio = io.BytesIO(file_bytes)
        if file_name_lower.endswith('.csv'):
            df = pd.read_csv(bio, encoding='gb18030')
            df_stock = None
            df_q4_raw = None
            df_perf_raw = None
        else:
            xl = pd.ExcelFile(bio)
            df = xl.parse(0)
            df_stock = xl.parse(1) if len(xl.sheet_names) > 1 else None
            df_q4_raw = xl.parse(2) if len(xl.sheet_names) > 2 else None
            df_perf_raw = None
            
            debug_logs.append(f"Total Sheets: {len(xl.sheet_names)} | Names: {xl.sheet_names}")

            # Sheet 4 Detection Logic (Robust)
            if len(xl.sheet_names) > 3:
                preferred = next((s for s in xl.sheet_names if 'sheet4' in str(s).strip().lower()), None)
                candidate_names = [preferred] if preferred else []
                candidate_names += [s for s in xl.sheet_names if s not in candidate_names]
                
                for sname in candidate_names:
                    try:
                        # Optimization: Read only header first (0 rows) to check columns
                        tmp_header = xl.parse(sname, nrows=0)
                        cols = [str(c).strip() for c in tmp_header.columns]
                    except Exception as e:
                        debug_logs.append(f"Error parsing header of {sname}: {str(e)}")
                        continue
                    
                    # Fuzzy match for keys
                    key_hits = sum(1 for k in ['年份', '月份', '省区'] if any(k in c for c in cols))
                    signal_hits = sum(1 for k in ['发货仓', '原价金额', '基本数量', '大分类', '月分析', '客户简称'] if any(k in c for c in cols))
                    
                    debug_logs.append(f"Checking '{sname}': keys={key_hits}, signals={signal_hits}")
                    
                    if key_hits >= 2 and signal_hits >= 1:
                        # Found it! Now read the full sheet
                        try:
                            df_perf_raw = xl.parse(sname)
                            debug_logs.append(f"-> MATCHED Sheet4: {sname}")
                            break
                        except Exception as e:
                            debug_logs.append(f"Error reading body of {sname}: {e}")
            else:
                 debug_logs.append("Warning: Less than 4 sheets found.")
            
        # --- Process Sheet 1 (Sales) ---
        # Ensure column names are clean
        df.columns = [str(c).strip() for c in df.columns]
        
        # --- Handle Long Format (Rows) -> Wide Format (Columns) ---
        # User indicates Time (Month) is in Column F (index 5)
        # Potential Columns: F=Time, I=Prov, J=Dist, K=Qty (based on user info)
        is_long_format = False
        time_col = None
        
        # Check if Column F exists and looks like Month
        if len(df.columns) > 5:
            col_f = df.columns[5]
            # Check a sample of values in Col F for "月" or date-like
            sample_vals = df[col_f].dropna().head(10).astype(str).tolist()
            if any('月' in v for v in sample_vals):
                is_long_format = True
                time_col = col_f
        
        if is_long_format:
            # Identify Key Columns for Pivot
            # Try to map by name or index
            # User hints: Prov(I=8), Dist(J=9), Qty(K=10)
            
            col_prov = df.columns[8] if len(df.columns) > 8 else None
            col_dist = df.columns[9] if len(df.columns) > 9 else None
            col_qty = df.columns[10] if len(df.columns) > 10 else None
            
            # Fallback: Search by name
            if col_prov is None: col_prov = next((c for c in df.columns if '省' in c), None)
            if col_dist is None: col_dist = next((c for c in df.columns if '经销' in c or '客户' in c), None)
            if col_qty is None: col_qty = next((c for c in df.columns if '数' in c or 'Qty' in c or '箱' in c), None)
            
            # Store Column? If not found, default to Dist or blank
            col_store = next((c for c in df.columns if '门店' in c), None)
            
            if col_prov and col_dist and col_qty and time_col:
                # Prepare for Pivot
                pivot_index = [col_prov, col_dist]
                if col_store:
                    pivot_index.append(col_store)
                
                # Pivot
                # Ensure Qty is numeric
                df[col_qty] = pd.to_numeric(df[col_qty], errors='coerce').fillna(0)
                
                df_wide = df.pivot_table(
                    index=pivot_index,
                    columns=time_col,
                    values=col_qty,
                    aggfunc='sum'
                ).reset_index()
                
                # Handle Missing Store Column if needed
                if not col_store:
                    df_wide['门店名称'] = df_wide[col_dist] # Use Dist as Store if missing
                    
                df = df_wide
                # Reset clean columns
                df.columns = [str(c).strip() for c in df.columns]
                
        # Identify Month Columns (Assume '1月', '2月', etc. or columns 4 onwards if strict structure)
        # Based on user requirement: Col 1-3 are dimensions, 4+ are months.
        # Let's try to detect "X月" pattern first, fallback to index.
        month_cols = [c for c in df.columns if '月' in c and c not in ['品牌省区名称', '经销商名称', '门店名称']]
        
        # If headers are standard: 品牌省区名称, 经销商名称, 门店名称
        # Normalize dimension columns
        rename_map = {}
        if '品牌省区名称' in df.columns: rename_map['品牌省区名称'] = '省区'
        if '经销商名称' not in df.columns and len(df.columns) > 1: rename_map[df.columns[1]] = '经销商名称'
        if '门店名称' not in df.columns and len(df.columns) > 2: rename_map[df.columns[2]] = '门店名称'
        
        df = df.rename(columns=rename_map)
        
        # Validate critical columns
        required = ['省区', '经销商名称', '门店名称']
        for req in required:
            if req not in df.columns:
                # Fallback: Assume positional 0, 1, 2
                if len(df.columns) >= 3:
                    df.columns.values[0] = '省区'
                    df.columns.values[1] = '经销商名称'
                    df.columns.values[2] = '门店名称'
                else:
                    st.error(f"数据格式错误：缺失列 {req}")
                    return None, None, None, None

        # Re-identify month cols after rename
        month_cols = [c for c in df.columns if c not in required]
        
        # Ensure numeric
        for col in month_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
        # --- Core Metric Calculation ---
        # 1. Total Shipment
        df['总出库数'] = df[month_cols].sum(axis=1)
        
        # 2. Effective Months (Count where Shipment > 0)
        df['有效月份数'] = df[month_cols].gt(0).sum(axis=1).astype(int)
        
        # 3. Avg Monthly Shipment
        # Optimized: Vectorized calculation instead of apply
        df['平均每月出库数'] = np.where(df['有效月份数'] > 0, df['总出库数'] / df['有效月份数'], 0.0)
        
        # Classification
        # Optimized: Vectorized select
        conditions = [
            df['平均每月出库数'] >= 4,
            (df['平均每月出库数'] >= 2) & (df['平均每月出库数'] < 4),
            (df['平均每月出库数'] >= 1) & (df['平均每月出库数'] < 2)
        ]
        choices = ['A类门店 (>=4)', 'B类门店 (2-4)', 'C类门店 (1-2)']
        df['门店分类'] = np.select(conditions, choices, default='D类门店 (<1)')
        
        # --- Process Sheet 2 (Stock) ---
        if df_stock is not None:
            # Clean columns
            df_stock.columns = [str(c).strip() for c in df_stock.columns]
            
            # Validate Stock Columns (A-L strict structure check or Name check)
            # User defined: 经销商编码(A), 经销商名称(B), 产品编码(C), 产品名称(D), 库存数量(E), 箱数(F), 省区(G), 客户简称(H), 大类(I), 小类(J), 重量(K), 规格(L)
            # We map by index to be safe if names vary slightly, or by expected names.
            # Let's use expected names map based on index to standardize.
            # UPDATE: Use '客户简称' (H列, index 7) as the primary '经销商名称' for analysis.
            # Rename original '经销商名称' (B列, index 1) to '经销商全称' for reference.
            stock_cols_map = {
                0: '经销商编码', 1: '经销商全称', 2: '产品编码', 3: '产品名称', 
                4: '库存数量(听/盒)', 5: '箱数', 6: '省区名称', 7: '经销商名称', # Map '客户简称' to '经销商名称'
                8: '产品大类', 9: '产品小类', 10: '重量', 11: '规格'
            }
            
            if len(df_stock.columns) >= 12:
                # Rename columns by index to ensure standard access
                new_cols = list(df_stock.columns)
                for idx, name in stock_cols_map.items():
                    new_cols[idx] = name
                df_stock.columns = new_cols
                
                # Ensure numeric '箱数'
                df_stock['箱数'] = pd.to_numeric(df_stock['箱数'], errors='coerce').fillna(0)
                
                # Clean Distributor Name (客户简称)
                df_stock['经销商名称'] = df_stock['经销商名称'].astype(str).str.strip()
                
                # Fix PyArrow mixed type error for mixed columns
                df_stock['重量'] = df_stock['重量'].astype(str)
                df_stock['规格'] = df_stock['规格'].astype(str)
                
                # --- Smart Classification Logic (Specific Category) ---
                # Rules:
                # - 雅系列：仅当产品名称包含「雅赋/雅耀/雅舒/雅护」之一时命中
                # - 分段：仅在「产品大类=美思雅段粉」范围内，且产品名称包含「1段/2段/3段」之一时命中
                
                # Optimized: Vectorized Logic using np.select and str.contains
                # Pre-calculate boolean masks
                name_series = df_stock['产品名称'].astype(str)
                cat_series = df_stock['产品大类'].astype(str)
                
                mask_ya = name_series.str.contains('雅赋|雅耀|雅舒|雅护', regex=True)
                mask_seg_cat = cat_series == '美思雅段粉'
                
                # For segments, we need to extract which segment it is. 
                # Since we need the specific string ('1段' etc), np.select is good but we need to know WHICH one.
                # Let's use extraction for segments.
                seg_extract = name_series.str.extract(r'(1段|2段|3段)')[0]
                
                # Logic:
                # 1. If '雅系列' keyword -> return keyword. (Need to extract which one? Old logic returned the keyword itself e.g. '雅赋')
                # 2. If '美思雅段粉' and has segment -> return segment.
                # 3. Else '其他'
                
                # Extract Ya keyword
                ya_extract = name_series.str.extract(r'(雅赋|雅耀|雅舒|雅护)')[0]
                
                # Construct final series
                # Priority: Ya > Segment (if logic follows original sequence, Ya was checked first)
                
                df_stock['具体分类'] = np.where(
                    mask_ya, ya_extract,
                    np.where(
                        mask_seg_cat & seg_extract.notna(), seg_extract,
                        '其他'
                    )
                )
                df_stock['具体分类'] = df_stock['具体分类'].fillna('其他').astype(str)
                 
                # --- Filter Stock Data (Hardcoded Rules) ---
                # Rule 1: Weight (重量) must be '700', '800', '800-新包装'
                if '重量' in df_stock.columns:
                    valid_weights = ['700', '800', '800-新包装']
                    # Ensure weight column is string for comparison (already done above)
                    # Handle potential float/int like 700.0 or 700
                    # We converted to string, so 700 might become '700' or '700.0' depending on source.
                    # Let's normalize: check if string contains the valid weight or exact match.
                    # Exact match is safer if data is clean. Let's try exact match first, assuming '700' in Excel is '700' or 700.
                    # If it was 700 (int), astype(str) makes it '700'.
                    df_stock = df_stock[df_stock['重量'].isin(valid_weights)]
            else:
                st.warning("库存表 (Sheet2) 列数不足 12 列，无法进行库存分析。")
                df_stock = None

        # --- Process Sheet 3 (Outbound Base Table) ---
        if df_q4_raw is not None:
            df_q4_raw.columns = [str(c).strip() for c in df_q4_raw.columns]

            df_out = df_q4_raw.copy()

            month_src = df_out.columns[5] if len(df_out.columns) > 5 else None
            prov_src = df_out.columns[8] if len(df_out.columns) > 8 else None
            dist_src = df_out.columns[9] if len(df_out.columns) > 9 else None
            qty_src = df_out.columns[10] if len(df_out.columns) > 10 else None

            rename_map = {}
            if month_src: rename_map[month_src] = '月份'
            if prov_src: rename_map[prov_src] = '省区'
            if dist_src: rename_map[dist_src] = '经销商名称'
            if qty_src: rename_map[qty_src] = '数量(箱)'

            cat_src = next((c for c in df_out.columns if '产品大类' in str(c)), None)
            if cat_src is None:
                cat_src = next((c for c in df_out.columns if ('大类' in str(c)) and ('省区' not in str(c))), None)
            sub_src = next((c for c in df_out.columns if '产品小类' in str(c)), None)
            if sub_src is None:
                sub_src = next((c for c in df_out.columns if ('小类' in str(c)) and ('产品' in str(c))), None)
            if cat_src is None and len(df_out.columns) > 11:
                cat_src = df_out.columns[11]
            if sub_src is None and len(df_out.columns) > 12:
                sub_src = df_out.columns[12]

            if cat_src: rename_map[cat_src] = '产品大类'
            if sub_src: rename_map[sub_src] = '产品小类'

            df_out = df_out.rename(columns=rename_map)
            df_out = df_out.loc[:, ~df_out.columns.duplicated()]

            if '经销商名称' in df_out.columns:
                df_out['经销商名称'] = df_out['经销商名称'].astype(str).str.strip()
            if '数量(箱)' in df_out.columns:
                df_out['数量(箱)'] = pd.to_numeric(df_out['数量(箱)'], errors='coerce').fillna(0)
            if '产品大类' in df_out.columns:
                df_out['产品大类'] = df_out['产品大类'].astype(str).str.strip()
            if '产品小类' in df_out.columns:
                df_out['产品小类'] = df_out['产品小类'].astype(str).str.strip()

            df_q4_raw = df_out

        # --- Process Sheet 4 (Performance / Shipment) ---
        if df_perf_raw is not None:
            df_perf_raw.columns = [str(c).strip() for c in df_perf_raw.columns]
            df_perf = df_perf_raw.copy()

            col_year = next((c for c in df_perf.columns if c == '年份' or '年' in c), None)
            col_month = next((c for c in df_perf.columns if c == '月份' or '月' in c), None)
            col_prov = next((c for c in df_perf.columns if c == '省区' or '省区' in c), None)
            col_dist = next((c for c in df_perf.columns if c == '经销商名称' or c == '客户简称' or '客户简称' in c), None)
            col_qty = next((c for c in df_perf.columns if c == '箱数' or c == '基本数量' or '数量' in c), None)
            col_amt = next((c for c in df_perf.columns if c == '发货金额' or c == '原价金额' or '金额' in c), None)
            col_wh = next((c for c in df_perf.columns if c == '发货仓' or '发货仓' in c), None)
            col_mid = next((c for c in df_perf.columns if c == '中类' or '中类' in c), None)
            col_grp = next((c for c in df_perf.columns if c == '归类' or '归类' in c), None)
            col_bigcat = next((c for c in df_perf.columns if c == '大分类' or '大分类' in c), None)
            col_big = next((c for c in df_perf.columns if c == '大类' or '大类' in c), None)
            col_small = next((c for c in df_perf.columns if c == '小类' or '小类' in c), None)
            col_cat = next((c for c in df_perf.columns if c == '月分析' or '月分析' in c), None)

            rename_perf = {}
            if col_year: rename_perf[col_year] = '年份'
            if col_month: rename_perf[col_month] = '月份'
            if col_prov: rename_perf[col_prov] = '省区'
            if col_dist: rename_perf[col_dist] = '经销商名称'
            if col_qty: rename_perf[col_qty] = '发货箱数'
            if col_amt: rename_perf[col_amt] = '发货金额'
            if col_wh: rename_perf[col_wh] = '发货仓'
            if col_mid: rename_perf[col_mid] = '中类'
            if col_grp: rename_perf[col_grp] = '归类'
            if col_bigcat:
                rename_perf[col_bigcat] = '大分类'
            elif col_cat:
                rename_perf[col_cat] = '大分类'
            if col_big: rename_perf[col_big] = '大类'
            if col_small: rename_perf[col_small] = '小类'

            df_perf = df_perf.rename(columns=rename_perf)

            for c in ['省区', '经销商名称', '发货仓', '中类', '归类', '大分类', '大类', '小类']:
                if c in df_perf.columns:
                    df_perf[c] = df_perf[c].fillna('').astype(str).str.strip()
            
            # --- FIX: Ensure '经销商名称' exists ---
            if '经销商名称' not in df_perf.columns:
                # Try to find alias
                alt_dist = next((c for c in df_perf.columns if '客户' in c or '经销' in c), None)
                if alt_dist:
                    df_perf = df_perf.rename(columns={alt_dist: '经销商名称'})
                else:
                    # Fallback: Create empty if absolutely necessary (but better to warn)
                    df_perf['经销商名称'] = '未知经销商'
            # --------------------------------------

            if '大分类' in df_perf.columns and '类目' not in df_perf.columns:
                df_perf['类目'] = df_perf['大分类']

            if '年份' in df_perf.columns:
                # Handle "25年" or "2025" strings by extracting digits
                # NOTE: Use regex extraction to handle "25年" -> "25"
                df_perf['年份'] = df_perf['年份'].astype(str).str.extract(r'(\d+)')[0].astype(float).fillna(0).astype(int)
                # Normalize 2-digit years to 4-digit (e.g. 25 -> 2025)
                df_perf['年份'] = df_perf['年份'].apply(lambda y: y + 2000 if 0 < y < 100 else y)

            if '月份' in df_perf.columns:
                 # Handle "1月" or "01" strings
                df_perf['月份'] = df_perf['月份'].astype(str).str.extract(r'(\d+)')[0].astype(float).fillna(0).astype(int)
            if '发货箱数' in df_perf.columns:
                df_perf['发货箱数'] = pd.to_numeric(df_perf['发货箱数'], errors='coerce').fillna(0)
            if '发货金额' in df_perf.columns:
                df_perf['发货金额'] = pd.to_numeric(df_perf['发货金额'], errors='coerce').fillna(0)

            if '年份' in df_perf.columns and '月份' in df_perf.columns:
                df_perf = df_perf[(df_perf['年份'] > 0) & (df_perf['月份'].between(1, 12))]
                df_perf['年月'] = pd.to_datetime(df_perf['年份'].astype(str) + '-' + df_perf['月份'].astype(str).str.zfill(2) + '-01')
            else:
                df_perf['年月'] = pd.NaT

            df_perf_raw = df_perf

        # --- Process Sheet 5 (Target) ---
        df_target_raw = None
        try:
            if len(xl.sheet_names) > 4:
                df_target_raw = xl.parse(4)
                df_target_raw.columns = [str(c).strip() for c in df_target_raw.columns]
                
                # Expected Cols: D(品类), E(月份), F(任务量) -> Index 3, 4, 5
                # Rename by index to be safe
                rename_target = {}
                if len(df_target_raw.columns) > 3: rename_target[df_target_raw.columns[3]] = '品类'
                if len(df_target_raw.columns) > 4: rename_target[df_target_raw.columns[4]] = '月份'
                if len(df_target_raw.columns) > 5: rename_target[df_target_raw.columns[5]] = '任务量'
                
                df_target_raw = df_target_raw.rename(columns=rename_target)
                
                # Basic Cleaning
                if '月份' in df_target_raw.columns:
                     # Handle "1月" or "01" strings
                    df_target_raw['月份'] = df_target_raw['月份'].astype(str).str.extract(r'(\d+)')[0].astype(float).fillna(0).astype(int)
                if '任务量' in df_target_raw.columns:
                    df_target_raw['任务量'] = pd.to_numeric(df_target_raw['任务量'], errors='coerce').fillna(0)
            else:
                 debug_logs.append("Warning: Sheet5 (Target) not found.")
        except Exception as e:
            debug_logs.append(f"Error parsing Sheet5: {e}")
            df_target_raw = None

        return df, month_cols, df_stock, df_q4_raw, df_perf_raw, df_target_raw, debug_logs
        
    except Exception as e:
        st.error(f"数据加载失败: {str(e)}")
        return None, None, None, None, None, None, [str(e)]

@st.cache_data(ttl=3600)
def load_data_v3(file_bytes: bytes, file_name: str):
    debug_logs = []
    try:
        file_name_lower = (file_name or "").lower()
        bio = io.BytesIO(file_bytes)
        
        # Init Returns
        df = None
        month_cols = []
        df_stock = None
        df_q4_raw = None
        df_perf_raw = None
        df_target_raw = None
        df_scan_raw = None

        if file_name_lower.endswith('.csv'):
            df = pd.read_csv(bio, encoding='gb18030')
        else:
            xl = pd.ExcelFile(bio)
            debug_logs.append(f"Sheet Names: {xl.sheet_names}")
            
            # Sheet 1: Sales
            if len(xl.sheet_names) > 0: df = xl.parse(0)
            
            # Sheet 2: Stock
            if len(xl.sheet_names) > 1: df_stock = xl.parse(1)
            
            # Sheet 3: Outbound (Q4)
            if len(xl.sheet_names) > 2: df_q4_raw = xl.parse(2)
            
            # Sheet 4: Performance
            if len(xl.sheet_names) > 3:
                preferred = next((s for s in xl.sheet_names if 'sheet4' in str(s).strip().lower()), None)
                candidate_names = [preferred] if preferred else []
                candidate_names += [s for s in xl.sheet_names if s not in candidate_names]
                for sname in candidate_names:
                    try:
                        tmp_header = xl.parse(sname, nrows=0)
                        cols = [str(c).strip() for c in tmp_header.columns]
                        key_hits = sum(1 for k in ['年份', '月份', '省区'] if any(k in c for c in cols))
                        signal_hits = sum(1 for k in ['发货仓', '原价金额', '基本数量', '大分类', '月分析', '客户简称'] if any(k in c for c in cols))
                        if key_hits >= 2 and signal_hits >= 1:
                            df_perf_raw = xl.parse(sname)
                            debug_logs.append(f"-> MATCHED Sheet4: {sname}")
                            break
                    except: continue
            
            # Sheet 5: Target
            if len(xl.sheet_names) > 4: df_target_raw = xl.parse(4)

            # Sheet 6: Scan Data
            if len(xl.sheet_names) > 5: df_scan_raw = xl.parse(5)

        # --- Process Sheet 1 (Sales) ---
        if df is not None:
            df.columns = [str(c).strip() for c in df.columns]
            
            # Identify Month Columns
            is_long_format = False
            time_col = None
            if len(df.columns) > 5:
                col_f = df.columns[5]
                sample_vals = df[col_f].dropna().head(10).astype(str).tolist()
                if any('月' in v for v in sample_vals):
                    is_long_format = True
                    time_col = col_f
            
            if is_long_format:
                col_prov = df.columns[8] if len(df.columns) > 8 else None
                col_dist = df.columns[9] if len(df.columns) > 9 else None
                col_qty = df.columns[10] if len(df.columns) > 10 else None
                
                if col_prov is None: col_prov = next((c for c in df.columns if '省' in c), None)
                if col_dist is None: col_dist = next((c for c in df.columns if '经销' in c or '客户' in c), None)
                if col_qty is None: col_qty = next((c for c in df.columns if '数' in c or 'Qty' in c or '箱' in c), None)
                col_store = next((c for c in df.columns if '门店' in c), None)
                
                if col_prov and col_dist and col_qty and time_col:
                    df[col_qty] = pd.to_numeric(df[col_qty], errors='coerce').fillna(0)
                    pivot_index = [col_prov, col_dist]
                    if col_store: pivot_index.append(col_store)
                    df_wide = df.pivot_table(index=pivot_index, columns=time_col, values=col_qty, aggfunc='sum').reset_index()
                    if not col_store: df_wide['门店名称'] = df_wide[col_dist]
                    df = df_wide
                    df.columns = [str(c).strip() for c in df.columns]
            
            rename_map = {}
            if '品牌省区名称' in df.columns: rename_map['品牌省区名称'] = '省区'
            if '经销商名称' not in df.columns and len(df.columns) > 1: rename_map[df.columns[1]] = '经销商名称'
            if '门店名称' not in df.columns and len(df.columns) > 2: rename_map[df.columns[2]] = '门店名称'
            df = df.rename(columns=rename_map)
            
            required = ['省区', '经销商名称', '门店名称']
            for req in required:
                if req not in df.columns:
                    if len(df.columns) >= 3:
                        df.columns.values[0] = '省区'
                        df.columns.values[1] = '经销商名称'
                        df.columns.values[2] = '门店名称'
            
            month_cols = [c for c in df.columns if '月' in c and c not in required]
            for col in month_cols:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
            df['总出库数'] = df[month_cols].sum(axis=1)
            df['有效月份数'] = df[month_cols].gt(0).sum(axis=1).astype(int)
            df['平均每月出库数'] = np.where(df['有效月份数'] > 0, df['总出库数'] / df['有效月份数'], 0.0)
            
            conditions = [df['平均每月出库数'] >= 4, (df['平均每月出库数'] >= 2) & (df['平均每月出库数'] < 4), (df['平均每月出库数'] >= 1) & (df['平均每月出库数'] < 2)]
            choices = ['A类门店 (>=4)', 'B类门店 (2-4)', 'C类门店 (1-2)']
            df['门店分类'] = np.select(conditions, choices, default='D类门店 (<1)')

        # --- Process Sheet 2 (Stock) ---
        if df_stock is not None:
            df_stock.columns = [str(c).strip() for c in df_stock.columns]
            stock_cols_map = {
                0: '经销商编码', 1: '经销商全称', 2: '产品编码', 3: '产品名称', 
                4: '库存数量(听/盒)', 5: '箱数', 6: '省区名称', 7: '经销商名称', # 7=客户简称
                8: '产品大类', 9: '产品小类', 10: '重量', 11: '规格'
            }
            if len(df_stock.columns) >= 12:
                new_cols = list(df_stock.columns)
                for idx, name in stock_cols_map.items():
                    if idx < len(new_cols): new_cols[idx] = name
                df_stock.columns = new_cols
                df_stock['箱数'] = pd.to_numeric(df_stock['箱数'], errors='coerce').fillna(0)
                
                # CLEAN DISTRIBUTOR NAME STRICTLY
                if '经销商名称' in df_stock.columns:
                    df_stock['经销商名称'] = df_stock['经销商名称'].astype(str).str.replace(r'\s+', '', regex=True)
                
                df_stock['重量'] = df_stock['重量'].astype(str)
                df_stock['规格'] = df_stock['规格'].astype(str)
                
                name_series = df_stock['产品名称'].astype(str)
                mask_ya = name_series.str.contains('雅赋|雅耀|雅舒|雅护', regex=True)
                mask_seg_cat = df_stock['产品大类'].astype(str) == '美思雅段粉'
                seg_extract = name_series.str.extract(r'(1段|2段|3段)')[0]
                ya_extract = name_series.str.extract(r'(雅赋|雅耀|雅舒|雅护)')[0]
                
                df_stock['具体分类'] = np.where(mask_ya, ya_extract, np.where(mask_seg_cat & seg_extract.notna(), seg_extract, '其他'))
                df_stock['具体分类'] = df_stock['具体分类'].fillna('其他').astype(str)
                 
                if '重量' in df_stock.columns:
                    valid_weights = ['700', '800', '800-新包装']
                    df_stock = df_stock[df_stock['重量'].isin(valid_weights)]
            else:
                df_stock = None

        # --- Process Sheet 3 (Outbound) FIX ---
        if df_q4_raw is not None:
            # Deduplicate
            cols = pd.Series(df_q4_raw.columns)
            for dup in cols[cols.duplicated()].unique(): 
                cols[cols[cols == dup].index.values.tolist()] = [dup + '.' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
            df_q4_raw.columns = cols
            
            df_out = df_q4_raw.copy()
            
            # Map Indices (User Requirement: +8 shift)
            # M(12)=Year, N(13)=Month, Q(16)=Prov, R(17)=Dist(CustomerAbbr), S(18)=Qty, U(20)=SubCat
            idx_map = {
                12: '年份',
                13: '月份',
                16: '省区',
                17: '经销商名称',
                18: '数量(箱)',
                20: '产品小类',
                19: '产品大类'
            }
            curr_cols = list(df_out.columns)
            
            # Avoid Name Collision: Rename existing columns that clash with target names
            target_names = list(idx_map.values())
            for i, c in enumerate(curr_cols):
                if c in target_names and i not in idx_map:
                    new_n = f"{c}_old_{i}"
                    df_out.rename(columns={c: new_n}, inplace=True)
                    debug_logs.append(f"Renamed collision '{c}' -> '{new_n}'")
            
            # Refresh columns after collision avoidance
            curr_cols = list(df_out.columns)
            
            for idx, name in idx_map.items():
                if idx < len(curr_cols):
                    df_out.rename(columns={curr_cols[idx]: name}, inplace=True)
            
            # Clean Dist Name
            if '经销商名称' in df_out.columns:
                 df_out['经销商名称'] = df_out['经销商名称'].astype(str).str.replace(r'\s+', '', regex=True)
                 debug_logs.append(f"Sheet3 Dist Sample: {df_out['经销商名称'].head(3).tolist()}")

            if '数量(箱)' in df_out.columns:
                df_out['数量(箱)'] = pd.to_numeric(df_out['数量(箱)'], errors='coerce').fillna(0)
            
            if '产品大类' in df_out.columns: df_out['产品大类'] = df_out['产品大类'].astype(str).str.strip()
            if '产品小类' in df_out.columns: df_out['产品小类'] = df_out['产品小类'].astype(str).str.strip()
            
            # Clean Year
            if '年份' in df_out.columns:
                # Extract digits and normalize
                df_out['年份'] = df_out['年份'].astype(str).str.extract(r'(\d+)')[0].astype(float).fillna(0).astype(int)
                # Normalize 25 -> 2025
                df_out['年份'] = df_out['年份'].apply(lambda y: y + 2000 if 20 <= y < 100 else y)

            df_q4_raw = df_out

        # --- Process Sheet 4 (Perf) ---
        if df_perf_raw is not None:
            df_perf_raw.columns = [str(c).strip() for c in df_perf_raw.columns]
            df_perf = df_perf_raw.copy()
            col_year = next((c for c in df_perf.columns if str(c).strip() == '年份' or ('年' in str(c))), None)
            col_month = next((c for c in df_perf.columns if str(c).strip() == '月份' or ('月' in str(c))), None)
            col_prov = next((c for c in df_perf.columns if str(c).strip() == '省区' or ('省区' in str(c)) or (str(c).strip() == '省')), None)
            col_dist = next((c for c in df_perf.columns if str(c).strip() == '经销商名称' or str(c).strip() == '客户简称' or ('客户简称' in str(c)) or ('经销商' in str(c))), None)
            col_qty = next((c for c in df_perf.columns if str(c).strip() == '发货箱数' or str(c).strip() == '箱数' or str(c).strip() == '基本数量' or ('数量' in str(c)) or ('箱' in str(c))), None)
            col_amt = next((c for c in df_perf.columns if str(c).strip() == '发货金额' or str(c).strip() == '原价金额' or ('金额' in str(c))), None)
            col_wh = next((c for c in df_perf.columns if str(c).strip() == '发货仓' or ('发货仓' in str(c))), None)
            col_mid = next((c for c in df_perf.columns if str(c).strip() == '中类' or ('中类' in str(c))), None)
            col_grp = next((c for c in df_perf.columns if str(c).strip() == '归类' or ('归类' in str(c))), None)
            col_bigcat = next((c for c in df_perf.columns if str(c).strip() == '大分类' or ('大分类' in str(c))), None)
            col_big = next((c for c in df_perf.columns if str(c).strip() == '大类' or ('大类' in str(c))), None)
            col_small = next((c for c in df_perf.columns if str(c).strip() == '小类' or ('小类' in str(c))), None)
            col_cat = next((c for c in df_perf.columns if str(c).strip() == '月分析' or ('月分析' in str(c))), None)

            rename_perf = {}
            if col_year: rename_perf[col_year] = '年份'
            if col_month: rename_perf[col_month] = '月份'
            if col_prov: rename_perf[col_prov] = '省区'
            if col_dist: rename_perf[col_dist] = '经销商名称'
            if col_qty: rename_perf[col_qty] = '发货箱数'
            if col_amt: rename_perf[col_amt] = '发货金额'
            if col_wh: rename_perf[col_wh] = '发货仓'
            if col_mid: rename_perf[col_mid] = '中类'
            if col_grp: rename_perf[col_grp] = '归类'
            if col_bigcat:
                rename_perf[col_bigcat] = '大分类'
            elif col_cat:
                rename_perf[col_cat] = '大分类'
            if col_big: rename_perf[col_big] = '大类'
            if col_small: rename_perf[col_small] = '小类'

            df_perf = df_perf.rename(columns=rename_perf)

            if '经销商名称' not in df_perf.columns:
                alt_dist = next((c for c in df_perf.columns if ('客户' in str(c)) or ('经销' in str(c))), None)
                if alt_dist:
                    df_perf = df_perf.rename(columns={alt_dist: '经销商名称'})
                else:
                    df_perf['经销商名称'] = ''
                    debug_logs.append("Warning: Sheet4 missing distributor column; set '经销商名称' to empty.")

            for c in ['省区', '经销商名称', '发货仓', '中类', '归类', '大分类', '大类', '小类']:
                if c in df_perf.columns:
                    df_perf[c] = df_perf[c].fillna('').astype(str).str.strip()

            if '年份' in df_perf.columns:
                df_perf['年份'] = df_perf['年份'].astype(str).str.extract(r'(\d+)')[0].astype(float).fillna(0).astype(int)
                df_perf['年份'] = df_perf['年份'].apply(lambda y: y + 2000 if 0 < y < 100 else y)
            if '月份' in df_perf.columns:
                df_perf['月份'] = df_perf['月份'].astype(str).str.extract(r'(\d+)')[0].astype(float).fillna(0).astype(int)
            if '发货箱数' in df_perf.columns:
                df_perf['发货箱数'] = pd.to_numeric(df_perf['发货箱数'], errors='coerce').fillna(0)
            if '发货金额' in df_perf.columns:
                df_perf['发货金额'] = pd.to_numeric(df_perf['发货金额'], errors='coerce').fillna(0)

            if '年份' in df_perf.columns and '月份' in df_perf.columns:
                df_perf = df_perf[(df_perf['年份'] > 0) & (df_perf['月份'].between(1, 12))]
                df_perf['年月'] = pd.to_datetime(
                    df_perf['年份'].astype(str) + '-' + df_perf['月份'].astype(str).str.zfill(2) + '-01',
                    errors='coerce'
                )
            else:
                df_perf['年月'] = pd.NaT
            df_perf_raw = df_perf

        # --- Process Sheet 5 (Target) ---
        if df_target_raw is not None:
            df_target_raw.columns = [str(c).strip() for c in df_target_raw.columns]
            rename_target = {}
            if len(df_target_raw.columns) > 3: rename_target[df_target_raw.columns[3]] = '品类'
            if len(df_target_raw.columns) > 4: rename_target[df_target_raw.columns[4]] = '月份'
            if len(df_target_raw.columns) > 5: rename_target[df_target_raw.columns[5]] = '任务量'
            df_target_raw = df_target_raw.rename(columns=rename_target)
            if '月份' in df_target_raw.columns:
                df_target_raw['月份'] = df_target_raw['月份'].astype(str).str.extract(r'(\d+)')[0].astype(float).fillna(0).astype(int)
            if '任务量' in df_target_raw.columns:
                df_target_raw['任务量'] = pd.to_numeric(df_target_raw['任务量'], errors='coerce').fillna(0)

        # --- Process Sheet 6 (Scan Data) ---
        if df_scan_raw is not None:
            df0 = df_scan_raw

            def _col(idx: int):
                if idx < df0.shape[1]:
                    return df0.iloc[:, idx]
                return pd.Series([None] * len(df0))

            df_scan_raw = pd.DataFrame({
                "门店名称": _col(1),
                "经销商名称": _col(18),
                "省区": _col(17),
                "产品大类": _col(19),
                "产品小类": _col(20),
                "经纬度": _col(12),
                "年份": _col(13),
                "月份": _col(14),
                "日": _col(15),
            })

            df_scan_raw["年份"] = df_scan_raw["年份"].astype(str).str.extract(r"(\d+)")[0].astype(float).fillna(0).astype(int)
            df_scan_raw["年份"] = df_scan_raw["年份"].apply(lambda y: y + 2000 if 0 < y < 100 else y)
            df_scan_raw["月份"] = df_scan_raw["月份"].astype(str).str.extract(r"(\d+)")[0].astype(float).fillna(0).astype(int)
            df_scan_raw["日"] = df_scan_raw["日"].astype(str).str.extract(r"(\d+)")[0].astype(float).fillna(0).astype(int)

            for c in ["门店名称", "省区", "经销商名称", "产品大类", "产品小类"]:
                df_scan_raw[c] = df_scan_raw[c].fillna("").astype(str).str.strip()

            coords = df_scan_raw["经纬度"].apply(_parse_lon_lat)
            df_scan_raw["经度"] = coords.apply(lambda x: x[0])
            df_scan_raw["纬度"] = coords.apply(lambda x: x[1])

        return df, month_cols, df_stock, df_q4_raw, df_perf_raw, df_target_raw, df_scan_raw, debug_logs
        
    except Exception as e:
        import traceback
        return None, None, None, None, None, None, None, [f"Error: {str(e)}", traceback.format_exc()]

@st.cache_data(ttl=3600)
def load_builtin_perf_2025():
    base_dir = os.path.dirname(__file__) if "__file__" in globals() else os.getcwd()
    path = os.path.join(base_dir, "分析底表0115.xlsx")
    if not os.path.exists(path):
        return pd.DataFrame()
    xl = pd.ExcelFile(path)
    sheet_name = next((s for s in xl.sheet_names if "发货" in str(s)), None)
    if sheet_name is None and len(xl.sheet_names) > 3:
        sheet_name = xl.sheet_names[3]
    if sheet_name is None:
        return pd.DataFrame()
    df0 = xl.parse(sheet_name)
    df0.columns = [str(c).strip() for c in df0.columns]
    col_year = next((c for c in df0.columns if str(c).strip() == "年份" or "年" in str(c)), None)
    col_month = next((c for c in df0.columns if str(c).strip() == "月份" or "月" in str(c)), None)
    col_prov = next((c for c in df0.columns if "省区" in str(c)), None)
    col_dist = next((c for c in df0.columns if "客户简称" in str(c)), None) or next((c for c in df0.columns if "购货单位" in str(c)), None)
    col_qty = next((c for c in df0.columns if "基本数量" in str(c)), None) or next((c for c in df0.columns if "箱" in str(c) or "数量" in str(c)), None)
    col_amt = next((c for c in df0.columns if "原价金额" in str(c)), None) or next((c for c in df0.columns if "金额" in str(c)), None)
    col_wh = next((c for c in df0.columns if "发货仓" in str(c)), None)
    col_grp = next((c for c in df0.columns if "归类" in str(c)), None)
    col_bigcat = next((c for c in df0.columns if str(c).strip() == "大分类"), None) or next((c for c in df0.columns if "月分析" in str(c)), None)
    col_big = next((c for c in df0.columns if str(c).strip() == "大类"), None)
    col_mid = next((c for c in df0.columns if str(c).strip() == "中类"), None)
    col_small = next((c for c in df0.columns if str(c).strip() == "小类"), None)

    df = pd.DataFrame()
    if col_year is not None: df["年份"] = df0[col_year]
    if col_month is not None: df["月份"] = df0[col_month]
    if col_prov is not None: df["省区"] = df0[col_prov]
    if col_dist is not None: df["经销商名称"] = df0[col_dist]
    if col_qty is not None: df["发货箱数"] = df0[col_qty]
    if col_amt is not None: df["发货金额"] = df0[col_amt]
    if col_wh is not None: df["发货仓"] = df0[col_wh]
    if col_mid is not None: df["中类"] = df0[col_mid]
    if col_grp is not None: df["归类"] = df0[col_grp]
    if col_bigcat is not None: df["大分类"] = df0[col_bigcat]
    if col_big is not None: df["大类"] = df0[col_big]
    if col_small is not None: df["小类"] = df0[col_small]

    for c in ["省区", "经销商名称", "发货仓", "中类", "归类", "大分类", "大类", "小类"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()
    if "年份" in df.columns:
        df["年份"] = df["年份"].astype(str).str.extract(r"(\d+)")[0].astype(float).fillna(0).astype(int)
        df["年份"] = df["年份"].apply(lambda y: y + 2000 if 0 < y < 100 else y)
    if "月份" in df.columns:
        df["月份"] = df["月份"].astype(str).str.extract(r"(\d+)")[0].astype(float).fillna(0).astype(int)
    if "发货箱数" in df.columns:
        df["发货箱数"] = pd.to_numeric(df["发货箱数"], errors="coerce").fillna(0)
    if "发货金额" in df.columns:
        df["发货金额"] = pd.to_numeric(df["发货金额"], errors="coerce").fillna(0)
    if "年份" in df.columns and "月份" in df.columns:
        df = df[(df["年份"] == 2025) & (df["月份"].between(1, 12))]
        df["年月"] = pd.to_datetime(df["年份"].astype(str) + "-" + df["月份"].astype(str).str.zfill(2) + "-01", errors="coerce")
    else:
        return pd.DataFrame()
    return df

@st.cache_data(ttl=3600)
def load_builtin_scan_2025():
    # Attempt to load built-in file if it exists, otherwise return empty
    base_dir = os.path.dirname(__file__) if "__file__" in globals() else os.getcwd()
    path = os.path.join(base_dir, "分析底表0115.xlsx")
    if not os.path.exists(path):
        return pd.DataFrame()
    
    try:
        xl = pd.ExcelFile(path)
        if len(xl.sheet_names) <= 5:
            return pd.DataFrame()

        df0 = xl.parse(5)
        if df0 is None or df0.empty:
            return pd.DataFrame()

        def _col(idx: int):
            if idx < df0.shape[1]:
                return df0.iloc[:, idx]
            return pd.Series([None] * len(df0))

        df = pd.DataFrame({
            "门店名称": _col(1),
            "经销商名称": _col(18),
            "省区": _col(17),
            "产品大类": _col(19),
            "产品小类": _col(20),
            "经纬度": _col(12),
            "年份": _col(13),
            "月份": _col(14),
            "日": _col(15),
        })

        df["年份"] = df["年份"].astype(str).str.extract(r"(\d+)")[0].astype(float).fillna(0).astype(int)
        df["年份"] = df["年份"].apply(lambda y: y + 2000 if 0 < y < 100 else y)
        df["月份"] = df["月份"].astype(str).str.extract(r"(\d+)")[0].astype(float).fillna(0).astype(int)
        df["日"] = df["日"].astype(str).str.extract(r"(\d+)")[0].astype(float).fillna(0).astype(int)

        for c in ["门店名称", "省区", "经销商名称", "产品大类", "产品小类"]:
            df[c] = df[c].fillna("").astype(str).str.strip()

        coords = df["经纬度"].apply(_parse_lon_lat)
        df["经度"] = coords.apply(lambda x: x[0])
        df["纬度"] = coords.apply(lambda x: x[1])

        df = df[df["年份"] == 2025]
        return df
    except Exception:
        return pd.DataFrame()

# -----------------------------------------------------------------------------
# 4. Layout
# -----------------------------------------------------------------------------

st.markdown("## 🛠️ 数据控制台")

if 'hc_mode' not in st.session_state:
    st.session_state.hc_mode = False

st.toggle("高对比模式", key="hc_mode")

if st.session_state.get("hc_mode"):
    st.markdown("""
    <style>
      :root {
        --tbl-header-bg: #0B57D0;
        --tbl-header-bg-hover: #0846AB;
        --tbl-header-border: #06357F;
        --tbl-header-fg: #FFFFFF;
        --tbl-header-icon: #FFFFFF;
        --tbl-header-shadow: 0 10px 22px rgba(0, 0, 0, 0.38);
      }
    </style>
    """, unsafe_allow_html=True)

if 'exp_upload' not in st.session_state:
    st.session_state.exp_upload = True
if 'exp_filter' not in st.session_state:
    st.session_state.exp_filter = True

with st.expander("📥 数据导入", expanded=st.session_state.exp_upload):
    uploaded_file = st.file_uploader("导入数据表 (Excel/CSV)", type=['xlsx', 'xls', 'csv'], key="main_uploader")

if uploaded_file is None:
    st.markdown(
        """
        <div style='text-align: center; padding: 60px 20px; background-color: white; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); margin-bottom: 40px;'>
            <h1 style='color: #4096ff; margin-bottom: 16px;'>👋 欢迎使用美思雅数据分析系统</h1>
            <p style='color: #666; font-size: 16px; margin-bottom: 0;'>请上传 Excel 数据文件以解锁完整分析面板</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    # st.stop()  # Streamlit Cloud Health Check Fix: Avoid blocking the app here

# Main Logic
if uploaded_file:
    uploaded_name = uploaded_file.name
    cached_bytes = st.session_state.get("_uploaded_bytes")
    cached_sig = st.session_state.get("_uploaded_sig")
    cached_name = st.session_state.get("_uploaded_name")

    if cached_bytes is None or cached_sig is None or cached_name != uploaded_name:
        cached_bytes = uploaded_file.getvalue()
        cached_sig = hashlib.md5(cached_bytes).hexdigest()
        st.session_state["_uploaded_bytes"] = cached_bytes
        st.session_state["_uploaded_sig"] = cached_sig
        st.session_state["_uploaded_name"] = uploaded_name

    if st.session_state.get("_active_file_sig") != cached_sig:
        st.session_state["_active_file_sig"] = cached_sig
        st.session_state["run_analysis"] = False

    parsed_cache = st.session_state.get("_parsed_cache", {})
    if cached_sig in parsed_cache:
        # Check if cache is old version (7 elements) or new version (8 elements)
        cache_val = parsed_cache[cached_sig]
        if len(cache_val) == 8:
            df_raw, month_cols, df_stock_raw, df_q4_raw, df_perf_raw, df_target_raw, df_scan_raw, debug_logs = cache_val
        else:
            # Re-parse if cache format mismatch (Old cache had 7 items)
            df_raw, month_cols, df_stock_raw, df_q4_raw, df_perf_raw, df_target_raw, df_scan_raw, debug_logs = load_data_v3(cached_bytes, uploaded_name)
            parsed_cache[cached_sig] = (df_raw, month_cols, df_stock_raw, df_q4_raw, df_perf_raw, df_target_raw, df_scan_raw, debug_logs)
            st.session_state["_parsed_cache"] = parsed_cache
    else:
        df_raw, month_cols, df_stock_raw, df_q4_raw, df_perf_raw, df_target_raw, df_scan_raw, debug_logs = load_data_v3(cached_bytes, uploaded_name)
        parsed_cache[cached_sig] = (df_raw, month_cols, df_stock_raw, df_q4_raw, df_perf_raw, df_target_raw, df_scan_raw, debug_logs)
        if len(parsed_cache) > 2:
            for k in list(parsed_cache.keys())[:-2]:
                parsed_cache.pop(k, None)
        st.session_state["_parsed_cache"] = parsed_cache

    df_perf_2025 = load_builtin_perf_2025()
    if df_perf_2025 is not None and not df_perf_2025.empty:
        if df_perf_raw is None or getattr(df_perf_raw, "empty", True):
            df_perf_raw = df_perf_2025.copy()
        else:
            years = pd.to_numeric(df_perf_raw.get("年份", pd.Series(dtype=float)), errors="coerce")
            if not bool((years == 2025).any()):
                df_perf_raw = pd.concat([df_perf_2025, df_perf_raw], ignore_index=True, sort=False)
                
    df_scan_2025 = load_builtin_scan_2025()
    if df_scan_2025 is not None and not df_scan_2025.empty:
        if df_scan_raw is None or getattr(df_scan_raw, "empty", True):
            df_scan_raw = df_scan_2025.copy()
        else:
            years_s = pd.to_numeric(df_scan_raw.get("年份", pd.Series(dtype=float)), errors="coerce")
            if not bool((years_s == 2025).any()):
                df_scan_raw = pd.concat([df_scan_2025, df_scan_raw], ignore_index=True, sort=False)

    if df_raw is None and debug_logs:
        st.error("数据加载失败。详细日志如下：")
        st.text("\n".join(debug_logs))

    if df_raw is not None:
        # --- Filters Area ---
        with st.expander("🔎 筛选搜索", expanded=st.session_state.exp_filter):
            # Province Filter
            provinces = ['全部'] + sorted(list(df_raw['省区'].unique()))
            sel_prov = st.selectbox("选择省区 (Province)", provinces)
            
            # Distributor Filter
            if sel_prov != '全部':
                dist_options = ['全部'] + sorted(list(df_raw[df_raw['省区']==sel_prov]['经销商名称'].unique()))
            else:
                dist_options = ['全部'] + sorted(list(df_raw['经销商名称'].unique()))
            sel_dist = st.selectbox("选择经销商 (Distributor)", dist_options)

            cat_set = set()
            for _df, _col in [
                (df_perf_raw, '大分类'),
                (df_perf_raw, '产品大类'),
                (df_q4_raw, '产品大类'),
                (df_stock_raw, '产品大类'),
                (df_scan_raw, '产品大类'),
            ]:
                if _df is not None and not getattr(_df, "empty", True) and _col in _df.columns:
                    cat_set |= set(_df[_col].fillna('').astype(str).str.strip().tolist())
            cat_options = ['全部'] + sorted([x for x in cat_set if x])
            sel_cat = st.selectbox("选择产品大类 (Category)", cat_options, key="main_sel_cat")
        
        # Apply Filters
        df = df_raw.copy()
        if sel_prov != '全部':
            df = df[df['省区'] == sel_prov]
        if sel_dist != '全部':
            df = df[df['经销商名称'] == sel_dist]
            
        if not st.session_state.get('run_analysis', False):
            st.markdown("### ✅ 数据已加载")
            st.caption("点击「开始分析 🚀」进入分析页面。")
            if st.button("开始分析 🚀", type="primary", key="main_start_analysis"):
                st.session_state['run_analysis'] = True

        # Share / external-access UI intentionally removed
            
        if st.session_state.get('run_analysis', False):
            
            # --- Header ---
            st.title("📈 美思雅数据分析系统")
            st.markdown(f"当前数据范围: **{sel_prov}** / **{sel_dist}** | 包含 **{len(df)}** 家门店")
            
            # --- Tabs ---
            tab1, tab7, tab6, tab_out, tab_scan, tab3, tab_other = st.tabs(["📊 核心概览", "🚀 业绩分析", "📦 库存分析", "🚚 出库分析", "📱 扫码分析", "📈 ABCD效能分析", "其他分析"])
            
            # === TAB 1: OVERVIEW ===
            with tab1:
                st.caption(f"筛选口径：省区={sel_prov}｜经销商={sel_dist}｜产品大类={st.session_state.get('main_sel_cat', '全部')}")

                # --- Common Helpers for Tab 1 ---
                def _fmt_wan(x): return fmt_num((x or 0) / 10000)
                def _fmt_pct(x): return fmt_pct_ratio(x) if x is not None else "—"
                def _arrow(x): return "↑" if x and x>0 else ("↓" if x and x<0 else "")
                def _trend_cls(x): return "trend-up" if x and x > 0 else ("trend-down" if x and x < 0 else "trend-neutral")

                # Card Renderer for Performance (Tab 7 Style)
                def _render_perf_card(title, icon, val_wan, target_wan, rate, yoy_val_wan, yoy_pct):
                    trend_cls = _trend_cls(yoy_pct)
                    arrow = _arrow(yoy_pct)
                    rate_txt = _fmt_pct(rate)
                    yoy_txt = _fmt_pct(yoy_pct)
                    pct_val = min(max(rate * 100 if rate else 0, 0), 100)
                    prog_color = "#28A745" if rate and rate >= 1.0 else ("#FFC107" if rate and rate >= 0.8 else "#DC3545")

                    st.markdown(f"""
                    <div class="out-kpi-card">
                        <div class="out-kpi-bar"></div>
                        <div class="out-kpi-head">
                            <div class="out-kpi-ico">{icon}</div>
                            <div class="out-kpi-title">{title}</div>
                        </div>
                        <div class="out-kpi-val">¥ {val_wan}万</div>
                        <div class="out-kpi-sub2" style="margin-top:8px;">
                            <span>达成率</span>
                            <span style="font-weight:800; color:{prog_color}">{rate_txt}</span>
                        </div>
                        <div class="out-kpi-progress" style="margin-top:6px;">
                            <div class="out-kpi-progress-bar" style="background:{prog_color}; width:{pct_val}%;"></div>
                        </div>
                        <div class="out-kpi-sub2" style="margin-top:10px;">
                            <span>目标</span>
                            <span>{target_wan}万</span>
                        </div>
                        <div class="out-kpi-sub2">
                            <span>同期</span>
                            <span>{yoy_val_wan}万</span>
                        </div>
                        <div class="out-kpi-sub2">
                            <span>同比</span>
                            <span class="{trend_cls}">{arrow} {yoy_txt}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                # Card Renderer for Outbound/Scan (General Style)
                def _render_general_card(title, icon, main_val, sub_items):
                    # sub_items: list of (label, value_html)
                    rows_html = ""
                    for label, val_html in sub_items:
                        rows_html += f'<div class="out-kpi-sub2"><span>{label}</span><span>{val_html}</span></div>'
                    
                    st.markdown(f"""
                    <div class="out-kpi-card">
                        <div class="out-kpi-bar"></div>
                        <div class="out-kpi-head">
                            <div class="out-kpi-ico">{icon}</div>
                            <div class="out-kpi-title">{title}</div>
                        </div>
                        <div class="out-kpi-val">{main_val}</div>
                        <div style="margin-top:10px;">{rows_html}</div>
                    </div>
                    """, unsafe_allow_html=True)

                sel_bigcat = st.session_state.get("main_sel_cat", "全部")

                def _filter_common(_df):
                    if _df is None or getattr(_df, "empty", True):
                        return pd.DataFrame()
                    d = _df.copy()
                    for c in ['省区', '经销商名称', '产品大类', '大分类']:
                        if c in d.columns:
                            d[c] = d[c].fillna('').astype(str).str.strip()
                    if sel_prov != '全部' and '省区' in d.columns:
                        d = d[d['省区'] == sel_prov]
                    if sel_dist != '全部' and '经销商名称' in d.columns:
                        d = d[d['经销商名称'] == sel_dist]
                    if sel_bigcat != '全部':
                        if '产品大类' in d.columns:
                            d = d[d['产品大类'] == sel_bigcat]
                        elif '大分类' in d.columns:
                            d = d[d['大分类'] == sel_bigcat]
                    return d

                # ---------------------------------------------------------
                # 1. 核心业绩指标 (From Tab 7)
                # ---------------------------------------------------------
                st.markdown("### 🚀 核心业绩指标")
                df_perf = _filter_common(df_perf_raw)
                if not df_perf.empty:
                    # Data Prep
                    if '年份' in df_perf.columns:
                        df_perf['年份'] = pd.to_numeric(df_perf['年份'], errors='coerce').fillna(0).astype(int)
                    if '月份' in df_perf.columns:
                        df_perf['月份'] = pd.to_numeric(df_perf['月份'], errors='coerce').fillna(0).astype(int)
                    amt_col = '发货金额' if '发货金额' in df_perf.columns else None
                    if amt_col:
                        df_perf[amt_col] = pd.to_numeric(df_perf[amt_col], errors='coerce').fillna(0)
                    
                    years_avail = sorted([y for y in df_perf['年份'].unique().tolist() if y > 2000])
                    perf_y = max(years_avail) if years_avail else 2025
                    months_avail = sorted([m for m in df_perf[df_perf['年份'] == perf_y]['月份'].unique().tolist() if 1 <= m <= 12])
                    perf_m = max(months_avail) if months_avail else 1
                    last_y = perf_y - 1

                    # Actuals
                    cur_m_amt = df_perf[(df_perf['年份'] == perf_y) & (df_perf['月份'] == perf_m)][amt_col].sum() if amt_col else 0
                    last_m_amt = df_perf[(df_perf['年份'] == last_y) & (df_perf['月份'] == perf_m)][amt_col].sum() if amt_col else 0
                    cur_y_amt = df_perf[df_perf['年份'] == perf_y][amt_col].sum() if amt_col else 0
                    last_y_amt = df_perf[df_perf['年份'] == last_y][amt_col].sum() if amt_col else 0

                    yoy_m = (cur_m_amt - last_m_amt) / last_m_amt if last_m_amt > 0 else 0
                    yoy_y = (cur_y_amt - last_y_amt) / last_y_amt if last_y_amt > 0 else 0

                    # Targets
                    t_cur_m = 0.0
                    t_cur_y = 0.0
                    if df_target_raw is not None and not getattr(df_target_raw, "empty", True):
                        df_t = df_target_raw.copy()
                        for c in ['省区', '品类']:
                            if c in df_t.columns: df_t[c] = df_t[c].fillna('').astype(str).str.strip()
                        if '月份' in df_t.columns: df_t['月份'] = pd.to_numeric(df_t['月份'], errors='coerce').fillna(0).astype(int)
                        if '任务量' in df_t.columns: df_t['任务量'] = pd.to_numeric(df_t['任务量'], errors='coerce').fillna(0)
                        
                        if sel_prov != '全部' and '省区' in df_t.columns:
                            df_t = df_t[df_t['省区'] == sel_prov]
                        # Target usually doesn't filter by Distributor, but filters by Category
                        if sel_bigcat != '全部' and '品类' in df_t.columns:
                            df_t = df_t[df_t['品类'] == sel_bigcat]
                        
                        t_cur_m = df_t[df_t['月份'] == perf_m]['任务量'].sum()
                        t_cur_y = df_t['任务量'].sum() # Total Year Target

                    rate_m = (cur_m_amt / t_cur_m) if t_cur_m > 0 else None
                    rate_y = (cur_y_amt / t_cur_y) if t_cur_y > 0 else None

                    c1, c2 = st.columns(2)
                    with c1:
                        _render_perf_card(f"本月业绩（{perf_m}月）", "📅", _fmt_wan(cur_m_amt), _fmt_wan(t_cur_m), rate_m, _fmt_wan(last_m_amt), yoy_m)
                    with c2:
                        _render_perf_card(f"年度累计业绩（{perf_y}年）", "🏆", _fmt_wan(cur_y_amt), _fmt_wan(t_cur_y), rate_y, _fmt_wan(last_y_amt), yoy_y)
                else:
                    st.info("业绩数据为空或不含匹配字段")

                st.markdown("---")

                # ---------------------------------------------------------
                # 2. 库存关键指标 (From Tab 6)
                # ---------------------------------------------------------
                st.markdown("### 📦 库存关键指标")
                df_stock = _filter_common(df_stock_raw)
                if not df_stock.empty:
                    # Prepare Data for Metrics
                    stock_box_col = '箱数' if '箱数' in df_stock.columns else next((c for c in df_stock.columns if '箱' in str(c)), None)
                    stock_boxes = float(pd.to_numeric(df_stock[stock_box_col], errors='coerce').fillna(0).sum()) if stock_box_col else 0.0
                    
                    # Q4 Avg Sales (Need logic from Tab 6)
                    total_q4_avg = 0.0
                    if df_q4_raw is not None and not getattr(df_q4_raw, "empty", True):
                        # Simple estimation: Filter Q4 raw by current filters -> Sum Q4 months -> Divide by 3
                        # Tab 6 logic is more complex (Distributor based), but for Overview Total, simple sum is close enough.
                        # However, let's try to match Tab 6 logic: Sum 'Q4_Avg' of relevant distributors.
                        
                        # 1. Get filtered distributors
                        valid_dists = df_stock['经销商名称'].unique()
                        
                        # 2. Calculate Q4 Sales for these distributors
                        df_q4_f = df_q4_raw.copy()
                        if '年份' in df_q4_f.columns: df_q4_f = df_q4_f[df_q4_f['年份'] == 2025] # Q4 assumption
                        if '经销商名称' in df_q4_f.columns:
                            df_q4_f = df_q4_f[df_q4_f['经销商名称'].isin(valid_dists)]
                        
                        # Filter for Oct, Nov, Dec
                        if '月份' in df_q4_f.columns:
                            df_q4_f['月份'] = pd.to_numeric(df_q4_f['月份'], errors='coerce').fillna(0).astype(int)
                            df_q4_f = df_q4_f[df_q4_f['月份'].isin([10, 11, 12])]
                        
                        qty_col = '数量(箱)' if '数量(箱)' in df_q4_f.columns else next((c for c in df_q4_f.columns if '数量' in str(c)), None)
                        if qty_col:
                            total_q4_sales = pd.to_numeric(df_q4_f[qty_col], errors='coerce').sum()
                            total_q4_avg = total_q4_sales / 3.0

                    dos = stock_boxes / total_q4_avg if total_q4_avg > 0 else 0.0
                    
                    # Abnormal Count (Simplify for Overview)
                    # Tab 6 calculates per distributor. Here we just show global metrics.
                    
                    m1, m2, m3 = st.columns(3)
                    m1.metric("📦 总库存 (箱)", fmt_num(stock_boxes))
                    m2.metric("📉 Q4月均销", fmt_num(total_q4_avg))
                    m3.metric("📅 整体可销月 (DOS)", fmt_num(dos))
                else:
                    st.info("库存数据为空")

                st.markdown("---")

                # ---------------------------------------------------------
                # 3. 出库关键指标 (From Tab Out)
                # ---------------------------------------------------------
                st.markdown("### 🚚 出库关键指标")
                df_out = _filter_common(df_q4_raw)
                if not df_out.empty:
                    # Date Prep
                    tmp = df_out.copy()
                    for c in ['年份', '月份']: 
                        if c in tmp.columns: tmp[c] = pd.to_numeric(tmp[c], errors='coerce').fillna(0).astype(int)
                    if '日' in tmp.columns: tmp['日'] = pd.to_numeric(tmp['日'], errors='coerce').fillna(0).astype(int)
                    qty_col = '数量(箱)' if '数量(箱)' in tmp.columns else next((c for c in tmp.columns if '数量' in str(c) or '箱' in str(c)), None)
                    if qty_col:
                        tmp['数量(箱)'] = pd.to_numeric(tmp[qty_col], errors='coerce').fillna(0)
                        tmp = tmp[tmp['年份'] > 0]
                        
                        oy = int(tmp['年份'].max())
                        om = int(tmp[tmp['年份'] == oy]['月份'].max())
                        od = int(tmp[(tmp['年份'] == oy) & (tmp['月份'] == om)]['日'].max())
                        
                        # Current
                        today_boxes = tmp[(tmp['年份'] == oy) & (tmp['月份'] == om) & (tmp['日'] == od)]['数量(箱)'].sum()
                        month_boxes = tmp[(tmp['年份'] == oy) & (tmp['月份'] == om)]['数量(箱)'].sum()
                        year_boxes = tmp[tmp['年份'] == oy]['数量(箱)'].sum()
                        
                        # Last Year
                        ly = oy - 1
                        l_today_boxes = tmp[(tmp['年份'] == ly) & (tmp['月份'] == om) & (tmp['日'] == od)]['数量(箱)'].sum()
                        l_month_boxes = tmp[(tmp['年份'] == ly) & (tmp['月份'] == om)]['数量(箱)'].sum()
                        l_year_boxes = tmp[tmp['年份'] == ly]['数量(箱)'].sum()
                        
                        # YoY
                        yoy_d = (today_boxes - l_today_boxes) / l_today_boxes if l_today_boxes > 0 else 0
                        yoy_m = (month_boxes - l_month_boxes) / l_month_boxes if l_month_boxes > 0 else 0
                        yoy_y = (year_boxes - l_year_boxes) / l_year_boxes if l_year_boxes > 0 else 0
                        
                        k1, k2, k3 = st.columns(3)
                        with k1:
                            trend = _trend_cls(yoy_d)
                            arr = _arrow(yoy_d)
                            _render_general_card("本日出库", "🚚", f"{fmt_num(today_boxes)} 箱", [
                                ("同期", f"{fmt_num(l_today_boxes)} 箱"),
                                ("同比", f'<span class="{trend}">{arr} {_fmt_pct(yoy_d)}</span>')
                            ])
                        with k2:
                            trend = _trend_cls(yoy_m)
                            arr = _arrow(yoy_m)
                            _render_general_card(f"本月累计出库（{om}月）", "📦", f"{fmt_num(month_boxes)} 箱", [
                                ("同期", f"{fmt_num(l_month_boxes)} 箱"),
                                ("同比", f'<span class="{trend}">{arr} {_fmt_pct(yoy_m)}</span>')
                            ])
                        with k3:
                            trend = _trend_cls(yoy_y)
                            arr = _arrow(yoy_y)
                            _render_general_card(f"本年累计出库（{oy}年）", "🧾", f"{fmt_num(year_boxes)} 箱", [
                                ("同期", f"{fmt_num(l_year_boxes)} 箱"),
                                ("同比", f'<span class="{trend}">{arr} {_fmt_pct(yoy_y)}</span>')
                            ])
                else:
                    st.info("出库数据为空")

                st.markdown("---")

                # ---------------------------------------------------------
                # 4. 扫码率概览 (From Tab Scan)
                # ---------------------------------------------------------
                st.markdown("### 📱 扫码率概览")
                df_scan = _filter_common(df_scan_raw)
                # Re-use out_base from above or re-calc
                if not df_scan.empty and not df_out.empty:
                    # Ensure Date Cols
                    for c in ['年份', '月份', '日']:
                        if c in df_scan.columns: df_scan[c] = pd.to_numeric(df_scan[c], errors='coerce').fillna(0).astype(int)
                    
                    # Use same oy, om, od from Outbound
                    scan_today = len(df_scan[(df_scan['年份'] == oy) & (df_scan['月份'] == om) & (df_scan['日'] == od)]) / 6.0
                    scan_month = len(df_scan[(df_scan['年份'] == oy) & (df_scan['月份'] == om)]) / 6.0
                    scan_year = len(df_scan[df_scan['年份'] == oy]) / 6.0
                    
                    l_scan_today = len(df_scan[(df_scan['年份'] == ly) & (df_scan['月份'] == om) & (df_scan['日'] == od)]) / 6.0
                    l_scan_month = len(df_scan[(df_scan['年份'] == ly) & (df_scan['月份'] == om)]) / 6.0
                    l_scan_year = len(df_scan[df_scan['年份'] == ly]) / 6.0

                    rate_today = scan_today / today_boxes if today_boxes > 0 else 0
                    rate_month = scan_month / month_boxes if month_boxes > 0 else 0
                    rate_year = scan_year / year_boxes if year_boxes > 0 else 0
                    
                    yoy_rate_d = rate_today - (l_scan_today / l_today_boxes if l_today_boxes > 0 else 0)
                    yoy_rate_m = rate_month - (l_scan_month / l_month_boxes if l_month_boxes > 0 else 0)
                    yoy_rate_y = rate_year - (l_scan_year / l_year_boxes if l_year_boxes > 0 else 0)

                    s1, s2, s3 = st.columns(3)
                    with s1:
                        trend = _trend_cls(yoy_rate_d)
                        arr = _arrow(yoy_rate_d)
                        _render_general_card("本日扫码率", "📱", fmt_pct_ratio(rate_today), [
                            ("扫码 / 出库", f"{fmt_num(scan_today)} / {fmt_num(today_boxes)}"),
                            ("同比增减", f'<span class="{trend}">{arr} {fmt_pct_value(yoy_rate_d*100)}</span>')
                        ])
                    with s2:
                        trend = _trend_cls(yoy_rate_m)
                        arr = _arrow(yoy_rate_m)
                        _render_general_card("本月扫码率", "🗓️", fmt_pct_ratio(rate_month), [
                            ("扫码 / 出库", f"{fmt_num(scan_month)} / {fmt_num(month_boxes)}"),
                            ("同比增减", f'<span class="{trend}">{arr} {fmt_pct_value(yoy_rate_m*100)}</span>')
                        ])
                    with s3:
                        trend = _trend_cls(yoy_rate_y)
                        arr = _arrow(yoy_rate_y)
                        _render_general_card("本年扫码率", "📈", fmt_pct_ratio(rate_year), [
                            ("扫码 / 出库", f"{fmt_num(scan_year)} / {fmt_num(year_boxes)}"),
                            ("同比增减", f'<span class="{trend}">{arr} {fmt_pct_value(yoy_rate_y*100)}</span>')
                        ])
                else:
                    st.info("扫码数据为空")

            # === TAB SCAN: SCAN ANALYSIS ===
            with tab_scan:
                if df_scan_raw is not None and not df_scan_raw.empty:
                    st.subheader("📱 扫码分析")
                    
                    # 1. Date Calculation
                    # Today: Max date in max month of 2026
                    max_scan_date = None
                    df_scan_2026 = df_scan_raw[df_scan_raw['年份'] == 2026]
                    if not df_scan_2026.empty:
                        max_month = df_scan_2026['月份'].max()
                        max_day = df_scan_2026[df_scan_2026['月份'] == max_month]['日'].max()
                        max_scan_date = pd.Timestamp(year=2026, month=max_month, day=max_day)
                    
                    if max_scan_date:
                        cur_year = max_scan_date.year
                        cur_month = max_scan_date.month
                        cur_day = max_scan_date.day
                        st.info(f"📅 当前统计日期：{cur_year}年{cur_month}月{cur_day}日")
                    else:
                        st.warning("⚠️ 未找到2026年扫码数据，无法计算当日/当月指标")
                        cur_year, cur_month, cur_day = 2026, 1, 1

                    # 2. Filter Area
                    with st.expander("🔎 扫码筛选", expanded=True):
                        c_s1, c_s2, c_s3 = st.columns(3)
                        # Province
                        prov_opts_s = ['全部'] + sorted(df_scan_raw['省区'].unique().tolist())
                        sel_prov_s = c_s1.selectbox("省区", prov_opts_s, key="scan_prov")
                        
                        # Distributor
                        if sel_prov_s != '全部':
                            dist_opts_s = ['全部'] + sorted(df_scan_raw[df_scan_raw['省区'] == sel_prov_s]['经销商名称'].unique().tolist())
                        else:
                            dist_opts_s = ['全部'] + sorted(df_scan_raw['经销商名称'].unique().tolist())
                        sel_dist_s = c_s2.selectbox("经销商", dist_opts_s, key="scan_dist")
                        
                        # Category
                        cat_opts_s = ['全部'] + sorted(df_scan_raw['产品大类'].unique().tolist())
                        sel_cat_s = c_s3.selectbox("产品大类", cat_opts_s, key="scan_cat")

                    # Apply Filters
                    df_s_flt = df_scan_raw.copy()
                    last_year = cur_year - 1
                    out_base_df = None
                    out_day_df = None
                    out_day_last_df = None
                    if df_q4_raw is not None and not getattr(df_q4_raw, "empty", True):
                        tmp = df_q4_raw.copy()
                        for c in ['年份', '月份']:
                            if c in tmp.columns:
                                tmp[c] = pd.to_numeric(tmp[c], errors='coerce').fillna(0).astype(int)
                        day_col_out = None
                        if '日' in tmp.columns:
                            day_col_out = '日'
                            tmp['日'] = pd.to_numeric(tmp['日'], errors='coerce').fillna(0).astype(int)
                        else:
                            cand = next((c for c in tmp.columns if '日期' in str(c)), None)
                            if cand:
                                dt = pd.to_datetime(tmp[cand], errors='coerce')
                                tmp['年份'] = dt.dt.year
                                tmp['月份'] = dt.dt.month
                                tmp['日'] = dt.dt.day
                                day_col_out = '日'
                        qty_col_out = '数量(箱)' if '数量(箱)' in tmp.columns else next((c for c in tmp.columns if '数量' in str(c) or '箱' in str(c)), None)
                        if qty_col_out:
                            tmp['数量(箱)'] = pd.to_numeric(tmp[qty_col_out], errors='coerce').fillna(0)
                            if all(k in tmp.columns for k in ['年份', '月份', '日']):
                                out_base_df = tmp.copy()
                                for c in ['省区', '经销商名称', '产品大类', '大分类']:
                                    if c in out_base_df.columns:
                                        out_base_df[c] = out_base_df[c].fillna('').astype(str).str.strip()

                    if sel_prov_s != '全部':
                        df_s_flt = df_s_flt[df_s_flt['省区'] == sel_prov_s]
                        if out_base_df is not None and '省区' in out_base_df.columns:
                            out_base_df = out_base_df[out_base_df['省区'] == sel_prov_s]
                    if sel_dist_s != '全部':
                        df_s_flt = df_s_flt[df_s_flt['经销商名称'] == sel_dist_s]
                        if out_base_df is not None and '经销商名称' in out_base_df.columns:
                            out_base_df = out_base_df[out_base_df['经销商名称'] == sel_dist_s]
                    if sel_cat_s != '全部':
                        df_s_flt = df_s_flt[df_s_flt['产品大类'] == sel_cat_s]
                        if out_base_df is not None:
                            if '产品大类' in out_base_df.columns:
                                out_base_df = out_base_df[out_base_df['产品大类'] == sel_cat_s]
                            elif '大分类' in out_base_df.columns:
                                out_base_df = out_base_df[out_base_df['大分类'] == sel_cat_s]

                    if out_base_df is not None:
                        out_day_df = out_base_df[(out_base_df['年份'] == cur_year) & (out_base_df['月份'] == cur_month) & (out_base_df['日'] == cur_day)].copy()
                        out_day_last_df = out_base_df[(out_base_df['年份'] == last_year) & (out_base_df['月份'] == cur_month) & (out_base_df['日'] == cur_day)].copy()

                    # 3. Calculate Metrics (Scan vs Outbound)
                    # Unit: Box (6 tins = 1 box)
                    # Scan Count (Rows) / 6
                    
                    # --- Current Period (2026) ---
                    # Day
                    scan_day = len(df_s_flt[(df_s_flt['年份'] == cur_year) & (df_s_flt['月份'] == cur_month) & (df_s_flt['日'] == cur_day)]) / 6.0
                    out_day = 0
                    if out_day_df is not None:
                        qty_col_out = '数量(箱)' if '数量(箱)' in out_day_df.columns else next((c for c in out_day_df.columns if '数量' in str(c) or '箱' in str(c)), None)
                        if qty_col_out:
                            out_day = float(pd.to_numeric(out_day_df[qty_col_out], errors='coerce').fillna(0).sum())
                    out_day_last = 0
                    if out_day_last_df is not None:
                        qty_col_out = '数量(箱)' if '数量(箱)' in out_day_last_df.columns else next((c for c in out_day_last_df.columns if '数量' in str(c) or '箱' in str(c)), None)
                        if qty_col_out:
                            out_day_last = float(pd.to_numeric(out_day_last_df[qty_col_out], errors='coerce').fillna(0).sum())
                    
                    # Month
                    scan_month = len(df_s_flt[(df_s_flt['年份'] == cur_year) & (df_s_flt['月份'] == cur_month)]) / 6.0
                    out_month = float(pd.to_numeric(out_base_df[(out_base_df['年份'] == cur_year) & (out_base_df['月份'] == cur_month)]['数量(箱)'], errors='coerce').fillna(0).sum()) if out_base_df is not None else 0.0
                    
                    # Year
                    scan_year = len(df_s_flt[df_s_flt['年份'] == cur_year]) / 6.0
                    out_year = float(pd.to_numeric(out_base_df[out_base_df['年份'] == cur_year]['数量(箱)'], errors='coerce').fillna(0).sum()) if out_base_df is not None else 0.0

                    # --- Same Period Last Year (2025) ---
                    scan_day_last = len(df_s_flt[(df_s_flt['年份'] == last_year) & (df_s_flt['月份'] == cur_month) & (df_s_flt['日'] == cur_day)]) / 6.0
                    
                    # Month
                    scan_month_last = len(df_s_flt[(df_s_flt['年份'] == last_year) & (df_s_flt['月份'] == cur_month)]) / 6.0
                    out_month_last = float(pd.to_numeric(out_base_df[(out_base_df['年份'] == last_year) & (out_base_df['月份'] == cur_month)]['数量(箱)'], errors='coerce').fillna(0).sum()) if out_base_df is not None else 0.0
                    
                    # Year (YTD? or Full Year? Usually YTD for comparison or Full Year 2025)
                    # "同期" usually means same period. For Year, it means 2025 Full Year or YTD.
                    # Let's use Full Year 2025 for now as 2026 is incomplete.
                    scan_year_last = len(df_s_flt[df_s_flt['年份'] == last_year]) / 6.0
                    out_year_last = float(pd.to_numeric(out_base_df[out_base_df['年份'] == last_year]['数量(箱)'], errors='coerce').fillna(0).sum()) if out_base_df is not None else 0.0

                    # Rates
                    rate_month = (scan_month / out_month) if out_month > 0 else 0
                    rate_month_last = (scan_month_last / out_month_last) if out_month_last > 0 else 0
                    rate_year = (scan_year / out_year) if out_year > 0 else 0
                    rate_year_last = (scan_year_last / out_year_last) if out_year_last > 0 else 0
                    rate_day = (scan_day / out_day) if out_day > 0 else 0
                    rate_day_last = (scan_day_last / out_day_last) if out_day_last > 0 else 0

                    tab_overview, tab_s_cat, tab_s_prov, tab_s_map = st.tabs(["📊 扫码率概览", "🧩 分品类扫码率", "🗺️ 省区扫码率", "🧭 地图热力"])

                    with tab_overview:
                        st.caption(f"口径：今日 {cur_year}年{cur_month}月{cur_day}日｜本月 {cur_month}月｜本年 {cur_year}年")

                        def _trend_cls(x):
                            if x is None or (isinstance(x, float) and pd.isna(x)):
                                return "trend-neutral"
                            return "trend-up" if x > 0 else ("trend-down" if x < 0 else "trend-neutral")

                        def _arrow(x):
                            if x is None or (isinstance(x, float) and pd.isna(x)):
                                return ""
                            return "↑" if x > 0 else ("↓" if x < 0 else "")

                        yoy_rate_day = (rate_day - rate_day_last) if out_day_last > 0 else None
                        yoy_rate_month = (rate_month - rate_month_last) if out_month_last > 0 else None
                        yoy_rate_year = (rate_year - rate_year_last) if out_year_last > 0 else None
                        yoy_rate_day_pct = (yoy_rate_day * 100.0) if yoy_rate_day is not None else None
                        yoy_rate_month_pct = (yoy_rate_month * 100.0) if yoy_rate_month is not None else None
                        yoy_rate_year_pct = (yoy_rate_year * 100.0) if yoy_rate_year is not None else None

                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.markdown(f"""
                            <div class="out-kpi-card">
                                <div class="out-kpi-bar"></div>
                                <div class="out-kpi-head">
                                    <div class="out-kpi-ico">📱</div>
                                    <div class="out-kpi-title">本日扫码率</div>
                                </div>
                                <div class="out-kpi-val">{fmt_pct_ratio(rate_day)}</div>
                                <div class="out-kpi-sub"><span>出库(箱)</span><span>{fmt_num(out_day)}</span></div>
                                <div class="out-kpi-sub"><span>扫码(箱)</span><span>{fmt_num(scan_day)}</span></div>
                                <div class="out-kpi-sub2" style="margin-top:10px;"><span>同期({last_year})</span><span>{fmt_num(out_day_last)} 箱 / {fmt_num(scan_day_last)} 箱</span></div>
                                <div class="out-kpi-sub2"><span>同比（扫码率）</span><span class="{_trend_cls(yoy_rate_day)}">{_arrow(yoy_rate_day)} {fmt_pct_value(yoy_rate_day_pct) if yoy_rate_day_pct is not None else "—"}</span></div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div class="out-kpi-card">
                                <div class="out-kpi-bar"></div>
                                <div class="out-kpi-head">
                                    <div class="out-kpi-ico">🗓️</div>
                                    <div class="out-kpi-title">本月扫码率</div>
                                </div>
                                <div class="out-kpi-val">{fmt_pct_ratio(rate_month)}</div>
                                <div class="out-kpi-sub"><span>出库(箱)</span><span>{fmt_num(out_month)}</span></div>
                                <div class="out-kpi-sub"><span>扫码(箱)</span><span>{fmt_num(scan_month)}</span></div>
                                <div class="out-kpi-sub2" style="margin-top:10px;"><span>同期({last_year})</span><span>{fmt_num(out_month_last)} 箱 / {fmt_num(scan_month_last)} 箱</span></div>
                                <div class="out-kpi-sub2"><span>同比（扫码率）</span><span class="{_trend_cls(yoy_rate_month)}">{_arrow(yoy_rate_month)} {fmt_pct_value(yoy_rate_month_pct) if yoy_rate_month_pct is not None else "—"}</span></div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c3:
                            st.markdown(f"""
                            <div class="out-kpi-card">
                                <div class="out-kpi-bar"></div>
                                <div class="out-kpi-head">
                                    <div class="out-kpi-ico">📈</div>
                                    <div class="out-kpi-title">本年扫码率</div>
                                </div>
                                <div class="out-kpi-val">{fmt_pct_ratio(rate_year)}</div>
                                <div class="out-kpi-sub"><span>出库(箱)</span><span>{fmt_num(out_year)}</span></div>
                                <div class="out-kpi-sub"><span>扫码(箱)</span><span>{fmt_num(scan_year)}</span></div>
                                <div class="out-kpi-sub2" style="margin-top:10px;"><span>同期({last_year})</span><span>{fmt_num(out_year_last)} 箱 / {fmt_num(scan_year_last)} 箱</span></div>
                                <div class="out-kpi-sub2"><span>同比（扫码率）</span><span class="{_trend_cls(yoy_rate_year)}">{_arrow(yoy_rate_year)} {fmt_pct_value(yoy_rate_year_pct) if yoy_rate_year_pct is not None else "—"}</span></div>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    # --- Sub-Tab 1: Category ---
                    with tab_s_cat:
                        # Group by Big Category
                        
                        # --- Day Level (Sync) ---
                        s_cat_day = df_s_flt[(df_s_flt['年份'] == cur_year) & (df_s_flt['月份'] == cur_month) & (df_s_flt['日'] == cur_day)].groupby('产品大类').size().reset_index(name='本日扫码听数')
                        s_cat_day['本日扫码(箱)'] = s_cat_day['本日扫码听数'] / 6.0
                        o_cat_day = None
                        if out_day_df is not None:
                            if '产品大类' in out_day_df.columns:
                                group_col = '产品大类'
                            elif '大分类' in out_day_df.columns:
                                group_col = '大分类'
                            else:
                                group_col = None
                            qty_col_out = '数量(箱)' if '数量(箱)' in (out_day_df.columns if out_day_df is not None else []) else next((c for c in out_day_df.columns if '数量' in str(c) or '箱' in str(c)), None) if out_day_df is not None else None
                            if group_col and qty_col_out:
                                o_cat_day = out_day_df.groupby(group_col)[qty_col_out].sum().reset_index().rename(columns={group_col: '产品大类', qty_col_out: '今日出库(箱)'})
                        
                        # --- Month Level (Sync) ---
                        s_cat_month = df_s_flt[(df_s_flt['年份'] == cur_year) & (df_s_flt['月份'] == cur_month)].groupby('产品大类').size().reset_index(name='本月扫码听数')
                        s_cat_month['本月扫码(箱)'] = s_cat_month['本月扫码听数'] / 6.0
                        
                        o_cat_month = pd.DataFrame(columns=['产品大类', '本月出库(箱)'])
                        if out_base_df is not None:
                            group_col_m = '产品大类' if '产品大类' in out_base_df.columns else ('大分类' if '大分类' in out_base_df.columns else None)
                            if group_col_m:
                                o_cat_month = out_base_df[(out_base_df['年份'] == cur_year) & (out_base_df['月份'] == cur_month)].groupby(group_col_m)['数量(箱)'].sum().reset_index()
                                o_cat_month = o_cat_month.rename(columns={group_col_m: '产品大类', '数量(箱)': '本月出库(箱)'})

                        # --- Year Level (Sync) ---
                        s_cat_year = df_s_flt[df_s_flt['年份'] == cur_year].groupby('产品大类').size().reset_index(name='本年扫码听数')
                        s_cat_year['本年扫码(箱)'] = s_cat_year['本年扫码听数'] / 6.0
                        
                        o_cat_year = pd.DataFrame(columns=['产品大类', '本年出库(箱)'])
                        if out_base_df is not None:
                            group_col_y = '产品大类' if '产品大类' in out_base_df.columns else ('大分类' if '大分类' in out_base_df.columns else None)
                            if group_col_y:
                                o_cat_year = out_base_df[out_base_df['年份'] == cur_year].groupby(group_col_y)['数量(箱)'].sum().reset_index()
                                o_cat_year = o_cat_year.rename(columns={group_col_y: '产品大类', '数量(箱)': '本年出库(箱)'})
                            
                        # Merge All
                        cat_final = pd.merge(s_cat_day[['产品大类', '本日扫码(箱)']], s_cat_month[['产品大类', '本月扫码(箱)']], on='产品大类', how='outer')
                        if o_cat_day is not None:
                            cat_final = pd.merge(cat_final, o_cat_day, on='产品大类', how='outer')
                        cat_final = pd.merge(cat_final, o_cat_month, on='产品大类', how='outer')
                        cat_final = pd.merge(cat_final, s_cat_year[['产品大类', '本年扫码(箱)']], on='产品大类', how='outer')
                        cat_final = pd.merge(cat_final, o_cat_year, on='产品大类', how='outer').fillna(0)
                        
                        # Calculate Rates
                        # Day Rate: Outbound usually monthly, so Day Rate might not be accurate unless assumed uniform or N/A
                        # User requirement: "本日、本月的维度，也需要加到分品类和分省区". 
                        # Let's show Day Scan Qty. Day Rate is tricky without Day Outbound. We will show Day Scan Qty only or N/A for Rate.
                        # Month Rate
                        cat_final['本月扫码率'] = cat_final.apply(lambda x: x['本月扫码(箱)'] / x['本月出库(箱)'] if x['本月出库(箱)'] > 0 else 0, axis=1)
                        # Year Rate
                        cat_final['本年扫码率'] = cat_final.apply(lambda x: x['本年扫码(箱)'] / x['本年出库(箱)'] if x['本年出库(箱)'] > 0 else 0, axis=1)
                        # Day Rate
                        if '今日出库(箱)' in cat_final.columns:
                            cat_final['本日扫码率'] = cat_final.apply(lambda x: x['本日扫码(箱)'] / x['今日出库(箱)'] if x['今日出库(箱)'] > 0 else 0, axis=1)
                        else:
                            cat_final['今日出库(箱)'] = 0.0
                            cat_final['本日扫码率'] = 0.0
                        
                        cat_final = cat_final.sort_values('本月扫码(箱)', ascending=False)
                        
                        # Format for display
                        # Display
                        cat_disp = cat_final[['产品大类', '今日出库(箱)', '本日扫码(箱)', '本日扫码率', '本月出库(箱)', '本月扫码(箱)', '本月扫码率', '本年出库(箱)', '本年扫码(箱)', '本年扫码率']].copy()
                        cat_disp = cat_disp.rename(columns={'本日扫码(箱)': '今日扫码(箱)'})
                        cat_column_defs = [
                            {"headerName": "产品大类", "field": "产品大类", "pinned": "left", "minWidth": 120},
                            {"headerName": f"今日（{cur_month}月{cur_day}日）", "children": [
                                {"headerName": "出库(箱)", "field": "今日出库(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码(箱)", "field": "今日扫码(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码率", "field": "本日扫码率", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_PCT_RATIO},
                            ]},
                            {"headerName": f"本月（{cur_month}月）", "children": [
                                {"headerName": "出库(箱)", "field": "本月出库(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码(箱)", "field": "本月扫码(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码率", "field": "本月扫码率", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_PCT_RATIO},
                            ]},
                            {"headerName": f"本年（{cur_year}年）", "children": [
                                {"headerName": "出库(箱)", "field": "本年出库(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码(箱)", "field": "本年扫码(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码率", "field": "本年扫码率", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_PCT_RATIO},
                            ]},
                        ]

                        show_aggrid_table(cat_disp, key="scan_cat_ag", column_defs=cat_column_defs)

                    # --- Sub-Tab 2: Province ---
                    with tab_s_prov:
                        # --- Day Level ---
                        s_prov_day = df_s_flt[(df_s_flt['年份'] == cur_year) & (df_s_flt['月份'] == cur_month) & (df_s_flt['日'] == cur_day)].groupby('省区').size().reset_index(name='本日扫码听数')
                        s_prov_day['本日扫码(箱)'] = s_prov_day['本日扫码听数'] / 6.0
                        o_prov_day = None
                        if out_day_df is not None:
                            o_prov_day = out_day_df.groupby('省区')['数量(箱)'].sum().reset_index().rename(columns={'数量(箱)': '今日出库(箱)'})

                        # --- Month Level (Current) ---
                        s_prov_cur = df_s_flt[(df_s_flt['年份'] == cur_year) & (df_s_flt['月份'] == cur_month)].groupby('省区').size().reset_index(name='扫码听数')
                        s_prov_cur['扫码箱数'] = s_prov_cur['扫码听数'] / 6.0
                        o_prov_cur = pd.DataFrame(columns=['省区', '本月出库(箱)'])
                        if out_base_df is not None:
                            o_prov_cur = out_base_df[(out_base_df['年份'] == cur_year) & (out_base_df['月份'] == cur_month)].groupby('省区')['数量(箱)'].sum().reset_index().rename(columns={'数量(箱)': '本月出库(箱)'})
                        prov_cur = pd.merge(s_prov_cur[['省区', '扫码箱数']], o_prov_cur, on='省区', how='outer').fillna(0)
                        prov_cur['本月扫码(箱)'] = prov_cur['扫码箱数']
                        prov_cur['本月扫码率'] = prov_cur.apply(lambda x: x['本月扫码(箱)'] / x['本月出库(箱)'] if x['本月出库(箱)'] > 0 else 0, axis=1)
                        prov_cur = prov_cur[['省区', '本月出库(箱)', '本月扫码(箱)', '本月扫码率']]

                        # --- Same Period Last Year (Month) ---
                        s_prov_last = df_s_flt[(df_s_flt['年份'] == last_year) & (df_s_flt['月份'] == cur_month)].groupby('省区').size().reset_index(name='扫码听数')
                        s_prov_last['扫码箱数'] = s_prov_last['扫码听数'] / 6.0
                        o_prov_last = pd.DataFrame(columns=['省区', '同期出库(箱)'])
                        if out_base_df is not None:
                            o_prov_last = out_base_df[(out_base_df['年份'] == last_year) & (out_base_df['月份'] == cur_month)].groupby('省区')['数量(箱)'].sum().reset_index().rename(columns={'数量(箱)': '同期出库(箱)'})
                        prov_last = pd.merge(s_prov_last[['省区', '扫码箱数']], o_prov_last, on='省区', how='outer').fillna(0)
                        prov_last['同期扫码(箱)'] = prov_last['扫码箱数']
                        prov_last['同期扫码率'] = prov_last.apply(lambda x: x['同期扫码(箱)'] / x['同期出库(箱)'] if x['同期出库(箱)'] > 0 else 0, axis=1)
                        prov_last = prov_last[['省区', '同期出库(箱)', '同期扫码(箱)', '同期扫码率']]

                        # --- Ring Period (Month) ---
                        if cur_month == 1:
                            ring_year = cur_year - 1
                            ring_month = 12
                        else:
                            ring_year = cur_year
                            ring_month = cur_month - 1

                        s_prov_ring = df_s_flt[(df_s_flt['年份'] == ring_year) & (df_s_flt['月份'] == ring_month)].groupby('省区').size().reset_index(name='扫码听数')
                        s_prov_ring['扫码箱数'] = s_prov_ring['扫码听数'] / 6.0
                        o_prov_ring = pd.DataFrame(columns=['省区', '环比出库(箱)'])
                        if out_base_df is not None:
                            o_prov_ring = out_base_df[(out_base_df['年份'] == ring_year) & (out_base_df['月份'] == ring_month)].groupby('省区')['数量(箱)'].sum().reset_index().rename(columns={'数量(箱)': '环比出库(箱)'})
                        prov_ring = pd.merge(s_prov_ring[['省区', '扫码箱数']], o_prov_ring, on='省区', how='outer').fillna(0)
                        prov_ring['环比扫码(箱)'] = prov_ring['扫码箱数']
                        prov_ring['环比扫码率'] = prov_ring.apply(lambda x: x['环比扫码(箱)'] / x['环比出库(箱)'] if x['环比出库(箱)'] > 0 else 0, axis=1)
                        prov_ring = prov_ring[['省区', '环比扫码率']]

                        # Merge All
                        prov_final = pd.merge(prov_cur, s_prov_day[['省区', '本日扫码(箱)']], on='省区', how='outer')
                        if o_prov_day is not None:
                            prov_final = pd.merge(prov_final, o_prov_day, on='省区', how='outer')
                        prov_final = pd.merge(prov_final, prov_last[['省区', '同期出库(箱)', '同期扫码(箱)', '同期扫码率']], on='省区', how='outer')
                        prov_final = pd.merge(prov_final, prov_ring[['省区', '环比扫码率']], on='省区', how='left').fillna(0)
                        prov_final['环比增长'] = prov_final['本月扫码率'] - prov_final['环比扫码率']
                        if '今日出库(箱)' not in prov_final.columns:
                            prov_final['今日出库(箱)'] = 0.0
                        prov_final['本日扫码率'] = prov_final.apply(lambda x: x['本日扫码(箱)'] / x['今日出库(箱)'] if x.get('今日出库(箱)', 0) > 0 else 0, axis=1)

                        prov_disp = prov_final[['省区', '本日扫码(箱)', '今日出库(箱)', '本日扫码率', '本月出库(箱)', '本月扫码(箱)', '本月扫码率', '同期出库(箱)', '同期扫码(箱)', '同期扫码率', '环比扫码率', '环比增长']].copy()
                        prov_disp = prov_disp.sort_values('本月扫码(箱)', ascending=False)
                        prov_disp = prov_disp.rename(columns={'本日扫码(箱)': '今日扫码(箱)'})
                        prov_column_defs = [
                            {"headerName": "省区", "field": "省区", "pinned": "left", "minWidth": 110},
                            {"headerName": f"今日（{cur_month}月{cur_day}日）", "children": [
                                {"headerName": "出库(箱)", "field": "今日出库(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码(箱)", "field": "今日扫码(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码率", "field": "本日扫码率", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_PCT_RATIO},
                            ]},
                            {"headerName": f"本月（{cur_month}月）", "children": [
                                {"headerName": "出库(箱)", "field": "本月出库(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码(箱)", "field": "本月扫码(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码率", "field": "本月扫码率", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_PCT_RATIO},
                            ]},
                            {"headerName": f"同期（{last_year}年{cur_month}月）", "children": [
                                {"headerName": "出库(箱)", "field": "同期出库(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码(箱)", "field": "同期扫码(箱)", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_NUM},
                                {"headerName": "扫码率", "field": "同期扫码率", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_PCT_RATIO},
                            ]},
                            {"headerName": "环比", "children": [
                                {"headerName": "扫码率", "field": "环比扫码率", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_PCT_RATIO},
                                {"headerName": "增长", "field": "环比增长", "type": ["numericColumn", "numberColumnFilter"], "valueFormatter": JS_FMT_PCT_RATIO, "cellStyle": JS_COLOR_CONDITIONAL},
                            ]},
                        ]

                        show_aggrid_table(prov_disp, key="scan_prov_ag", column_defs=prov_column_defs)

                    with tab_s_map:
                        if ("经度" not in df_s_flt.columns) or ("纬度" not in df_s_flt.columns):
                            st.info("未检测到经纬度列：请确认扫码数据Sheet的M列为经纬度（形如 116.4,39.9 或 39.9,116.4）。")
                        else:
                            c_map1, c_map2, c_map3 = st.columns([1.1, 1.1, 1.2])
                            metric_mode = c_map1.radio("对比口径", ["扫码数", "扫码率"], horizontal=True, key="scan_map_metric_mode")
                            period_mode = c_map2.radio("时间范围", ["今日", "本月", "本年"], horizontal=True, key="scan_map_period")
                            style_mode = c_map3.radio("地图样式", ["详细", "简洁"], horizontal=True, key="scan_map_style_mode")

                            c_map4, c_map5 = st.columns([1.3, 1.0])
                            prov_opts_map = ["全国"] + sorted([p for p in df_s_flt["省区"].unique().tolist() if str(p).strip() != ""])
                            focus_prov = c_map4.selectbox("省区聚焦", prov_opts_map, key="scan_map_focus_prov")
                            palette_mode = c_map5.radio("配色", ["高对比", "色盲友好"], horizontal=True, key="scan_map_palette")

                            c_map6, c_map7 = st.columns([1.2, 1.1])
                            basemap_provider = c_map6.selectbox("底图来源", ["高德(国内)", "OpenStreetMap(外网)", "无底图(离线)", "自定义瓦片(内网/自建)"], key="scan_map_basemap_provider")
                            custom_tile_url = ""
                            if basemap_provider == "自定义瓦片(内网/自建)":
                                custom_tile_url = c_map7.text_input("瓦片URL模板", value="http://127.0.0.1:8080/{z}/{x}/{y}.png", key="scan_map_custom_tile_url")
                            else:
                                c_map7.write("")

                            show_cb_key = "scan_map_show_colorbar"
                            if show_cb_key not in st.session_state:
                                st.session_state[show_cb_key] = False
                            cb_label = "显示颜色刻度" if not st.session_state[show_cb_key] else "隐藏颜色刻度"
                            if st.button(cb_label, key="scan_map_toggle_colorbar"):
                                st.session_state[show_cb_key] = not bool(st.session_state[show_cb_key])
                                st.rerun()

                            df_map = df_s_flt.copy()
                            if period_mode == "今日":
                                df_map = df_map[(df_map["年份"] == cur_year) & (df_map["月份"] == cur_month) & (df_map["日"] == cur_day)]
                            elif period_mode == "本月":
                                df_map = df_map[(df_map["年份"] == cur_year) & (df_map["月份"] == cur_month)]
                            else:
                                df_map = df_map[df_map["年份"] == cur_year]

                            if focus_prov != "全国":
                                df_map = df_map[df_map["省区"] == focus_prov]

                            df_map = df_map.dropna(subset=["经度", "纬度"])
                            df_map = df_map[df_map["经度"].between(70, 140) & df_map["纬度"].between(0, 60)]

                            if df_map.empty:
                                st.info("当前筛选与口径下没有可用的经纬度数据。")
                            else:
                                center_lat = float(df_map["纬度"].mean())
                                center_lon = float(df_map["经度"].mean())
                                default_zoom = 3.1 if focus_prov == "全国" else 4.9
                                min_zoom, max_zoom = 2.2, 10.5

                                zoom_key = "scan_map_zoom"
                                if zoom_key not in st.session_state:
                                    st.session_state[zoom_key] = default_zoom
                                if st.session_state[zoom_key] < min_zoom or st.session_state[zoom_key] > max_zoom:
                                    st.session_state[zoom_key] = default_zoom

                                zc1, zc2, zc3, zc4 = st.columns([0.13, 0.13, 0.18, 0.56])
                                if zc1.button("＋", key="scan_map_zoom_in"):
                                    st.session_state[zoom_key] = min(max_zoom, float(st.session_state[zoom_key]) + 0.6)
                                    st.rerun()
                                if zc2.button("－", key="scan_map_zoom_out"):
                                    st.session_state[zoom_key] = max(min_zoom, float(st.session_state[zoom_key]) - 0.6)
                                    st.rerun()
                                if zc3.button("复位", key="scan_map_zoom_reset"):
                                    st.session_state[zoom_key] = default_zoom
                                    st.rerun()
                                zc4.slider("缩放", min_value=min_zoom, max_value=max_zoom, value=float(st.session_state[zoom_key]), step=0.1, key=zoom_key)

                                basemap_layers = None
                                if basemap_provider == "OpenStreetMap(外网)":
                                    map_style = "carto-positron" if style_mode == "简洁" else "open-street-map"
                                elif basemap_provider == "高德(国内)":
                                    map_style = "white-bg"
                                    gaode_style = "7" if style_mode == "详细" else "8"
                                    gaode_url = f"https://webrd02.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style={gaode_style}&x={{x}}&y={{y}}&z={{z}}"
                                    basemap_layers = [{"sourcetype": "raster", "source": [gaode_url], "below": "traces"}]
                                elif basemap_provider == "自定义瓦片(内网/自建)":
                                    map_style = "white-bg"
                                    _u = (custom_tile_url or "").strip()
                                    if _u:
                                        basemap_layers = [{"sourcetype": "raster", "source": [_u], "below": "traces"}]
                                else:
                                    map_style = "white-bg"
                                marker_opacity = 0.86
                                color_scale_count = "Turbo" if palette_mode == "高对比" else "Cividis"
                                color_scale_rate = "Viridis" if palette_mode == "高对比" else "Cividis"
                                point_scale = [
                                    [0.0, "#00C853"],
                                    [0.35, "#00C853"],
                                    [0.65, "#FFEB3B"],
                                    [0.82, "#FF9800"],
                                    [1.0, "#F44336"],
                                ]

                                if metric_mode == "扫码数":
                                    c_u1, c_u2 = st.columns([0.55, 0.45])
                                    unit_mode = c_u1.radio("单位", ["听", "箱"], horizontal=True, key="scan_map_unit")
                                    render_mode = c_u2.radio("渲染方式", ["热力", "标点"], horizontal=True, key="scan_map_render_mode")
                                    precision = st.slider("坐标聚合精度(小数位)", 0, 3, 2, key="scan_map_precision")
                                    df_grid = df_map[["经度", "纬度"]].copy()
                                    df_grid["经度"] = df_grid["经度"].round(int(precision))
                                    df_grid["纬度"] = df_grid["纬度"].round(int(precision))
                                    df_grid = df_grid.groupby(["经度", "纬度"]).size().reset_index(name="扫码听数")
                                    df_grid["扫码箱数"] = df_grid["扫码听数"] / 6.0
                                    val_col = "扫码听数" if unit_mode == "听" else "扫码箱数"

                                    if render_mode == "热力":
                                        fig = px.density_mapbox(
                                            df_grid,
                                            lat="纬度",
                                            lon="经度",
                                            z=val_col,
                                            radius=18 if focus_prov == "全国" else 14,
                                            zoom=float(st.session_state[zoom_key]),
                                            center={"lat": center_lat, "lon": center_lon},
                                            color_continuous_scale=color_scale_count,
                                            hover_data={"扫码听数": ":,.0f", "扫码箱数": ":,.2f"}
                                        )
                                        fig.update_traces(opacity=0.82)
                                    else:
                                        fig = px.scatter_mapbox(
                                            df_grid,
                                            lat="纬度",
                                            lon="经度",
                                            color=val_col,
                                            size=val_col,
                                            size_max=26,
                                            zoom=float(st.session_state[zoom_key]),
                                            center={"lat": center_lat, "lon": center_lon},
                                            color_continuous_scale=point_scale,
                                            hover_data={"扫码听数": ":,.0f", "扫码箱数": ":,.2f"}
                                        )
                                        fig.update_traces(marker={"opacity": marker_opacity})

                                    _layout_kwargs = {
                                        "mapbox_style": map_style,
                                        "margin": {"r": 0, "t": 0, "l": 0, "b": 0},
                                        "transition": {"duration": 260, "easing": "cubic-in-out"},
                                    }
                                    if basemap_layers is not None:
                                        _layout_kwargs["mapbox_layers"] = basemap_layers
                                    fig.update_layout(**_layout_kwargs)
                                    show_cb = bool(st.session_state.get(show_cb_key, False))
                                    cb_style = {"thickness": 10, "len": 0.55, "x": 1.0, "xpad": 0, "y": 0.5, "bgcolor": "rgba(255,255,255,0.25)", "outlinewidth": 0, "title": {"text": ""}}
                                    for _ax_name in [k for k in fig.layout if str(k).startswith("coloraxis")]:
                                        try:
                                            fig.layout[_ax_name].showscale = show_cb
                                        except Exception:
                                            pass
                                        if show_cb:
                                            try:
                                                fig.layout[_ax_name].colorbar = cb_style
                                            except Exception:
                                                pass
                                    for _t in fig.data:
                                        try:
                                            _t.update(showscale=show_cb)
                                        except Exception:
                                            pass
                                    st.plotly_chart(
                                        fig,
                                        use_container_width=True,
                                        config={"scrollZoom": True, "displayModeBar": True, "displaylogo": False, "responsive": True}
                                    )
                                else:
                                    render_mode_rate = st.radio("渲染方式", ["热力", "标点"], horizontal=True, key="scan_map_render_mode_rate")
                                    scan_by_prov = df_map.groupby("省区").size().reset_index(name="扫码听数")
                                    scan_by_prov["扫码箱数"] = scan_by_prov["扫码听数"] / 6.0
                                    cent = df_map.groupby("省区")[["经度", "纬度"]].mean().reset_index()
                                    prov_map = pd.merge(scan_by_prov, cent, on="省区", how="left")
                                    prov_map["出库(箱)"] = 0.0

                                    if out_base_df is not None and not getattr(out_base_df, "empty", True) and ("省区" in out_base_df.columns) and ("数量(箱)" in out_base_df.columns):
                                        out_map = out_base_df.copy()
                                        if period_mode == "今日":
                                            out_map = out_map[(out_map["年份"] == cur_year) & (out_map["月份"] == cur_month) & (out_map["日"] == cur_day)]
                                        elif period_mode == "本月":
                                            out_map = out_map[(out_map["年份"] == cur_year) & (out_map["月份"] == cur_month)]
                                        else:
                                            out_map = out_map[out_map["年份"] == cur_year]
                                        out_prov = out_map.groupby("省区")["数量(箱)"].sum().reset_index().rename(columns={"数量(箱)": "出库(箱)"})
                                        prov_map = pd.merge(prov_map.drop(columns=["出库(箱)"], errors="ignore"), out_prov, on="省区", how="left")
                                        prov_map["出库(箱)"] = pd.to_numeric(prov_map.get("出库(箱)"), errors="coerce").fillna(0.0)

                                    prov_map["扫码率"] = prov_map.apply(lambda x: x["扫码箱数"] / x["出库(箱)"] if x["出库(箱)"] > 0 else None, axis=1)
                                    prov_map = prov_map.dropna(subset=["经度", "纬度"])

                                    if render_mode_rate == "热力":
                                        fig = px.density_mapbox(
                                            prov_map.dropna(subset=["扫码率"]),
                                            lat="纬度",
                                            lon="经度",
                                            z="扫码率",
                                            radius=36 if focus_prov == "全国" else 24,
                                            zoom=float(st.session_state[zoom_key]),
                                            center={"lat": center_lat, "lon": center_lon},
                                            color_continuous_scale=color_scale_rate,
                                            hover_data={"省区": True, "扫码听数": ":,.0f", "扫码箱数": ":,.2f", "出库(箱)": ":,.0f", "扫码率": ":.2%"}
                                        )
                                        fig.update_traces(opacity=0.82)
                                    else:
                                        fig = px.scatter_mapbox(
                                            prov_map,
                                            lat="纬度",
                                            lon="经度",
                                            color="扫码率",
                                            size="扫码箱数",
                                            size_max=42,
                                            zoom=float(st.session_state[zoom_key]),
                                            center={"lat": center_lat, "lon": center_lon},
                                            color_continuous_scale=point_scale,
                                            hover_name="省区",
                                            hover_data={"扫码听数": ":,.0f", "扫码箱数": ":,.2f", "出库(箱)": ":,.0f", "扫码率": ":.2%"}
                                        )
                                        fig.update_traces(marker={"opacity": marker_opacity})
                                    _layout_kwargs = {
                                        "mapbox_style": map_style,
                                        "margin": {"r": 0, "t": 0, "l": 0, "b": 0},
                                        "transition": {"duration": 260, "easing": "cubic-in-out"},
                                    }
                                    if basemap_layers is not None:
                                        _layout_kwargs["mapbox_layers"] = basemap_layers
                                    fig.update_layout(**_layout_kwargs)
                                    show_cb = bool(st.session_state.get(show_cb_key, False))
                                    cb_style = {"thickness": 10, "len": 0.55, "x": 1.0, "xpad": 0, "y": 0.5, "bgcolor": "rgba(255,255,255,0.25)", "outlinewidth": 0, "title": {"text": ""}}
                                    for _ax_name in [k for k in fig.layout if str(k).startswith("coloraxis")]:
                                        try:
                                            fig.layout[_ax_name].showscale = show_cb
                                        except Exception:
                                            pass
                                        if show_cb:
                                            try:
                                                fig.layout[_ax_name].colorbar = cb_style
                                            except Exception:
                                                pass
                                    for _t in fig.data:
                                        try:
                                            _t.update(showscale=show_cb)
                                        except Exception:
                                            pass
                                    st.plotly_chart(
                                        fig,
                                        use_container_width=True,
                                        config={"scrollZoom": True, "displayModeBar": True, "displaylogo": False, "responsive": True}
                                    )

                else:
                    st.info("请在Excel中包含第6个Sheet（扫码数据）以查看此分析。")

            # === TAB 3: ABCD ANALYSIS ===
            with tab3:
                st.subheader("📊 Q3 vs Q4 门店效能对比分析")
                
                # Check for Q3/Q4 columns
                q3_cols = [c for c in month_cols if c in ['7月', '8月', '9月']]
                q4_cols = [c for c in month_cols if c in ['10月', '11月', '12月']]
                
                if not q3_cols or not q4_cols:
                    st.warning("⚠️ 数据源缺失7-12月的完整数据，无法进行Q3 vs Q4对比分析")
                else:
                    # Logic
                    # Calculate Q3 Class
                    df['Q3_Sum'] = df[q3_cols].sum(axis=1)
                    df['Q3_Avg'] = df['Q3_Sum'] / 3
                    
                    # Calculate Q4 Class
                    df['Q4_Sum'] = df[q4_cols].sum(axis=1)
                    df['Q4_Avg'] = df['Q4_Sum'] / 3
                    
                    def classify_score(x):
                        if x >= 4: return 'A'
                        elif 2 <= x < 4: return 'B'
                        elif 1 <= x < 2: return 'C'
                        else: return 'D'
                        
                    df['Class_Q3'] = df['Q3_Avg'].apply(classify_score)
                    df['Class_Q4'] = df['Q4_Avg'].apply(classify_score)
                    
                    # Comparison Metrics
                    q3_counts = df['Class_Q3'].value_counts().sort_index()
                    q4_counts = df['Class_Q4'].value_counts().sort_index()
                    
                    # Overview Cards
                    c1, c2, c3, c4 = st.columns(4)
                    
                    def render_metric(col, cls_label):
                        curr = q4_counts.get(cls_label, 0)
                        prev = q3_counts.get(cls_label, 0)
                        delta = curr - prev
                        col.metric(f"{cls_label}类门店 (Q4)", fmt_num(curr), f"{fmt_num(delta)} (环比)")
                        
                    render_metric(c1, 'A')
                    render_metric(c2, 'B')
                    render_metric(c3, 'C')
                    render_metric(c4, 'D')
                    
                    st.markdown("---")
                    
                    # Province Comparison Chart
                    st.subheader("🗺️ 各省区ABCD类门店数量对比 (Q3 vs Q4)")
                    
                    # Prepare Data for Chart
                    # Group by Province and Class for Q3
                    prov_q3 = df.groupby(['省区', 'Class_Q3']).size().reset_index(name='Count')
                    prov_q3['Period'] = 'Q3'
                    prov_q3.rename(columns={'Class_Q3': 'Class'}, inplace=True)
                    
                    # Group by Province and Class for Q4
                    prov_q4 = df.groupby(['省区', 'Class_Q4']).size().reset_index(name='Count')
                    prov_q4['Period'] = 'Q4'
                    prov_q4.rename(columns={'Class_Q4': 'Class'}, inplace=True)
                    
                    # Combine
                    prov_comp = pd.concat([prov_q3, prov_q4])
                    
                    # Interactive Selection
                    sel_period = st.radio("选择展示周期:", ["Q4 (本期)", "Q3 (上期)"], horizontal=True)
                    target_period = 'Q4' if 'Q4' in sel_period else 'Q3'
                    
                    chart_data = prov_comp[prov_comp['Period'] == target_period]
                    
                    fig_bar_prov_class = px.bar(chart_data, x='省区', y='Count', color='Class',
                                               title=f"各省区门店等级分布 ({target_period})",
                                               category_orders={"Class": ["A", "B", "C", "D"]},
                                               color_discrete_map={'A':'#FFC400', 'B':'#6A3AD0', 'C':'#B79BFF', 'D':'#8A8AA3'},
                                               text='Count')
                    fig_bar_prov_class.update_traces(textposition='inside', texttemplate='%{y:,.1~f}', hovertemplate='省区: %{x}<br>数量: %{y:,.1~f}<extra></extra>')
                    fig_bar_prov_class.update_layout(yaxis_title="门店数量", xaxis_title="省区", yaxis=dict(tickformat=",.1~f"), paper_bgcolor='rgba(255,255,255,0.25)', plot_bgcolor='rgba(255,255,255,0.25)')
                    st.plotly_chart(fig_bar_prov_class, use_container_width=True)
                    
                    st.markdown("---")
                    
                    # Migration Matrix
                    st.subheader("🔄 门店等级变动明细")
                    
                    # Define Change Type
                    def get_change_type(row):
                        order = {'A': 4, 'B': 3, 'C': 2, 'D': 1}
                        score_q3 = order[row['Class_Q3']]
                        score_q4 = order[row['Class_Q4']]
                        
                        if score_q3 == score_q4: return '持平'
                        elif score_q4 > score_q3: return '升级 ⬆️'
                        else: return '降级 ⬇️'
                        
                    df['变动类型'] = df.apply(get_change_type, axis=1)
                    
                    # Summary of Changes
                    change_counts = df['变动类型'].value_counts()
                    st.info(f"📊 变动概览: 升级 {fmt_num(change_counts.get('升级 ⬆️', 0), na='')} 家 | 降级 {fmt_num(change_counts.get('降级 ⬇️', 0), na='')} 家 | 持平 {fmt_num(change_counts.get('持平', 0), na='')} 家")
                    
                    # Detailed Table
                    # Filters
                    c_f1, c_f2, c_f3 = st.columns(3)
                    filter_prov = c_f1.selectbox("筛选省区", ['全部'] + list(df['省区'].unique()), key='abcd_prov')
                    
                    # Distributor Filter (Dependent on Province)
                    if filter_prov != '全部':
                        dist_opts = ['全部'] + sorted(list(df[df['省区'] == filter_prov]['经销商名称'].unique()))
                    else:
                        dist_opts = ['全部'] + sorted(list(df['经销商名称'].unique()))
                    filter_dist = c_f2.selectbox("筛选经销商", dist_opts, key='abcd_dist')
                    
                    filter_change = c_f3.selectbox("筛选变动类型", ['全部', '升级 ⬆️', '降级 ⬇️', '持平'], key='abcd_change')
                    
                    view_df = df.copy()
                    if filter_prov != '全部':
                        view_df = view_df[view_df['省区'] == filter_prov]
                    if filter_dist != '全部':
                        view_df = view_df[view_df['经销商名称'] == filter_dist]
                    if filter_change != '全部':
                        view_df = view_df[view_df['变动类型'] == filter_change]
                        
                    show_aggrid_table(view_df[['省区', '经销商名称', '门店名称', 'Class_Q3', 'Class_Q4', '变动类型', 'Q3_Avg', 'Q4_Avg']])

            with tab_other:
                other_rank, other_query, other_detail, other_review_2025 = st.tabs(["🏆 榜单排名", "🔍 查询分析", "📝 数据明细", "📅 2025年复盘"])
                
                with other_rank:
                    # Initialize df_perf for this scope if not already present
                    if 'df_perf' not in locals():
                         if df_perf_raw is not None:
                             df_perf = df_perf_raw.copy()
                             if '年份' in df_perf.columns:
                                 df_perf['年份'] = pd.to_numeric(df_perf['年份'], errors='coerce').fillna(0).astype(int)
                             if '月份' in df_perf.columns:
                                 df_perf['月份'] = pd.to_numeric(df_perf['月份'], errors='coerce').fillna(0).astype(int)
                             
                             if '发货金额' not in df_perf.columns:
                                     if '发货箱数' in df_perf.columns:
                                         df_perf['发货金额'] = df_perf['发货箱数']
                                     else:
                                         df_perf['发货金额'] = 0.0
                             df_perf['发货金额'] = pd.to_numeric(df_perf['发货金额'], errors='coerce').fillna(0.0)

                             for c in ['省区', '经销商名称', '归类', '发货仓', '大分类', '月分析']:
                                 if c in df_perf.columns:
                                     df_perf[c] = df_perf[c].fillna('').astype(str).str.strip()

                             if '年份' in df_perf.columns and '月份' in df_perf.columns:
                                 df_perf = df_perf[(df_perf['年份'] > 0) & (df_perf['月份'].between(1, 12))]
                                 df_perf['年月'] = pd.to_datetime(
                                     df_perf['年份'].astype(str) + '-' + df_perf['月份'].astype(str).str.zfill(2) + '-01',
                                     errors='coerce'
                                 )
                             else:
                                 df_perf['年月'] = pd.NaT
                         else:
                             df_perf = pd.DataFrame()

                    c_filter, c_main = st.columns([0.26, 0.74])
                    
                    with c_filter:
                        st.markdown("### 🧭 筛选区")
                        
                        # Calculate Date Range from Data
                        if df_perf is not None and not df_perf.empty and '年月' in df_perf.columns:
                            valid_dates = df_perf['年月'].dropna()
                            if not valid_dates.empty:
                                max_ym = valid_dates.max()
                                min_ym = valid_dates.min()
                            else:
                                max_ym = pd.Timestamp.now()
                                min_ym = max_ym - pd.DateOffset(months=12)
                        else:
                            max_ym = pd.Timestamp.now()
                            min_ym = max_ym - pd.DateOffset(months=12)

                        time_mode = st.selectbox(
                            "时间范围",
                            ["近3个月", "近6个月", "近12个月", "自定义年月"],
                            index=["近3个月", "近6个月", "近12个月", "自定义年月"].index(st.session_state.perf_time_mode) if st.session_state.perf_time_mode in ["近3个月", "近6个月", "近12个月", "自定义年月"] else 2,
                            key="perf_time_mode"
                        )

                        # Initialize default values before any condition
                        # Use pd.Timestamp to ensure correct type for comparisons
                        start_ym = pd.Timestamp(max_ym.year, max_ym.month, 1) - pd.DateOffset(months=11)
                        end_ym = pd.Timestamp(max_ym.year, max_ym.month, 1)

                        def _months_back(n):
                            end = pd.Timestamp(max_ym.year, max_ym.month, 1)
                            start = (end - pd.DateOffset(months=n - 1))
                            return start, end

                        if time_mode == "近3个月":
                            start_ym, end_ym = _months_back(3)
                        elif time_mode == "近6个月":
                            start_ym, end_ym = _months_back(6)
                        elif time_mode == "近12个月":
                            start_ym, end_ym = _months_back(12)
                        else:
                                c_from, c_to = st.columns(2)
                                with c_from:
                                    start_ym = st.date_input("开始月", value=pd.Timestamp(max_ym.year, max_ym.month, 1) - pd.DateOffset(months=11), min_value=min_ym.date(), max_value=max_ym.date(), key="perf_start")
                                with c_to:
                                    end_ym = st.date_input("结束月", value=max_ym.date(), min_value=min_ym.date(), max_value=max_ym.date(), key="perf_end")
                                start_ym = pd.Timestamp(pd.to_datetime(start_ym).year, pd.to_datetime(start_ym).month, 1)
                                end_ym = pd.Timestamp(pd.to_datetime(end_ym).year, pd.to_datetime(end_ym).month, 1)

                        prov_col = df_perf.get('省区', pd.Series(dtype=str))
                        prov_opts = sorted(prov_col.dropna().astype(str).str.strip().unique().tolist())
                        selected_provs = st.multiselect("省区（多选）", prov_opts, default=prov_opts if not st.session_state.perf_provs else st.session_state.perf_provs, key="perf_provs")

                        wh_opts = sorted([x for x in df_perf.get('发货仓', pd.Series(dtype=str)).dropna().astype(str).str.strip().unique().tolist() if x])
                        wh_sel = st.selectbox("发货仓", ["全部"] + wh_opts, index=0, key="perf_wh")

                        mid_opts = sorted([x for x in df_perf.get('中类', pd.Series(dtype=str)).dropna().astype(str).str.strip().unique().tolist() if x])
                        mid_sel = st.selectbox("中类", ["全部"] + mid_opts, index=0, key="perf_mid")

                        grp_opts = sorted([x for x in df_perf.get('归类', pd.Series(dtype=str)).dropna().astype(str).str.strip().unique().tolist() if x])
                        grp_sel = st.selectbox("归类", ["全部"] + grp_opts, index=0, key="perf_grp")

                        cat_col = '类目' if '类目' in df_perf.columns else ('大类' if '大类' in df_perf.columns else None)
                        cat_opts = sorted([x for x in df_perf.get(cat_col, pd.Series(dtype=str)).dropna().astype(str).str.strip().unique().tolist() if x]) if cat_col else []
                        
                        default_cats = []
                        if st.session_state.perf_cats:
                            default_cats = [c for c in st.session_state.perf_cats if c in cat_opts]
                        else:
                            default_cats = cat_opts
                            
                        cat_sel = st.multiselect("类目（多选）", cat_opts, default=default_cats, key="perf_cats")

                        dist_opts = sorted([x for x in df_perf.get('经销商名称', pd.Series(dtype=str)).dropna().astype(str).str.strip().unique().tolist() if x])
                        dist_sel = st.multiselect("经销商（可选）", dist_opts, default=[], key="perf_dists")

                        df_f = df_perf.copy()
                        df_f = df_f[(df_f['年月'] >= pd.Timestamp(start_ym)) & (df_f['年月'] <= pd.Timestamp(end_ym))]
                        if selected_provs:
                            df_f = df_f[df_f['省区'].astype(str).isin([str(x) for x in selected_provs])]
                        if wh_sel != "全部" and '发货仓' in df_f.columns:
                            df_f = df_f[df_f['发货仓'].astype(str) == str(wh_sel)]
                        if mid_sel != "全部" and '中类' in df_f.columns:
                            df_f = df_f[df_f['中类'].astype(str) == str(mid_sel)]
                        if grp_sel != "全部" and '归类' in df_f.columns:
                            df_f = df_f[df_f['归类'].astype(str) == str(grp_sel)]
                        if cat_col and cat_sel:
                            df_f = df_f[df_f[cat_col].astype(str).isin([str(x) for x in cat_sel])]
                        if dist_sel:
                            df_f = df_f[df_f['经销商名称'].astype(str).isin([str(x) for x in dist_sel])]

                        months_in_scope = sorted(df_f['年月'].dropna().unique().tolist())
                        months_n = len(months_in_scope) if months_in_scope else 0

                        def _sum_by_month(_df):
                            return _df.groupby('年月', as_index=False)['发货箱数'].sum().rename(columns={'发货箱数': '实际'})

                        actual_total = float(df_f['发货箱数'].sum()) if '发货箱数' in df_f.columns else 0.0

                        base_start = pd.Timestamp(start_ym) - pd.DateOffset(years=1)
                        base_end = pd.Timestamp(end_ym) - pd.DateOffset(years=1)
                        df_base = df_perf.copy()
                        df_base = df_base[(df_base['年月'] >= base_start) & (df_base['年月'] <= base_end)]
                        if selected_provs:
                            df_base = df_base[df_base['省区'].astype(str).isin([str(x) for x in selected_provs])]
                        if wh_sel != "全部" and '发货仓' in df_base.columns:
                            df_base = df_base[df_base['发货仓'].astype(str) == str(wh_sel)]
                        if mid_sel != "全部" and '中类' in df_base.columns:
                            df_base = df_base[df_base['中类'].astype(str) == str(mid_sel)]
                        if grp_sel != "全部" and '归类' in df_base.columns:
                            df_base = df_base[df_base['归类'].astype(str) == str(grp_sel)]
                        if cat_col and cat_sel:
                            df_base = df_base[df_base[cat_col].astype(str).isin([str(x) for x in cat_sel])]
                        if dist_sel:
                            df_base = df_base[df_base['经销商名称'].astype(str).isin([str(x) for x in dist_sel])]
                        plan_total = float(df_base['发货箱数'].sum()) if '发货箱数' in df_base.columns else 0.0

                        yoy_pct = None
                        if plan_total > 0:
                            yoy_pct = (actual_total - plan_total) / plan_total

                        prev_start = pd.Timestamp(start_ym) - pd.DateOffset(months=months_n) if months_n else pd.Timestamp(start_ym) - pd.DateOffset(months=12)
                        prev_end = pd.Timestamp(end_ym) - pd.DateOffset(months=months_n) if months_n else pd.Timestamp(end_ym) - pd.DateOffset(months=12)
                        df_prev = df_perf.copy()
                        df_prev = df_prev[(df_prev['年月'] >= prev_start) & (df_prev['年月'] <= prev_end)]
                        if selected_provs:
                            df_prev = df_prev[df_prev['省区'].astype(str).isin([str(x) for x in selected_provs])]
                        if wh_sel != "全部" and '发货仓' in df_prev.columns:
                            df_prev = df_prev[df_prev['发货仓'].astype(str) == str(wh_sel)]
                        if mid_sel != "全部" and '中类' in df_prev.columns:
                            df_prev = df_prev[df_prev['中类'].astype(str) == str(mid_sel)]
                        if grp_sel != "全部" and '归类' in df_prev.columns:
                            df_prev = df_prev[df_prev['归类'].astype(str) == str(grp_sel)]
                        if cat_col and cat_sel:
                            df_prev = df_prev[df_prev[cat_col].astype(str).isin([str(x) for x in cat_sel])]
                        if dist_sel:
                            df_prev = df_prev[df_prev['经销商名称'].astype(str).isin([str(x) for x in dist_sel])]
                        prev_total = float(df_prev['发货箱数'].sum()) if '发货箱数' in df_prev.columns else 0.0

                        mom_pct = None
                        if prev_total > 0:
                            mom_pct = (actual_total - prev_total) / prev_total
        
                    with c_main:
                        st.subheader("🏪 TOP 10 门店")
                        store_rank = df.nlargest(10, '总出库数')[['门店名称', '总出库数', '省区']]
                        fig_store = px.bar(store_rank, x='总出库数', y='门店名称', orientation='h', text='总出库数',
                                          title="门店出库排行 (前10)", color='省区')
                        fig_store.update_traces(texttemplate='%{x:,.1~f}', hovertemplate='门店: %{y}<br>总出库数: %{x:,.1~f}<extra></extra>')
                        fig_store.update_layout(yaxis_title="", yaxis={'categoryorder':'total ascending'}, xaxis=dict(tickformat=",.1~f"))
                        st.plotly_chart(fig_store, use_container_width=True)
                        
                        st.subheader("🌍 省区排名")
                        prov_rank = df.groupby('省区')['总出库数'].sum().sort_values(ascending=False).reset_index()
                        prov_rank['总出库数'] = prov_rank['总出库数'].astype(int)
                        n_rows = len(prov_rank)
                        calc_height = (n_rows + 1) * 35 + 10
                        final_height = max(150, min(calc_height, 2000))
                        show_aggrid_table(
                            prov_rank,
                            height=final_height,
                            columns_props={'总出库数': {'type': 'bar'}}
                        )

                with other_query:
                    st.subheader("🔍 多维度查询分析")
                    
                    sc1, sc2, sc3 = st.columns(3)
                    search_provinces = ['全部'] + sorted(list(df['省区'].unique()))
                    s_prov = sc1.selectbox("选择省区 (Province)", search_provinces, key='s_prov')
                    
                    if s_prov != '全部':
                        s_dist_opts = ['全部'] + sorted(list(df[df['省区'] == s_prov]['经销商名称'].unique()))
                    else:
                        s_dist_opts = ['全部'] + sorted(list(df['经销商名称'].unique()))
                    s_dist = sc2.selectbox("选择经销商 (Distributor)", s_dist_opts, key='s_dist')
                    
                    df_store_filter = df.copy()
                    if s_prov != '全部':
                        df_store_filter = df_store_filter[df_store_filter['省区'] == s_prov]
                    if s_dist != '全部':
                        df_store_filter = df_store_filter[df_store_filter['经销商名称'] == s_dist]
                        
                    s_store_opts = ['全部'] + sorted(list(df_store_filter['门店名称'].unique()))
                    s_store = sc3.selectbox("选择门店 (Store)", s_store_opts, key='s_store')

                    st.markdown("---")
                    
                    if s_store != '全部':
                        store_row = df_store_filter[df_store_filter['门店名称'] == s_store].iloc[0]
                        st.markdown(f"### 🏪 门店详情: {s_store}")
                        st.caption(f"所属经销商: {store_row['经销商名称']} | 所属省区: {store_row['省区']}")
                        
                        if month_cols:
                            row_trend = pd.DataFrame({'月份': month_cols, '出库数': [store_row[c] for c in month_cols]})
                            row_trend['Month_Num'] = row_trend['月份'].str.extract(r'(\d+)')[0].astype(int)
                            row_trend = row_trend.sort_values('Month_Num')
                            fig_s = px.line(row_trend, x='月份', y='出库数', markers=True, text='出库数', title=f"{s_store} - 月度出库趋势")
                            fig_s.update_traces(
                                mode='lines+markers+text',
                                line_color='#6A3AD0',
                                line_width=3,
                                hovertemplate='月份: %{x}<br>出库数: %{y:,.1~f}<extra></extra>',
                                texttemplate='%{y:,.1~f}',
                                textposition="top center"
                            )
                            fig_s.update_layout(yaxis=dict(tickformat=",.1~f"), paper_bgcolor='rgba(255,255,255,0.25)', plot_bgcolor='rgba(255,255,255,0.25)')
                            st.plotly_chart(fig_s, use_container_width=True)
                            show_aggrid_table(pd.DataFrame([store_row]), height=150, key="s_store_table")

                    elif s_dist != '全部':
                        st.markdown(f"### 🏢 经销商详情: {s_dist}")
                        dist_sub = df[df['经销商名称'] == s_dist]
                        st.caption(f"覆盖省区: {', '.join(dist_sub['省区'].unique())} | 旗下门店数: {len(dist_sub)}")
                        
                        if month_cols:
                            dist_trend = pd.DataFrame({'月份': month_cols, '出库数': dist_sub[month_cols].sum().values})
                            dist_trend['Month_Num'] = dist_trend['月份'].str.extract(r'(\d+)')[0].astype(int)
                            dist_trend = dist_trend.sort_values('Month_Num')
                            fig_d = px.line(dist_trend, x='月份', y='出库数', markers=True, text='出库数', title=f"{s_dist} - 整体月度出库趋势")
                            fig_d.update_traces(
                                mode='lines+markers+text',
                                line_color='#FFC400',
                                line_width=3,
                                hovertemplate='月份: %{x}<br>合计出库: %{y:,.1~f}<extra></extra>',
                                texttemplate='%{y:,.1~f}',
                                textposition="top center"
                            )
                            fig_d.update_layout(yaxis=dict(tickformat=",.1~f"), paper_bgcolor='rgba(255,255,255,0.25)', plot_bgcolor='rgba(255,255,255,0.25)')
                            st.plotly_chart(fig_d, use_container_width=True)
                            st.markdown("#### 旗下门店列表")
                            show_aggrid_table(dist_sub[['省区', '门店名称', '总出库数', '门店分类']], height=300, key="s_dist_table")

                    elif s_prov != '全部':
                        st.markdown(f"### 🏙️ 省区详情: {s_prov}")
                        prov_sub = df[df['省区'] == s_prov]
                        st.caption(f"经销商数量: {prov_sub['经销商名称'].nunique()} | 门店总数: {len(prov_sub)}")
                        
                        if month_cols:
                            prov_trend = pd.DataFrame({'月份': month_cols, '出库数': prov_sub[month_cols].sum().values})
                            prov_trend['Month_Num'] = prov_trend['月份'].str.extract(r'(\d+)')[0].astype(int)
                            prov_trend = prov_trend.sort_values('Month_Num')
                            fig_p = px.line(prov_trend, x='月份', y='出库数', markers=True, text='出库数', title=f"{s_prov} - 全省月度出库趋势")
                            fig_p.update_traces(
                                mode='lines+markers+text',
                                line_color='#5B2EA6',
                                line_width=3,
                                hovertemplate='月份: %{x}<br>合计出库: %{y:,.1~f}<extra></extra>',
                                texttemplate='%{y:,.1~f}',
                                textposition="top center"
                            )
                            fig_p.update_layout(yaxis=dict(tickformat=",.1~f"), paper_bgcolor='rgba(255,255,255,0.25)', plot_bgcolor='rgba(255,255,255,0.25)')
                            st.plotly_chart(fig_p, use_container_width=True)
                            st.markdown("#### 省内经销商概览")
                            dist_summary = prov_sub.groupby('经销商名称')['总出库数'].sum().reset_index().sort_values('总出库数', ascending=False)
                            show_aggrid_table(dist_summary, height=400, key="s_prov_table")
                    else:
                        st.info("👈 请在上方选择 省区 / 经销商 / 门店 进行查询")

                with other_detail:
                    st.subheader("📝 数据明细")
                    ds_opts = ["门店出库汇总(Sheet1)", "库存明细(Sheet2)", "出库底表(Sheet3)", "发货业绩(Sheet4)", "任务表(Sheet5)"]
                    ds = st.selectbox("选择数据集", ds_opts, key="other_ds_sel")

                    df_show = None
                    if ds.startswith("门店出库"):
                        df_show = df.copy()
                    elif ds.startswith("库存") and (df_stock_raw is not None):
                        df_show = df_stock_raw.copy()
                    elif ds.startswith("出库底表") and (df_q4_raw is not None):
                        df_show = df_q4_raw.copy()
                    elif ds.startswith("发货业绩") and (df_perf_raw is not None):
                        df_show = df_perf_raw.copy()
                    elif ds.startswith("任务表") and (df_target_raw is not None):
                        df_show = df_target_raw.copy()

                    if df_show is None or df_show.empty:
                        st.info("当前数据集无数据。")
                    else:
                        show_aggrid_table(df_show, height=520, key="other_detail_table")
                        out_buf = io.BytesIO()
                        try:
                            with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
                                df_show.to_excel(writer, index=False, sheet_name='data')
                        except Exception:
                            with pd.ExcelWriter(out_buf, engine='xlsxwriter') as writer:
                                df_show.to_excel(writer, index=False, sheet_name='data')
                        st.download_button(
                            "📥 导出当前数据集（Excel）",
                            data=out_buf.getvalue(),
                            file_name=f"{ds}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="other_detail_download"
                        )
                
                # --- New Tab for 2025 Review ---
                with other_review_2025:
                     st.header("📅 2025年复盘 (2025 Review)")
                     
                     # Construct df_2025 from df_perf_raw and df_target_raw
                     # Needs: '省区', '实际发货', '年度任务', '同比增长', '产品品类'
                     df_2025 = None
                     if df_perf_raw is not None and df_target_raw is not None:
                         # 1. Actuals 2025
                         # Ensure numeric
                         if '年份' in df_perf_raw.columns:
                            perf_2025 = df_perf_raw[df_perf_raw['年份'] == 2025].copy()
                         else:
                            perf_2025 = pd.DataFrame()

                         if not perf_2025.empty:
                            if '发货金额' not in perf_2025.columns and '发货箱数' in perf_2025.columns:
                                perf_2025['发货金额'] = perf_2025['发货箱数'] # Fallback
                            
                            # Agg by Prov
                            act_prov = perf_2025.groupby('省区')['发货金额'].sum().reset_index().rename(columns={'发货金额': '实际发货'})
                            
                            # 2. Targets 2025 (Assuming df_target_raw is 2025 targets)
                            # Target sheet: C=Prov, D=Cat, E=Month, F=Task
                            # We need to sum Task by Prov
                            tgt_prov = df_target_raw.groupby('省区')['任务量'].sum().reset_index().rename(columns={'任务量': '年度任务'})
                            
                            # 3. Merge
                            df_2025 = pd.merge(act_prov, tgt_prov, on='省区', how='outer').fillna(0)
                            
                            # 4. YoY (Need 2024 data)
                            perf_2024 = df_perf_raw[df_perf_raw['年份'] == 2024].copy()
                            if not perf_2024.empty:
                                if '发货金额' not in perf_2024.columns and '发货箱数' in perf_2024.columns:
                                    perf_2024['发货金额'] = perf_2024['发货箱数']
                                act_2024 = perf_2024.groupby('省区')['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期'})
                                df_2025 = pd.merge(df_2025, act_2024, on='省区', how='left').fillna(0)
                                df_2025['同比增长'] = df_2025.apply(lambda x: ((x['实际发货'] - x['同期']) / x['同期']) if x['同期'] > 0 else 0, axis=1)
                            else:
                                df_2025['同比增长'] = 0.0
                                
                            # 5. Category Breakdown (Optional, if '产品品类' needed)
                            # Create a separate df for category view if needed, or try to add it to df_2025?
                            # The code expects df_2025 to have '产品品类' column if possible. 
                            # But df_2025 above is aggregated by Province. 
                            # If we want Category breakdown, we need a different aggregation.
                            # Let's check usage: 
                            # prov_summ = df_2025.groupby('省区')... -> This works on the prov-agg df
                            # cat_summ = df_2025.groupby('产品品类')... -> This implies df_2025 should be granular?
                            # If df_2025 is granular (Prov, Cat), we can do both.
                            
                            # Let's try to build granular df_2025 (Prov, Cat)
                            cat_col_p = '类目' if '类目' in perf_2025.columns else ('大类' if '大类' in perf_2025.columns else '大分类')
                            if cat_col_p not in perf_2025.columns: cat_col_p = '省区' # Fallback
                            
                            act_gran = perf_2025.groupby(['省区', cat_col_p])['发货金额'].sum().reset_index().rename(columns={'发货金额': '实际发货', cat_col_p: '产品品类'})
                            
                            # Target Granular
                            # Sheet 5: '品类' column exists?
                            if '品类' in df_target_raw.columns:
                                tgt_gran = df_target_raw.groupby(['省区', '品类'])['任务量'].sum().reset_index().rename(columns={'任务量': '年度任务', '品类': '产品品类'})
                            else:
                                tgt_gran = pd.DataFrame(columns=['省区', '产品品类', '年度任务'])
                                
                            df_2025_g = pd.merge(act_gran, tgt_gran, on=['省区', '产品品类'], how='outer').fillna(0)
                            
                            # YoY Granular
                            if not perf_2024.empty:
                                cat_col_24 = '类目' if '类目' in perf_2024.columns else ('大类' if '大类' in perf_2024.columns else '大分类')
                                if cat_col_24 not in perf_2024.columns: cat_col_24 = '省区'
                                act_2024_g = perf_2024.groupby(['省区', cat_col_24])['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期', cat_col_24: '产品品类'})
                                df_2025_g = pd.merge(df_2025_g, act_2024_g, on=['省区', '产品品类'], how='left').fillna(0)
                                df_2025_g['同比增长'] = df_2025_g.apply(lambda x: ((x['实际发货'] - x['同期']) / x['同期']) if x['同期'] > 0 else 0, axis=1)
                            else:
                                df_2025_g['同比增长'] = 0.0
                                
                            df_2025 = df_2025_g

                     if df_2025 is None or df_2025.empty:
                         st.warning("⚠️ 未找到 2025 年复盘数据 (Sheet2)。请检查上传文件。")
                     else:
                         # 1. Total KPI
                         st.subheader("1. 整体关键指标")
                         c1, c2, c3, c4 = st.columns(4)
                         total_sales = df_2025['实际发货'].sum()
                         total_target = df_2025['年度任务'].sum()
                         ach_rate = total_sales / total_target if total_target else 0
                         yoy_growth = df_2025['同比增长'].mean()  # This might need weighted avg
                         
                         c1.metric("2025总实际发货", fmt_num(total_sales), delta=fmt_pct_value(yoy_growth))
                         c2.metric("2025总年度任务", fmt_num(total_target))
                         c3.metric("年度达成率", fmt_pct_ratio(ach_rate))
                         
                         # 2. Province Performance
                         st.subheader("2. 省区表现概览")
                         prov_summ = df_2025.groupby('省区')[['年度任务', '实际发货', '同比增长']].sum().reset_index()
                         prov_summ['达成率'] = prov_summ['实际发货'] / prov_summ['年度任务']
                         prov_summ = prov_summ.sort_values('实际发货', ascending=False)
                         
                         show_aggrid_table(prov_summ, height=400, key="review_2025_prov")
                         
                         # 3. Category Breakdown (if available)
                         if '产品品类' in df_2025.columns:
                             st.subheader("3. 品类表现")
                             cat_summ = df_2025.groupby('产品品类')[['实际发货', '年度任务']].sum().reset_index()
                             cat_summ['达成率'] = cat_summ['实际发货'] / cat_summ['年度任务']
                             
                             c_chart, c_data = st.columns([2, 1])
                             with c_chart:
                                 fig_cat = px.bar(cat_summ, x='产品品类', y=['实际发货', '年度任务'], barmode='group', title="品类任务 vs 实际")
                                 st.plotly_chart(fig_cat, use_container_width=True)
                             with c_data:
                                 show_aggrid_table(cat_summ, height=300, key="review_2025_cat")


            # --- Tab 6: Inventory Analysis ---
            with tab6:
                if df_stock_raw is None:
                    st.warning("⚠️ 未检测到库存数据 (Sheet2)。请确保上传的 Excel 文件包含第二个 Sheet 页，且格式正确。")
                    st.info("数据格式要求：\nSheet2 需包含 A-L 列，顺序为：经销商编码、经销商名称、产品编码、产品名称、库存数量、箱数、省区名称、客户简称、产品大类、产品小类、重量、规格。")
                else:
                    st.caption(f"🕒 数据更新时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    
                    with st.expander("🛠️ 库存筛选", expanded=False):
                        # Prepare filter lists
                        stock_provs = ['全部'] + sorted(list(df_stock_raw['省区名称'].dropna().unique()))
                        stock_dists = ['全部'] + sorted(list(df_stock_raw['经销商名称'].dropna().unique()))
                        stock_cats = ['全部'] + sorted(list(df_stock_raw['产品大类'].dropna().unique()))
                        
                        # Helper to reset drill status
                        def reset_inv_drill():
                            st.session_state.drill_level = 1
                            st.session_state.selected_prov = None
                            st.session_state.selected_dist = None

                        # --- Subcategory Logic Adjustment ---
                        # User requirement: "Subcategory" dropdown should include 'Segment' and 'Ya Series'.
                        # When 'Segment' is selected, Specific Category options are ['1段', '2段', '3段'].
                        # When 'Ya Series' is selected, Specific Category options are ['雅赋', '雅耀', '雅舒', '雅护'].
                        
                        # 1. Base Subcategories
                        base_subcats = sorted(list(df_stock_raw['产品小类'].dropna().unique()))
                        # 2. Add Virtual Subcategories (Ensure uniqueness)
                        virtual_subcats = ['分段', '雅系列']
                        stock_subcats = ['全部'] + virtual_subcats + [s for s in base_subcats if s not in virtual_subcats]
                        
                        c1, c2, c3, c4, c5 = st.columns(5)
                        with c1: s_prov = st.selectbox("省区名称", stock_provs, key='stock_s_prov', on_change=reset_inv_drill)
                        with c2: 
                            if s_prov != '全部':
                                valid_dists = df_stock_raw[df_stock_raw['省区名称'] == s_prov]['经销商名称'].unique()
                                s_dist_opts = ['全部'] + sorted(list(valid_dists))
                            else:
                                s_dist_opts = stock_dists
                            s_dist = st.selectbox("经销商名称", s_dist_opts, key='stock_s_dist', on_change=reset_inv_drill)
                            
                        with c3: s_cat = st.selectbox("产品大类", stock_cats, key='stock_s_cat', on_change=reset_inv_drill)
                        
                        with c4: 
                            # Dynamic filter for subcat based on cat
                            # If we are using virtual subcats, we might want to show them regardless of Category?
                            # Or only if the Category allows? Assuming '美思雅段粉' allows them.
                            if s_cat != '全部':
                                valid_sub = df_stock_raw[df_stock_raw['产品大类'] == s_cat]['产品小类'].unique()
                                # Mix in virtuals if they make sense (assuming they are always available for filtering)
                                current_sub_opts = ['全部'] + virtual_subcats + sorted([s for s in valid_sub if s not in virtual_subcats])
                                s_sub_opts = current_sub_opts
                            else:
                                s_sub_opts = stock_subcats
                            if 'stock_s_sub' in st.session_state:
                                st.session_state.pop('stock_s_sub', None)
                            s_sub_selected = st.multiselect("产品小类(可多选)", s_sub_opts, default=['全部'], key='stock_s_sub_ms', on_change=reset_inv_drill)
                        
                        with c5:
                            # --- Dynamic Specific Category Options based on Subcategory Selection ---
                            if '分段' in s_sub_selected and '雅系列' in s_sub_selected:
                                stock_specs = ['1段', '2段', '3段', '雅赋', '雅耀', '雅舒', '雅护']
                            elif '分段' in s_sub_selected:
                                stock_specs = ['1段', '2段', '3段']
                            elif '雅系列' in s_sub_selected:
                                stock_specs = ['雅赋', '雅耀', '雅舒', '雅护']
                            else:
                                raw_specs = df_stock_raw['具体分类'].dropna().unique()
                                spec_opts = set(raw_specs)
                                stock_specs = sorted(list(spec_opts))
                                
                            s_spec = st.multiselect("具体分类 (支持多选)", stock_specs, default=[], placeholder="选择具体分类...", on_change=reset_inv_drill)
                        
                        # Apply Filters
                        df_s_filtered = df_stock_raw.copy()
                        if s_prov != '全部': df_s_filtered = df_s_filtered[df_s_filtered['省区名称'] == s_prov]
                        if s_dist != '全部': df_s_filtered = df_s_filtered[df_s_filtered['经销商名称'] == s_dist]
                        if s_cat != '全部': df_s_filtered = df_s_filtered[df_s_filtered['产品大类'] == s_cat]
                        
                        # --- Subcategory Filter Logic ---
                        if s_sub_selected and ('全部' not in s_sub_selected):
                            mask_sub = pd.Series(False, index=df_s_filtered.index)
                            if '分段' in s_sub_selected:
                                mask_sub = mask_sub | (
                                    (df_s_filtered['产品大类'].astype(str) == '美思雅段粉')
                                    & (df_s_filtered['具体分类'].fillna('').astype(str).isin(['1段', '2段', '3段']))
                                )
                            if '雅系列' in s_sub_selected:
                                mask_sub = mask_sub | (
                                    df_s_filtered['具体分类'].fillna('').astype(str).isin(['雅赋', '雅耀', '雅舒', '雅护'])
                                )
                            normal_subs = [x for x in s_sub_selected if x not in ['分段', '雅系列', '全部']]
                            if normal_subs:
                                mask_sub = mask_sub | df_s_filtered['产品小类'].astype(str).isin([str(x) for x in normal_subs])
                            df_s_filtered = df_s_filtered[mask_sub]
                        
                        # Apply Specific Category Filter
                        if s_spec:
                            def match_spec(row_val):
                                row_val = str(row_val)
                                for sel in s_spec:
                                    if sel in row_val: return True
                                return False
                            
                            mask = df_s_filtered['具体分类'].apply(match_spec)
                            df_s_filtered = df_s_filtered[mask]
                    
                    st.markdown("---")

                    outbound_pivot = pd.DataFrame()
                    df_o_filtered = pd.DataFrame()
                    sales_agg_q4 = pd.DataFrame(columns=['经销商名称', 'Q4_Total', 'Q4_Avg'])

                    with st.expander("🚚 出库筛选", expanded=False):
                        if df_q4_raw is None or df_q4_raw.empty:
                            st.warning("⚠️ 未检测到出库底表数据 (Sheet3)。")
                        else:
                            o_raw = df_q4_raw.copy()
                            required_out_cols = ['省区', '经销商名称', '数量(箱)', '月份']
                            missing_out = [c for c in required_out_cols if c not in o_raw.columns]

                            if missing_out:
                                st.warning(f"⚠️ 出库底表缺失字段：{', '.join(missing_out)}")
                            else:
                                if '产品大类' not in o_raw.columns:
                                    o_raw['产品大类'] = '全部'
                                if '产品小类' not in o_raw.columns:
                                    o_raw['产品小类'] = '全部'
                                else:
                                    o_raw['产品小类'] = o_raw['产品小类'].astype(str).str.strip()
                                    o_raw.loc[o_raw['产品小类'].isin(['', 'nan', 'None', 'NULL', 'NaN']), '产品小类'] = pd.NA

                            out_provs = ['全部'] + sorted(o_raw['省区'].dropna().astype(str).unique().tolist())
                            out_dists_all = ['全部'] + sorted(o_raw['经销商名称'].dropna().astype(str).unique().tolist())
                            out_cats = ['全部'] + sorted(o_raw['产品大类'].dropna().astype(str).unique().tolist())
                            out_subs_clean = o_raw['产品小类'].dropna().astype(str).str.strip()
                            out_subs_clean = out_subs_clean[out_subs_clean != '']
                            out_subs_list = sorted(out_subs_clean.unique().tolist())
                            out_subs = ['全部'] + out_subs_list
                            empty_sub_cnt = int(o_raw['产品小类'].isna().sum()) if '产品小类' in o_raw.columns else 0
                            dup_sub_cnt = int(out_subs_clean.shape[0] - out_subs_clean.nunique())
                            if empty_sub_cnt > 0:
                                st.warning(f"⚠️ Sheet3 的M列(产品小类)存在空值：{empty_sub_cnt} 行")
                            if dup_sub_cnt > 0:
                                st.info(f"ℹ️ Sheet3 的M列(产品小类)存在重复值：{dup_sub_cnt} 行（下拉已自动去重）")
                            out_month_opts = list(range(1, 13))

                            oc1, oc2, oc3, oc4, oc5 = st.columns(5)
                            with oc1:
                                o_prov = st.selectbox("省区", out_provs, key='out_s_prov')
                            with oc2:
                                if o_prov != '全部':
                                    dists_in_prov = o_raw[o_raw['省区'].astype(str) == str(o_prov)]['经销商名称'].dropna().astype(str).unique().tolist()
                                    out_dists = ['全部'] + sorted(dists_in_prov)
                                else:
                                    out_dists = out_dists_all
                                o_dist = st.selectbox("经销商", out_dists, key='out_s_dist')
                            with oc3:
                                o_cat = st.selectbox("产品大类", out_cats, key='out_s_cat')
                            with oc4:
                                if o_cat != '全部':
                                    subs_in_cat = o_raw[o_raw['产品大类'].astype(str) == str(o_cat)]['产品小类'].dropna().astype(str).unique().tolist()
                                    out_subs2 = ['全部'] + sorted(subs_in_cat)
                                else:
                                    out_subs2 = out_subs
                                if 'out_s_sub' in st.session_state:
                                    st.session_state.pop('out_s_sub', None)
                                o_sub_selected = st.multiselect("产品小类(可多选)", out_subs2, default=['全部'], key='out_s_sub_ms')
                            with oc5:
                                o_months = st.multiselect("时间（月）", out_month_opts, default=[10, 11, 12], key='out_s_months')

                            df_o_filtered = o_raw.copy()
                            
                            # Filter for Year 2025 (as per Q4 definition)
                            if '年份' in df_o_filtered.columns:
                                df_o_filtered = df_o_filtered[df_o_filtered['年份'] == 2025]
                                
                            if o_prov != '全部':
                                df_o_filtered = df_o_filtered[df_o_filtered['省区'].astype(str) == str(o_prov)]
                            if o_dist != '全部':
                                df_o_filtered = df_o_filtered[df_o_filtered['经销商名称'].astype(str) == str(o_dist)]
                            if o_cat != '全部':
                                df_o_filtered = df_o_filtered[df_o_filtered['产品大类'].astype(str) == str(o_cat)]
                            if o_sub_selected and ('全部' not in o_sub_selected):
                                df_o_filtered = df_o_filtered[df_o_filtered['产品小类'].astype(str).isin([str(x) for x in o_sub_selected])]

                            def _to_month(v):
                                if pd.isna(v):
                                    return None
                                if isinstance(v, (int, float)) and not pd.isna(v):
                                    m = int(v)
                                    return m if 1 <= m <= 12 else None
                                s = str(v).strip()
                                if s.isdigit():
                                    m = int(s)
                                    return m if 1 <= m <= 12 else None
                                if '月' in s:
                                    digits = ''.join([ch for ch in s if ch.isdigit()])
                                    if digits:
                                        for k in (2, 1):
                                            if len(digits) >= k:
                                                m = int(digits[-k:])
                                                if 1 <= m <= 12:
                                                    return m
                                    return None
                                dt = pd.to_datetime(s, errors='coerce')
                                if pd.isna(dt):
                                    return None
                                m = int(dt.month)
                                return m if 1 <= m <= 12 else None

                            df_o_filtered['月'] = df_o_filtered['月份'].apply(_to_month)
                            df_o_filtered = df_o_filtered[df_o_filtered['月'].notna()].copy()
                            df_o_filtered['月'] = df_o_filtered['月'].astype(int)

                            if o_months:
                                df_o_filtered = df_o_filtered[df_o_filtered['月'].isin(o_months)].copy()

                            df_o_filtered['月列'] = df_o_filtered['月'].astype(str) + '月'

                            idx_cols = ['省区', '经销商名称', '产品大类', '产品小类']
                            outbound_pivot = (
                                df_o_filtered
                                .pivot_table(index=idx_cols, columns='月列', values='数量(箱)', aggfunc='sum', fill_value=0)
                                .reset_index()
                            )

                            month_cols_full = [f"{i}月" for i in range(1, 13)]
                            for mc in month_cols_full:
                                if mc not in outbound_pivot.columns:
                                    outbound_pivot[mc] = 0

                            outbound_pivot['Q4月均销'] = (outbound_pivot['10月'] + outbound_pivot['11月'] + outbound_pivot['12月']) / 3
                            outbound_pivot = outbound_pivot[idx_cols + month_cols_full + ['Q4月均销']]

                            with st.expander("📄 出库分析底表（Sheet3）", expanded=False):
                                show_aggrid_table(outbound_pivot, height=520, key="outbound_pivot_table")

                            if not outbound_pivot.empty:
                                dist_q4 = outbound_pivot.groupby('经销商名称')[['10月', '11月', '12月']].sum().reset_index()
                                dist_q4['Q4_Total'] = dist_q4['10月'] + dist_q4['11月'] + dist_q4['12月']
                                dist_q4['Q4_Avg'] = dist_q4['Q4_Total'] / 3
                                sales_agg_q4 = dist_q4[['经销商名称', 'Q4_Total', 'Q4_Avg']].copy()

                            out_xlsx = io.BytesIO()
                            try:
                                with pd.ExcelWriter(out_xlsx, engine='openpyxl') as writer:
                                    outbound_pivot.to_excel(writer, index=False, sheet_name='Sheet3')
                            except Exception:
                                with pd.ExcelWriter(out_xlsx, engine='xlsxwriter') as writer:
                                    outbound_pivot.to_excel(writer, index=False, sheet_name='Sheet3')
                                st.download_button(
                                    "📥 下载出库分析底表 (Excel)",
                                    data=out_xlsx.getvalue(),
                                    file_name="出库分析底表_Sheet3.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                )
                    
                    # --- Drill-down State Management ---
                    # (Initialized at top of script)
                    
                    # Threshold Config
                    with st.expander("⚙️ 阈值配置", expanded=False):
                        c_th1, c_th2 = st.columns(2)
                        high_th = c_th1.number_input("库存过高阈值 (DOS >)", value=2.0, step=0.1)
                        low_th = c_th2.number_input("库存过低阈值 (DOS <)", value=0.5, step=0.1)

                    # Logic:
                    # 1. Sum Stock '箱数' by Distributor (from filtered stock df_s_filtered)
                    # 2. Match with Sheet3 Sales 'Q4_Avg' by Distributor
                    
                    # Note: df_s_filtered '经销商名称' is now '客户简称' (H column) due to load_data mapping
                    stock_agg = df_s_filtered.groupby(['省区名称', '经销商名称'])['箱数'].sum().reset_index()
                    stock_agg.rename(columns={'箱数': '当前库存_箱'}, inplace=True)
                    stock_agg['经销商名称'] = stock_agg['经销商名称'].astype(str).str.strip()
                    
                    # Merge with Q4 sales data from Sheet3
                    # LEFT JOIN ensures we only keep distributors present in the STOCK file (filtered by top filters)
                    # However, if we filter by province in top filter, df_s_filtered only has that province.
                    # sales_agg_q4 has ALL distributors from Sheet3.
                    # Merging them attaches Sales info to the Stock info.
                    analysis_df = pd.merge(stock_agg, sales_agg_q4[['经销商名称', 'Q4_Avg']], on='经销商名称', how='left')
                    analysis_df['Q4_Avg'] = analysis_df['Q4_Avg'].fillna(0)
                    
                    # 3. Calc DOS & Status
                    analysis_df['近三月未出库'] = (analysis_df['Q4_Avg'] <= 0) & (analysis_df['当前库存_箱'] > 0)

                    # Calculate DOS
                    # Optimized: Vectorized
                    q4_avg_series = analysis_df['Q4_Avg']
                    stock_series = analysis_df['当前库存_箱']
                    mask_no_outbound = analysis_df.get('近三月未出库', pd.Series(False, index=analysis_df.index)).astype(bool)
                    
                    analysis_df['可销月(DOS)'] = np.where(
                        mask_no_outbound, np.nan,
                        np.where(
                            q4_avg_series <= 0, 0.0,
                            (stock_series / q4_avg_series)
                        )
                    )
                    
                    # Ensure thresholds are defined before use
                    if 'high_th' not in locals(): high_th = 2.0
                    if 'low_th' not in locals(): low_th = 0.5

                    # Optimized: Vectorized select
                    # Pre-calculate boolean mask for '近三月未出库'
                    mask_no_outbound = analysis_df.get('近三月未出库', pd.Series(False, index=analysis_df.index)).astype(bool)
                    
                    dos_series = analysis_df['可销月(DOS)']
                    
                    conditions = [
                        mask_no_outbound,
                        pd.isna(dos_series),
                        dos_series > high_th,
                        dos_series < low_th
                    ]
                    choices = [
                        '⚫ 近三月未出库',
                        '🟢 正常',
                        '🔴 库存过高',
                        '🟠 库存不足'
                    ]
                    analysis_df['库存状态'] = np.select(conditions, choices, default='🟢 正常')

                    # --- OVERVIEW METRICS (Moved Back & Enhanced) ---
                    # Calculate metrics based on the CURRENT context (filtered data analysis_df)
                    # If drill level is 1 (All Provs), it shows total.
                    # If drill level is 2 (One Prov), we should filter analysis_df to that prov for metrics?
                    # Or should metrics always reflect the TOP filters (df_s_filtered)?
                    # User request: "When I select a specific province (in filter), real-time update."
                    # df_s_filtered IS filtered by the top dropdowns. analysis_df is derived from it.
                    # So calculating from analysis_df is correct for the top filters.
                    # However, if user clicks "Drill Down" to level 2, should the metrics update to that province?
                    # User said "When I select to specific province". 
                    # If the user uses the *Sidebar/Top Filter*, df_s_filtered updates, so analysis_df updates.
                    # If the user uses *Drill Down*, st.session_state.selected_prov is set.
                    # Usually Overview Metrics reflect the *Global Context* of the current view.
                    # Let's support both: If Drill Level > 1, filter metrics to selected scope.
                    
                    metrics_df = analysis_df.copy()
                    if st.session_state.drill_level == 2 and st.session_state.selected_prov:
                        metrics_df = metrics_df[metrics_df['省区名称'] == st.session_state.selected_prov]
                    elif st.session_state.drill_level == 3 and st.session_state.selected_dist:
                        # For level 3, it's single distributor
                         metrics_df = metrics_df[metrics_df['经销商名称'] == st.session_state.selected_dist]

                    # Calc Metrics
                    total_stock_show = metrics_df['当前库存_箱'].sum()
                    if sales_agg_q4 is not None and not sales_agg_q4.empty and 'Q4_Total' in sales_agg_q4.columns:
                        dist_scope = (
                            metrics_df['经销商名称']
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .unique()
                            .tolist()
                        )
                        sales_scope = sales_agg_q4[sales_agg_q4['经销商名称'].isin(dist_scope)].copy()
                        total_q4_avg_show = float(sales_scope['Q4_Total'].sum()) / 3 if not sales_scope.empty else 0.0
                    else:
                        total_q4_avg_show = 0.0
                    
                    # DOS = Total Stock / Total Sales
                    if total_q4_avg_show > 0:
                        dos_show = total_stock_show / total_q4_avg_show
                    else:
                        dos_show = 0.0
                    
                    if metrics_df is None or metrics_df.empty or '库存状态' not in metrics_df.columns:
                        abnormal_count_show = 0
                    else:
                        abnormal_count_show = int(
                            metrics_df['库存状态']
                            .fillna('')
                            .astype(str)
                            .str.contains('🔴|🟠|⚫', na=False)
                            .sum()
                        )
                    
                    st.markdown("### 📊 关键指标概览")
                    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                    col_m1.metric("📦 总库存 (箱)", fmt_num(total_stock_show))
                    col_m2.metric("📉 Q4月均销", fmt_num(total_q4_avg_show))
                    col_m3.metric("📅 整体可销月", fmt_num(dos_show))
                    col_m4.metric("🚨 异常客户数", f"{abnormal_count_show} 家")
                    st.markdown("---")

                    rank_stock = (
                        metrics_df.groupby('经销商名称', as_index=False)['当前库存_箱']
                        .sum()
                        .rename(columns={'当前库存_箱': '库存数(箱)'})
                    )
                    rank_stock['经销商名称'] = rank_stock['经销商名称'].astype(str).str.strip()
                    rank_stock = pd.merge(
                        rank_stock,
                        sales_agg_q4[['经销商名称', 'Q4_Avg']] if (sales_agg_q4 is not None and 'Q4_Avg' in sales_agg_q4.columns) else pd.DataFrame(columns=['经销商名称', 'Q4_Avg']),
                        on='经销商名称',
                        how='left'
                    )
                    rank_stock['Q4_Avg'] = pd.to_numeric(rank_stock.get('Q4_Avg', 0), errors='coerce').fillna(0)
                    rank_stock['近三月未出库'] = (rank_stock['Q4_Avg'] <= 0) & (rank_stock['库存数(箱)'] > 0)

                    def _rank_dos(row):
                        q4 = float(row.get('Q4_Avg', 0) or 0)
                        stk = float(row.get('库存数(箱)', 0) or 0)
                        if q4 <= 0:
                            return float('nan') if stk > 0 else 0.0
                        return stk / q4

                    rank_stock['可销月'] = rank_stock.apply(_rank_dos, axis=1)
                    rank_stock['过高差值'] = (rank_stock['可销月'] - float(high_th))
                    rank_stock['过低差值'] = (float(low_th) - rank_stock['可销月'])

                    rank_stock_rankable = rank_stock[~rank_stock['近三月未出库']].copy()
                    high_top = rank_stock_rankable[rank_stock_rankable['过高差值'] > 0].copy().sort_values('过高差值', ascending=False).head(10)
                    low_top = rank_stock_rankable[rank_stock_rankable['过低差值'] > 0].copy().sort_values('过低差值', ascending=False).head(10)

                    st.markdown("### 🏆 异常库存TOP10经销商")
                    r1, r2 = st.columns(2)
                    with r1:
                        st.subheader("🔴 库存过高 TOP10")
                        if high_top.empty:
                            st.info("当前范围无库存过高经销商")
                        else:
                            high_chart = high_top.sort_values('过高差值', ascending=True).copy()
                            high_chart['标注'] = high_chart['过高差值'].map(lambda x: f"+{fmt_num(x, na='')}")
                            high_chart['_库存数_fmt'] = high_chart['库存数(箱)'].map(lambda x: fmt_num(x, na=''))
                            high_chart['_q4_fmt'] = high_chart['Q4_Avg'].map(lambda x: fmt_num(x, na=''))
                            high_chart['_dos_fmt'] = high_chart['可销月'].map(lambda x: fmt_num(x, na=''))
                            high_chart['_diff_fmt'] = high_chart['过高差值'].map(lambda x: fmt_num(x, na=''))
                            fig_high = px.bar(
                                high_chart,
                                x='过高差值',
                                y='经销商名称',
                                orientation='h',
                                text='标注',
                                title="超出过高阈值的差值（可销月 - 阈值）",
                                color_discrete_sequence=['#E5484D'],
                                custom_data=['_库存数_fmt', '_q4_fmt', '_dos_fmt', '_diff_fmt']
                            )
                            fig_high.update_traces(
                                textposition='outside',
                                hovertemplate=(
                                    "经销商: %{y}<br>"
                                    "库存数(箱): %{customdata[0]}<br>"
                                    "Q4月均销: %{customdata[1]}<br>"
                                    "可销月: %{customdata[2]}<br>"
                                    "超阈值差值: +%{customdata[3]}<extra></extra>"
                                )
                            )
                            fig_high.update_layout(height=420, xaxis_title="差值", yaxis_title="")
                            st.plotly_chart(fig_high, use_container_width=True)
                            show_aggrid_table(high_top[['经销商名称', '库存数(箱)', 'Q4_Avg', '可销月', '过高差值']], height=250, key='high_stock_ag')

                    with r2:
                        st.subheader("🟠 库存过低 TOP10")
                        if low_top.empty:
                            st.info("当前范围无库存过低经销商")
                        else:
                            low_chart = low_top.sort_values('过低差值', ascending=True).copy()
                            low_chart['标注'] = low_chart['过低差值'].map(lambda x: f"+{fmt_num(x, na='')}")
                            low_chart['_库存数_fmt'] = low_chart['库存数(箱)'].map(lambda x: fmt_num(x, na=''))
                            low_chart['_q4_fmt'] = low_chart['Q4_Avg'].map(lambda x: fmt_num(x, na=''))
                            low_chart['_dos_fmt'] = low_chart['可销月'].map(lambda x: fmt_num(x, na=''))
                            low_chart['_diff_fmt'] = low_chart['过低差值'].map(lambda x: fmt_num(x, na=''))
                            fig_low = px.bar(
                                low_chart,
                                x='过低差值',
                                y='经销商名称',
                                orientation='h',
                                text='标注',
                                title="低于过低阈值的差值（阈值 - 可销月）",
                                color_discrete_sequence=['#FFB000'],
                                custom_data=['_库存数_fmt', '_q4_fmt', '_dos_fmt', '_diff_fmt']
                            )
                            fig_low.update_traces(
                                textposition='outside',
                                hovertemplate=(
                                    "经销商: %{y}<br>"
                                    "库存数(箱): %{customdata[0]}<br>"
                                    "Q4月均销: %{customdata[1]}<br>"
                                    "可销月: %{customdata[2]}<br>"
                                    "低于阈值差值: +%{customdata[3]}<extra></extra>"
                                )
                            )
                            fig_low.update_layout(height=420, xaxis_title="差值", yaxis_title="")
                            st.plotly_chart(fig_low, use_container_width=True)
                            show_aggrid_table(low_top[['经销商名称', '库存数(箱)', 'Q4_Avg', '可销月', '过低差值']], height=250, key='low_stock_ag')

                    with st.expander("🔍 对账信息", expanded=False):
                        if df_o_filtered is None or df_o_filtered.empty or '月' not in df_o_filtered.columns:
                            st.warning("当前筛选下无出库明细可对账。")
                        else:
                            s10 = float(df_o_filtered[df_o_filtered['月'] == 10]['数量(箱)'].sum()) if '数量(箱)' in df_o_filtered.columns else 0.0
                            s11 = float(df_o_filtered[df_o_filtered['月'] == 11]['数量(箱)'].sum()) if '数量(箱)' in df_o_filtered.columns else 0.0
                            s12 = float(df_o_filtered[df_o_filtered['月'] == 12]['数量(箱)'].sum()) if '数量(箱)' in df_o_filtered.columns else 0.0
                            st.write(f"当前筛选下Sheet3合计：10月={fmt_num(s10)}，11月={fmt_num(s11)}，12月={fmt_num(s12)}")
                            st.write(f"当前筛选下Q4月均销=(10+11+12)/3 = {fmt_num((s10+s11+s12)/3)}")
                            if sales_agg_q4 is not None and 'Q4_Total' in sales_agg_q4.columns:
                                dist_scope_dbg = (
                                    metrics_df['经销商名称']
                                    .dropna()
                                    .astype(str)
                                    .str.strip()
                                    .unique()
                                    .tolist()
                                )
                                matched = sales_agg_q4[sales_agg_q4['经销商名称'].isin(dist_scope_dbg)]
                                st.write(f"当前范围经销商数(去重)：{len(dist_scope_dbg)}，Sheet3匹配到：{len(matched)}")
                                st.write(f"当前范围Q4月均销=(sum(Q4_Total))/3 = {fmt_num(float(matched['Q4_Total'].sum())/3)}")

                    # --- Navigation & Breadcrumbs ---
                    cols_nav = st.columns([1, 8])
                    if st.session_state.drill_level > 1:
                        if cols_nav[0].button("⬅️ 返回"):
                            st.session_state.drill_level -= 1
                            st.rerun()
                    
                    breadcrumbs = "🏠 全部省区"
                    if st.session_state.drill_level >= 2:
                        breadcrumbs += f" > 📍 {st.session_state.selected_prov}"
                    if st.session_state.drill_level >= 3:
                        breadcrumbs += f" > 🏢 {st.session_state.selected_dist}"
                    cols_nav[1].markdown(f"**当前位置**: {breadcrumbs}")

                    # --- Level 1: Province View ---
                    if st.session_state.drill_level == 1:
                        
                        # Agg by Prov
                        prov_agg = analysis_df.groupby('省区名称').agg({
                            '当前库存_箱': 'sum',
                            'Q4_Avg': 'sum',
                            '经销商名称': 'count' # Count of distributors
                        }).reset_index()
                        
                        # Calc Prov DOS
                        prov_agg['可销月(DOS)'] = prov_agg.apply(lambda x: (x['当前库存_箱'] / x['Q4_Avg']) if x['Q4_Avg'] > 0 else (float('nan') if x['当前库存_箱'] > 0 else 0.0), axis=1)
                        
                        # Count Abnormal Distributors per Prov
                        abnormal_counts = analysis_df.groupby('省区名称')['库存状态'].value_counts().unstack(fill_value=0)
                        if '🔴 库存过高' not in abnormal_counts.columns: abnormal_counts['🔴 库存过高'] = 0
                        if '🟠 库存不足' not in abnormal_counts.columns: abnormal_counts['🟠 库存不足'] = 0
                        if '⚫ 近三月未出库' not in abnormal_counts.columns: abnormal_counts['⚫ 近三月未出库'] = 0
                        
                        prov_view = pd.merge(prov_agg, abnormal_counts[['🔴 库存过高', '🟠 库存不足', '⚫ 近三月未出库']], on='省区名称', how='left').fillna(0)
                        
                        # New Logic: Calculate Total Abnormal Count and Sort
                        prov_view['合计异常数'] = prov_view['🔴 库存过高'] + prov_view['🟠 库存不足'] + prov_view['⚫ 近三月未出库']
                        prov_view['经销商总数'] = prov_view['经销商名称'] # Rename for clarity
                        
                        # Filter slider
                        max_abnormal = int(prov_view['合计异常数'].max()) if not prov_view.empty else 10
                        c_filter, _ = st.columns([1, 2])
                        min_abnormal_filter = c_filter.slider("🔎 异常数过滤 (≥)", 0, max_abnormal, 0)
                        
                        prov_view_filtered = prov_view[prov_view['合计异常数'] >= min_abnormal_filter].copy()
                        
                        # Sort Descending by Total Abnormal Count
                        prov_view_filtered = prov_view_filtered.sort_values('合计异常数', ascending=False)
                        
                        st.markdown("### 📋 省区库存异常详情列表")
                        st.caption("💡 提示：**直接点击表格中的某一行**，即可下钻查看该省区的经销商详情。")
                        
                        # Prepare DF for display
                        display_df = prov_view_filtered[["省区名称", "合计异常数", "🔴 库存过高", "🟠 库存不足", "⚫ 近三月未出库", "当前库存_箱", "Q4_Avg", "可销月(DOS)"]].reset_index(drop=True)
                        
                        # Use interactive dataframe with selection
                        # Dynamic height to show all rows
                        n_rows = len(display_df)
                        # Estimate height: 35px per row + 35px header + buffer
                        calc_height = (n_rows + 1) * 35 + 10
                        # Ensure a minimum height and reasonable max height (e.g., 2000px)
                        final_height = max(150, min(calc_height, 2000))

                        ag_inv = show_aggrid_table(
                            display_df,
                            height=final_height,
                            columns_props={'合计异常数': {'type': 'bar_count'}, '可销月(DOS)': {'type': 'number'}},
                            on_row_selected='single',
                            key='inv_prov_ag'
                        )
                        
                        # Show all province names as tags below for quick view
                        with st.expander("查看所有省区名称列表 (点击展开)", expanded=False):
                            st.markdown("  ".join([f"`{p}`" for p in display_df['省区名称'].tolist()]))
                        
                        # Handle Selection
                        selected_rows = ag_inv.get('selected_rows') if ag_inv else None
                        if selected_rows is not None and len(selected_rows) > 0:
                            if isinstance(selected_rows, pd.DataFrame):
                                first_row = selected_rows.iloc[0]
                            else:
                                first_row = selected_rows[0]
                            
                            selected_prov_name = first_row.get("省区名称") if isinstance(first_row, dict) else first_row["省区名称"]
                            st.session_state.selected_prov = selected_prov_name
                            st.session_state.drill_level = 2
                            st.rerun()

                        # Visualization: Stacked Bar Chart of Abnormalities
                        if not prov_view_filtered.empty:
                            fig_abnormal = px.bar(
                                prov_view_filtered,
                                x='省区名称',
                                y=['🔴 库存过高', '🟠 库存不足'],
                                title='各省异常库存分布',
                                labels={'value': '经销商数量', 'variable': '异常类型'},
                                color_discrete_map={'🔴 库存过高': '#E5484D', '🟠 库存不足': '#FFB000'}
                            )
                            fig_abnormal.update_layout(height=350, margin=dict(l=20, r=20, t=40, b=20))
                            st.plotly_chart(fig_abnormal, use_container_width=True)

                    # --- Level 2: Distributor View ---
                    elif st.session_state.drill_level == 2:
                        prov = st.session_state.selected_prov
                        st.caption("💡 提示：**点击表格行** 可查看该经销商的 SKU 库存明细。")
                        
                        # Filter by Prov
                        dist_view = analysis_df[analysis_df['省区名称'] == prov].copy().reset_index(drop=True)
                        
                        # Interactive Table
                        ag_dist_inv = show_aggrid_table(
                            dist_view[['经销商名称', '当前库存_箱', 'Q4_Avg', '可销月(DOS)', '库存状态']],
                            height=520,
                            columns_props={'可销月(DOS)': {'type': 'number'}},
                            on_row_selected='single',
                            key='inv_dist_ag'
                        )
                        
                        # Handle Selection
                        selected_rows_d = ag_dist_inv.get('selected_rows') if ag_dist_inv else None
                        if selected_rows_d is not None and len(selected_rows_d) > 0:
                            if isinstance(selected_rows_d, pd.DataFrame):
                                first_row_d = selected_rows_d.iloc[0]
                            else:
                                first_row_d = selected_rows_d[0]
                            
                            selected_dist_name = first_row_d.get("经销商名称") if isinstance(first_row_d, dict) else first_row_d["经销商名称"]
                            st.session_state.selected_dist = selected_dist_name
                            st.session_state.drill_level = 3
                            st.rerun()

                    # --- Level 3: SKU/Store View ---
                    elif st.session_state.drill_level == 3:
                        dist = st.session_state.selected_dist
                        
                        # Get SKU details for this distributor from filtered stock data
                        # Note: We don't have store-level sales in Sheet3 (only Dist level), 
                        # so we can only show Stock Details here, potentially calculating SKU-level DOS if we had SKU-level sales (which we don't from Sheet3).
                        # We will show SKU stock details.
                        
                        sku_view = df_s_filtered[df_s_filtered['经销商名称'] == dist][['产品名称', '产品编码', '箱数', '规格', '重量']].copy()
                        
                        show_aggrid_table(sku_view, height=520, key='inv_sku_ag')
                        st.caption("注：因Q4出库数据仅精确到经销商层级，此处仅展示SKU库存明细，不计算单品DOS。")

            with tab_out:
                if df_q4_raw is None or df_q4_raw.empty:
                    st.warning("⚠️ 未检测到出库数据 (Sheet3)。请确认Excel包含Sheet3且数据完整。")
                    with st.expander("🛠️ 调试信息", expanded=False):
                        for log in debug_logs:
                            st.text(log)
                else:
                    st.caption(f"🕒 数据更新时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

                    o_raw = df_q4_raw.copy()

                    if '产品大类' not in o_raw.columns:
                        o_raw['产品大类'] = '全部'
                    if '产品小类' not in o_raw.columns:
                        o_raw['产品小类'] = '全部'

                    day_col = next((c for c in o_raw.columns if str(c).strip() == '日'), None)
                    if day_col is None:
                        day_col = next((c for c in o_raw.columns if ('日期' in str(c)) or (str(c).strip().endswith('日') and '月' not in str(c))), None)
                    if day_col is None and len(o_raw.columns) > 14:
                        day_col = o_raw.columns[14]

                    store_name_col = o_raw.columns[5] if len(o_raw.columns) > 5 else None

                    if '数量(箱)' in o_raw.columns:
                        o_raw['数量(箱)'] = pd.to_numeric(o_raw['数量(箱)'], errors='coerce').fillna(0.0)
                    else:
                        o_raw['数量(箱)'] = 0.0

                    if store_name_col is not None and store_name_col in o_raw.columns:
                        o_raw['_门店名'] = (
                            o_raw[store_name_col]
                            .fillna('')
                            .astype(str)
                            .str.replace(r'\s+', '', regex=True)
                        )
                        o_raw.loc[o_raw['_门店名'].isin(['', 'nan', 'None', 'NULL', 'NaN']), '_门店名'] = pd.NA
                    else:
                        o_raw['_门店名'] = pd.NA

                    def _to_month(v):
                        if pd.isna(v):
                            return None
                        if isinstance(v, (int, float)) and not pd.isna(v):
                            m = int(v)
                            return m if 1 <= m <= 12 else None
                        s = str(v).strip()
                        if s.isdigit():
                            m = int(s)
                            return m if 1 <= m <= 12 else None
                        if '月' in s:
                            digits = ''.join([ch for ch in s if ch.isdigit()])
                            if digits:
                                for k in (2, 1):
                                    if len(digits) >= k:
                                        m = int(digits[-k:])
                                        if 1 <= m <= 12:
                                            return m
                            return None
                        dt = pd.to_datetime(s, errors='coerce')
                        if pd.isna(dt):
                            return None
                        m = int(dt.month)
                        return m if 1 <= m <= 12 else None

                    def _to_day(v):
                        if pd.isna(v):
                            return None
                        if isinstance(v, (int, float)) and not pd.isna(v):
                            d = int(v)
                            return d if 1 <= d <= 31 else None
                        s = str(v).strip()
                        digits = ''.join([ch for ch in s if ch.isdigit()])
                        if not digits:
                            return None
                        d = int(digits[-2:]) if len(digits) >= 2 else int(digits)
                        return d if 1 <= d <= 31 else None

                    if '年份' in o_raw.columns:
                        o_raw['_年'] = pd.to_numeric(o_raw['年份'], errors='coerce').fillna(0).astype(int)
                    else:
                        o_raw['_年'] = 0
                    if '月份' in o_raw.columns:
                        o_raw['_月'] = o_raw['月份'].apply(_to_month)
                    else:
                        o_raw['_月'] = None

                    if day_col is not None and day_col in o_raw.columns:
                        if '日期' in str(day_col):
                            dt_series = pd.to_datetime(o_raw[day_col], errors='coerce')
                            o_raw['_年'] = np.where(dt_series.notna(), dt_series.dt.year, o_raw['_年']).astype(int)
                            o_raw['_月'] = np.where(dt_series.notna(), dt_series.dt.month, o_raw['_月'])
                            o_raw['_日'] = np.where(dt_series.notna(), dt_series.dt.day, None)
                        else:
                            o_raw['_日'] = o_raw[day_col].apply(_to_day)
                    else:
                        o_raw['_日'] = None

                    o_raw = o_raw[o_raw['_年'] > 0].copy()
                    o_raw = o_raw[o_raw['_月'].notna()].copy()
                    o_raw['_月'] = o_raw['_月'].astype(int)
                    o_raw['_日'] = pd.to_numeric(o_raw['_日'], errors='coerce')

                    with st.expander("🛠️ 出库筛选", expanded=False):
                        out_provs = ['全部'] + sorted(o_raw['省区'].dropna().astype(str).unique().tolist()) if '省区' in o_raw.columns else ['全部']
                        oc1, oc2, oc3, oc4, oc5 = st.columns(5)
                        with oc1:
                            o_prov = st.selectbox("省区", out_provs, key='out2_prov')
                        with oc2:
                            if '经销商名称' in o_raw.columns:
                                if o_prov != '全部' and '省区' in o_raw.columns:
                                    dists_in_prov = o_raw[o_raw['省区'].astype(str) == str(o_prov)]['经销商名称'].dropna().astype(str).unique().tolist()
                                    out_dists = ['全部'] + sorted(dists_in_prov)
                                else:
                                    out_dists = ['全部'] + sorted(o_raw['经销商名称'].dropna().astype(str).unique().tolist())
                            else:
                                out_dists = ['全部']
                            o_dist = st.selectbox("经销商", out_dists, key='out2_dist')
                        with oc3:
                            out_cats = ['全部'] + sorted(o_raw['产品大类'].dropna().astype(str).unique().tolist())
                            o_cat = st.selectbox("产品大类", out_cats, key='out2_cat')
                        with oc4:
                            if o_cat != '全部':
                                subs_in_cat = o_raw[o_raw['产品大类'].astype(str) == str(o_cat)]['产品小类'].dropna().astype(str).unique().tolist()
                                out_subs = ['全部'] + sorted(subs_in_cat)
                            else:
                                out_subs = ['全部'] + sorted(o_raw['产品小类'].dropna().astype(str).unique().tolist())
                            o_sub = st.selectbox("产品小类", out_subs, key='out2_sub')
                        with oc5:
                            year_opts = sorted([int(y) for y in o_raw['_年'].dropna().unique().tolist() if int(y) > 0])
                            default_year = 2025 if 2025 in year_opts else (max(year_opts) if year_opts else 2025)
                            y_index = year_opts.index(default_year) if default_year in year_opts else 0
                            o_year = st.selectbox("年份", year_opts if year_opts else [2025], index=y_index, key='out2_year')
                            month_in_year = sorted([int(m) for m in o_raw[o_raw['_年'] == int(o_year)]['_月'].dropna().unique().tolist() if 1 <= int(m) <= 12])
                            month_opts = ['全部'] + month_in_year
                            o_month = st.selectbox("月份", month_opts, index=0, key='out2_month')

                    df_o = o_raw.copy()
                    if o_prov != '全部' and '省区' in df_o.columns:
                        df_o = df_o[df_o['省区'].astype(str) == str(o_prov)]
                    if o_dist != '全部' and '经销商名称' in df_o.columns:
                        df_o = df_o[df_o['经销商名称'].astype(str) == str(o_dist)]
                    if o_cat != '全部':
                        df_o = df_o[df_o['产品大类'].astype(str) == str(o_cat)]
                    if o_sub != '全部':
                        df_o = df_o[df_o['产品小类'].astype(str) == str(o_sub)]

                    def _agg_scope(df_scope: pd.DataFrame):
                        boxes = float(df_scope.get('数量(箱)', 0).sum()) if df_scope is not None and not df_scope.empty else 0.0
                        if df_scope is None or df_scope.empty or '_门店名' not in df_scope.columns:
                            stores = 0.0
                        else:
                            df_s = df_scope[df_scope['数量(箱)'] > 0].copy()
                            stores = float(df_s['_门店名'].dropna().astype(str).nunique()) if not df_s.empty else 0.0
                        return boxes, stores

                    def _yoy(cur, last):
                        if last is None:
                            return None
                        last_v = float(last or 0)
                        if last_v <= 0:
                            return None
                        return (float(cur or 0) - last_v) / last_v

                    def _avg(boxes, stores):
                        try:
                            s = float(stores or 0)
                            return float(boxes or 0) / s if s > 0 else 0.0
                        except Exception:
                            return 0.0

                    def _fmt_num(x):
                        return fmt_num(x, na="0")

                    def _fmt_pct(x):
                        return fmt_pct_ratio(x) if x is not None else "—"

                    def _trend_cls(x):
                        if x is None or (isinstance(x, float) and pd.isna(x)):
                            return "trend-neutral"
                        return "trend-up" if x > 0 else ("trend-down" if x < 0 else "trend-neutral")

                    def _arrow(x):
                        if x is None or (isinstance(x, float) and pd.isna(x)):
                            return ""
                        return "↑" if x > 0 else ("↓" if x < 0 else "")

                    # === Use Native Tabs for Consistency with Other Modules ===
                    tab_kpi, tab_cat, tab_prov = st.tabs(["📊 关键指标", "📦 分品类", "🗺️ 分省区"])
                    
                    # Prepare Data Context (Shared)
                    sig = (o_prov, o_dist, o_cat, o_sub, o_year, o_month)
                    if "out_subtab_cache" not in st.session_state:
                        st.session_state.out_subtab_cache = {}
                    
                    def _get_ctx():
                        ck = ("ctx", sig)
                        if ck in st.session_state.out_subtab_cache:
                            return st.session_state.out_subtab_cache[ck]
                        
                        # No spinner here to avoid flashing on every rerun, 
                        # relying on Streamlit's natural execution speed or cache if possible.
                        # If slow, we might add st.spinner inside specific heavy blocks.
                        if o_month != '全部':
                            _kpi_year = int(o_year)
                            _kpi_month = int(o_month)
                        else:
                            years_avail = sorted([int(y) for y in df_o['_年'].dropna().unique().tolist() if int(y) > 0])
                            _kpi_year = 2026 if 2026 in years_avail else (max(years_avail) if years_avail else int(o_year))
                            months_avail = sorted([int(m) for m in df_o[df_o['_年'] == int(_kpi_year)]['_月'].dropna().unique().tolist() if 1 <= int(m) <= 12])
                            _kpi_month = max(months_avail) if months_avail else 1

                        days_avail = sorted([int(d) for d in df_o[(df_o['_年'] == int(_kpi_year)) & (df_o['_月'] == int(_kpi_month))]['_日'].dropna().unique().tolist() if 1 <= int(d) <= 31])
                        _kpi_day = max(days_avail) if days_avail else None
                        _cmp_year = int(_kpi_year) - 1

                        _cur_today = (df_o[(df_o['_年'] == int(_kpi_year)) & (df_o['_月'] == int(_kpi_month)) & (df_o['_日'] == int(_kpi_day))] if _kpi_day is not None else df_o.iloc[0:0])
                        _cur_month = df_o[(df_o['_年'] == int(_kpi_year)) & (df_o['_月'] == int(_kpi_month))]
                        _cur_year = df_o[(df_o['_年'] == int(_kpi_year))]

                        _last_today = (df_o[(df_o['_年'] == int(_cmp_year)) & (df_o['_月'] == int(_kpi_month)) & (df_o['_日'] == int(_kpi_day))] if _kpi_day is not None else df_o.iloc[0:0])
                        _last_month = df_o[(df_o['_年'] == int(_cmp_year)) & (df_o['_月'] == int(_kpi_month))]
                        _last_year = df_o[(df_o['_年'] == int(_cmp_year))]

                        ctx = {
                            "kpi_year": _kpi_year,
                            "kpi_month": _kpi_month,
                            "kpi_day": _kpi_day,
                            "cmp_year": _cmp_year,
                            "cur_today": _cur_today,
                            "cur_month": _cur_month,
                            "cur_year": _cur_year,
                            "last_today": _last_today,
                            "last_month": _last_month,
                            "last_year": _last_year,
                        }
                        st.session_state.out_subtab_cache[ck] = ctx
                        return ctx

                    ctx = _get_ctx()
                    
                    # Common Caption
                    st.caption(
                        f"当前默认口径：{ctx['kpi_year']}年{int(ctx['kpi_month'])}月"
                        + (f"{int(ctx['kpi_day'])}日" if ctx["kpi_day"] is not None else "")
                    )

                    # --- Tab 1: KPI ---
                    with tab_kpi:
                        ck = ("kpi", sig)
                        if ck not in st.session_state.out_subtab_cache:
                             t_boxes, t_stores = _agg_scope(ctx["cur_today"])
                             tm_boxes, tm_stores = _agg_scope(ctx["cur_month"])
                             ty_boxes, ty_stores = _agg_scope(ctx["cur_year"])
                             lt_boxes, lt_stores = _agg_scope(ctx["last_today"])
                             ltm_boxes, ltm_stores = _agg_scope(ctx["last_month"])
                             lty_boxes, lty_stores = _agg_scope(ctx["last_year"])
                             t_yoy = _yoy(t_boxes, lt_boxes)
                             tm_yoy = _yoy(tm_boxes, ltm_boxes)
                             ty_yoy = _yoy(ty_boxes, lty_boxes)
                             t_avg = _avg(t_boxes, t_stores)
                             tm_avg = _avg(tm_boxes, tm_stores)
                             ty_avg = _avg(ty_boxes, ty_stores)
                             lt_avg = _avg(lt_boxes, lt_stores)
                             ltm_avg = _avg(ltm_boxes, ltm_stores)
                             lty_avg = _avg(lty_boxes, lty_stores)
                             st.session_state.out_subtab_cache[ck] = {
                                "t_boxes": t_boxes, "t_stores": t_stores, "t_yoy": t_yoy, "t_avg": t_avg, "lt_boxes": lt_boxes, "lt_stores": lt_stores, "lt_avg": lt_avg,
                                "tm_boxes": tm_boxes, "tm_stores": tm_stores, "tm_yoy": tm_yoy, "tm_avg": tm_avg, "ltm_boxes": ltm_boxes, "ltm_stores": ltm_stores, "ltm_avg": ltm_avg,
                                "ty_boxes": ty_boxes, "ty_stores": ty_stores, "ty_yoy": ty_yoy, "ty_avg": ty_avg, "lty_boxes": lty_boxes, "lty_stores": lty_stores, "lty_avg": lty_avg,
                             }
                        m = st.session_state.out_subtab_cache[ck]

                        k1, k2, k3 = st.columns(3)
                        with k1:
                            st.markdown(f"""
                            <div class="out-kpi-card">
                                <div class="out-kpi-bar"></div>
                                <div class="out-kpi-head">
                                    <div class="out-kpi-ico">🚚</div>
                                    <div class="out-kpi-title">本日出库</div>
                                </div>
                                <div class="out-kpi-val">{_fmt_num(m['t_boxes'])} 箱</div>
                                <div class="out-kpi-sub"><span>门店数</span><span>{_fmt_num(m['t_stores'])}</span></div>
                                <div class="out-kpi-sub2"><span>店均（箱/店）</span><span>{fmt_num(m['t_avg'])} <span style="color:rgba(27,21,48,0.55);">｜同期 {fmt_num(m['lt_avg'])}</span></span></div>
                                <div class="out-kpi-sub2" style="margin-top:10px;"><span>同期({ctx['cmp_year']})</span><span>{_fmt_num(m['lt_boxes'])} 箱 / {_fmt_num(m['lt_stores'])} 店</span></div>
                                <div class="out-kpi-sub2"><span>同比（箱）</span><span class="{_trend_cls(m['t_yoy'])}">{_arrow(m['t_yoy'])} {_fmt_pct(m['t_yoy'])}</span></div>
                            </div>
                            """, unsafe_allow_html=True)

                        with k2:
                            st.markdown(f"""
                            <div class="out-kpi-card">
                                <div class="out-kpi-bar"></div>
                                <div class="out-kpi-head">
                                    <div class="out-kpi-ico">📦</div>
                                    <div class="out-kpi-title">本月累计出库</div>
                                </div>
                                <div class="out-kpi-val">{_fmt_num(m['tm_boxes'])} 箱</div>
                                <div class="out-kpi-sub"><span>门店数</span><span>{_fmt_num(m['tm_stores'])}</span></div>
                                <div class="out-kpi-sub2"><span>店均（箱/店）</span><span>{fmt_num(m['tm_avg'])} <span style="color:rgba(27,21,48,0.55);">｜同期 {fmt_num(m['ltm_avg'])}</span></span></div>
                                <div class="out-kpi-sub2" style="margin-top:10px;"><span>同期({ctx['cmp_year']})</span><span>{_fmt_num(m['ltm_boxes'])} 箱 / {_fmt_num(m['ltm_stores'])} 店</span></div>
                                <div class="out-kpi-sub2"><span>同比（箱）</span><span class="{_trend_cls(m['tm_yoy'])}">{_arrow(m['tm_yoy'])} {_fmt_pct(m['tm_yoy'])}</span></div>
                            </div>
                            """, unsafe_allow_html=True)

                        with k3:
                            st.markdown(f"""
                            <div class="out-kpi-card">
                                <div class="out-kpi-bar"></div>
                                <div class="out-kpi-head">
                                    <div class="out-kpi-ico">🏁</div>
                                    <div class="out-kpi-title">本年累计出库</div>
                                </div>
                                <div class="out-kpi-val">{_fmt_num(m['ty_boxes'])} 箱</div>
                                <div class="out-kpi-sub"><span>门店数</span><span>{_fmt_num(m['ty_stores'])}</span></div>
                                <div class="out-kpi-sub2"><span>店均（箱/店）</span><span>{fmt_num(m['ty_avg'])} <span style="color:rgba(27,21,48,0.55);">｜同期 {fmt_num(m['lty_avg'])}</span></span></div>
                                <div class="out-kpi-sub2" style="margin-top:10px;"><span>同期({ctx['cmp_year']})</span><span>{_fmt_num(m['lty_boxes'])} 箱 / {_fmt_num(m['lty_stores'])} 店</span></div>
                                <div class="out-kpi-sub2"><span>同比（箱）</span><span class="{_trend_cls(m['ty_yoy'])}">{_arrow(m['ty_yoy'])} {_fmt_pct(m['ty_yoy'])}</span></div>
                            </div>
                            """, unsafe_allow_html=True)

                    # --- Tab 2: Category ---
                    with tab_cat:
                        ck = ("cat", sig)
                        if ck not in st.session_state.out_subtab_cache:
                            with st.spinner("正在加载分品类…"):
                                cat_dim = '产品小类' if o_cat != '全部' else '产品大类'
                                st.session_state.out_subtab_cache[ck] = {"cat_dim": cat_dim}
                        cat_dim = st.session_state.out_subtab_cache[ck]["cat_dim"]
                        dim_label = '产品小类' if cat_dim == '产品小类' else '产品大类'

                        st.caption(f"统计维度：{dim_label}（随筛选条件实时更新）")

                        def _cat_agg(df_scope: pd.DataFrame):
                            if df_scope is None or df_scope.empty:
                                return pd.DataFrame(columns=[cat_dim, '箱数', '门店数'])
                            df_t = df_scope.copy()
                            if cat_dim not in df_t.columns:
                                df_t[cat_dim] = '未知'
                            df_t[cat_dim] = df_t[cat_dim].fillna('未知').astype(str).str.strip()
                            df_t = df_t[df_t['数量(箱)'] > 0].copy()
                            if df_t.empty:
                                return pd.DataFrame(columns=[cat_dim, '箱数', '门店数'])
                            g_box = df_t.groupby(cat_dim, as_index=False)['数量(箱)'].sum().rename(columns={'数量(箱)': '箱数'})
                            if '_门店名' in df_t.columns:
                                g_store = df_t[df_t['_门店名'].notna()].groupby(cat_dim, as_index=False)['_门店名'].nunique().rename(columns={'_门店名': '门店数'})
                            else:
                                g_store = pd.DataFrame({cat_dim: g_box[cat_dim], '门店数': 0})
                            out = pd.merge(g_box, g_store, on=cat_dim, how='left').fillna(0)
                            out = out.sort_values('箱数', ascending=False).reset_index(drop=True)
                            return out

                        def _topn_with_other(df_sum: pd.DataFrame, n: int = 15):
                            if df_sum is None or df_sum.empty:
                                return df_sum
                            head = df_sum.head(n).copy()
                            tail = df_sum.iloc[n:].copy()
                            if not tail.empty:
                                other = pd.DataFrame([{
                                    cat_dim: '其他',
                                    '箱数': float(tail['箱数'].sum()),
                                    '门店数': float(tail['门店数'].sum())
                                }])
                                head = pd.concat([head, other], ignore_index=True)
                            return head

                        def _cat_table(df_cur: pd.DataFrame, df_last: pd.DataFrame):
                            cur_sum = _topn_with_other(_cat_agg(df_cur), 15)
                            last_sum = _topn_with_other(_cat_agg(df_last), 15)
                            if cur_sum is None or cur_sum.empty:
                                cur_sum = pd.DataFrame(columns=[cat_dim, '箱数', '门店数'])
                            if last_sum is None or last_sum.empty:
                                last_sum = pd.DataFrame(columns=[cat_dim, '箱数', '门店数'])
                            m = pd.merge(
                                cur_sum.rename(columns={'箱数': '箱数', '门店数': '门店数'}),
                                last_sum[[cat_dim, '箱数']].rename(columns={'箱数': '同期（箱数）'}),
                                on=cat_dim,
                                how='outer'
                            ).fillna(0)
                            m['同比'] = np.where(m['同期（箱数）'] > 0, (m['箱数'] - m['同期（箱数）']) / m['同期（箱数）'], None)
                            m = m.sort_values('箱数', ascending=False).reset_index(drop=True)
                            m = m.rename(columns={cat_dim: '品类'})
                            return m[['品类', '箱数', '门店数', '同期（箱数）', '同比']]

                        tab_cat_today, tab_cat_month, tab_cat_year = st.tabs(["本日", "本月", "本年"])
                        with tab_cat_today:
                            cat_tbl = _cat_table(ctx["cur_today"], ctx["last_today"])
                            show_aggrid_table(cat_tbl, columns_props={'同比': {'type': 'percent'}}, auto_height_limit=520)
                        with tab_cat_month:
                            cat_tbl = _cat_table(ctx["cur_month"], ctx["last_month"])
                            show_aggrid_table(cat_tbl, columns_props={'同比': {'type': 'percent'}}, auto_height_limit=520)
                        with tab_cat_year:
                            cat_tbl = _cat_table(ctx["cur_year"], ctx["last_year"])
                            show_aggrid_table(cat_tbl, columns_props={'同比': {'type': 'percent'}}, auto_height_limit=520)

                    # --- Tab 3: Province ---
                    with tab_prov:

                        def _prov_agg(df_scope: pd.DataFrame):
                            if df_scope is None or df_scope.empty or '省区' not in df_scope.columns:
                                return pd.DataFrame(columns=['省区', '箱数', '门店数'])
                            g_box = (
                                df_scope
                                .groupby('省区', as_index=False)['数量(箱)']
                                .sum()
                                .rename(columns={'数量(箱)': '箱数'})
                            )

                            if '_门店名' in df_scope.columns:
                                tmp = df_scope[(df_scope['数量(箱)'] > 0) & (df_scope['_门店名'].notna())].copy()
                                g_store = (
                                    tmp
                                    .groupby('省区', as_index=False)['_门店名']
                                    .nunique()
                                    .rename(columns={'_门店名': '门店数'})
                                )
                            else:
                                g_store = pd.DataFrame(columns=['省区', '门店数'])

                            return pd.merge(g_box, g_store, on='省区', how='left').fillna(0)

                        p_cur_today = _prov_agg(ctx["cur_today"])
                        p_cur_month = _prov_agg(ctx["cur_month"])
                        p_cur_year = _prov_agg(ctx["cur_year"])
                        p_last_today = _prov_agg(ctx["last_today"])
                        p_last_month = _prov_agg(ctx["last_month"])
                        p_last_year = _prov_agg(ctx["last_year"])

                        prov_all = sorted(set(
                            p_cur_today['省区'].astype(str).tolist()
                            + p_cur_month['省区'].astype(str).tolist()
                            + p_cur_year['省区'].astype(str).tolist()
                        ))
                        prov_df = pd.DataFrame({'省区': prov_all})

                        def _merge(prov_base, df_left, prefix):
                            d = df_left.copy()
                            d.columns = ['省区'] + [f"{prefix}{c}" for c in d.columns if c != '省区']
                            return pd.merge(prov_base, d, on='省区', how='left').fillna(0)

                        prov_df = _merge(prov_df, p_cur_today, "今日")
                        prov_df = _merge(prov_df, p_last_today, "同期今日")
                        prov_df = _merge(prov_df, p_cur_month, "本月")
                        prov_df = _merge(prov_df, p_last_month, "同期本月")
                        prov_df = _merge(prov_df, p_cur_year, "本年")
                        prov_df = _merge(prov_df, p_last_year, "同期本年")

                        prov_df['今日同比(箱)'] = prov_df.apply(lambda r: _yoy(r.get('今日箱数', 0), r.get('同期今日箱数', 0)), axis=1)
                        prov_df['今日同比(门店)'] = prov_df.apply(lambda r: _yoy(r.get('今日门店数', 0), r.get('同期今日门店数', 0)), axis=1)
                        prov_df['本月同比(箱)'] = prov_df.apply(lambda r: _yoy(r.get('本月箱数', 0), r.get('同期本月箱数', 0)), axis=1)
                        prov_df['本月同比(门店)'] = prov_df.apply(lambda r: _yoy(r.get('本月门店数', 0), r.get('同期本月门店数', 0)), axis=1)
                        prov_df['本年同比(箱)'] = prov_df.apply(lambda r: _yoy(r.get('本年箱数', 0), r.get('同期本年箱数', 0)), axis=1)
                        prov_df['本年同比(门店)'] = prov_df.apply(lambda r: _yoy(r.get('本年门店数', 0), r.get('同期本年门店数', 0)), axis=1)

                        prov_show = pd.DataFrame({
                            '省区': prov_df['省区'],
                            '今日箱数': pd.to_numeric(prov_df.get('今日箱数', 0), errors='coerce').fillna(0),
                            '今日门店数': pd.to_numeric(prov_df.get('今日门店数', 0), errors='coerce').fillna(0),
                            '今日同期(箱数)': pd.to_numeric(prov_df.get('同期今日箱数', 0), errors='coerce').fillna(0),
                            '今日同比(箱)': pd.to_numeric(prov_df.get('今日同比(箱)', None), errors='coerce'),
                            '本月箱数': pd.to_numeric(prov_df.get('本月箱数', 0), errors='coerce').fillna(0),
                            '本月门店数': pd.to_numeric(prov_df.get('本月门店数', 0), errors='coerce').fillna(0),
                            '本月同期(箱数)': pd.to_numeric(prov_df.get('同期本月箱数', 0), errors='coerce').fillna(0),
                            '本月同比(箱)': pd.to_numeric(prov_df.get('本月同比(箱)', None), errors='coerce'),
                            '本年箱数': pd.to_numeric(prov_df.get('本年箱数', 0), errors='coerce').fillna(0),
                            '本年门店数': pd.to_numeric(prov_df.get('本年门店数', 0), errors='coerce').fillna(0),
                            '本年同期(箱数)': pd.to_numeric(prov_df.get('同期本年箱数', 0), errors='coerce').fillna(0),
                            '本年同比(箱)': pd.to_numeric(prov_df.get('本年同比(箱)', None), errors='coerce'),
                        }).fillna({'今日同比(箱)': np.nan, '本月同比(箱)': np.nan, '本年同比(箱)': np.nan})

                        day_txt = f"{int(ctx['kpi_month'])}月{int(ctx['kpi_day'])}日" if ctx["kpi_day"] is not None else f"{int(ctx['kpi_month'])}月"
                        grp_today = f"今日（{day_txt}）"
                        grp_month = f"本月（{int(ctx['kpi_month'])}月）"
                        grp_year = f"本年（{int(ctx['kpi_year'])}年）"

                        col_defs = [
                            {'headerName': '省区', 'field': '省区', 'minWidth': 110, 'headerClass': 'ag-header-center'},
                            {
                                'headerName': grp_today,
                                'children': [
                                    {'headerName': '箱数', 'field': '今日箱数', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '门店数', 'field': '今日门店数', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '同期（箱数）', 'field': '今日同期(箱数)', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '同比（箱）', 'field': '今日同比(箱)', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_PCT_RATIO}, 
                                ],
                            },
                            {
                                'headerName': grp_month,
                                'children': [
                                    {'headerName': '箱数', 'field': '本月箱数', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '门店数', 'field': '本月门店数', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '同期（箱数）', 'field': '本月同期(箱数)', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '同比（箱）', 'field': '本月同比(箱)', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_PCT_RATIO}, 
                                ],
                            },
                            {
                                'headerName': grp_year,
                                'children': [
                                    {'headerName': '箱数', 'field': '本年箱数', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '门店数', 'field': '本年门店数', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '同期（箱数）', 'field': '本年同期(箱数)', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_NUM},
                                    {'headerName': '同比（箱）', 'field': '本年同比(箱)', 'type': ['numericColumn', 'numberColumnFilter'], 'headerClass': 'ag-header-center', 'valueFormatter': JS_FMT_PCT_RATIO}, 
                                ],
                            },
                        ]

                        def _sum_col(col_name: str) -> float:
                            if col_name not in prov_show.columns:
                                return 0.0
                            return float(pd.to_numeric(prov_show[col_name], errors='coerce').fillna(0).sum())

                        _t_cur = _sum_col('今日箱数')
                        _t_last = _sum_col('今日同期(箱数)')
                        _m_cur = _sum_col('本月箱数')
                        _m_last = _sum_col('本月同期(箱数)')
                        _y_cur = _sum_col('本年箱数')
                        _y_last = _sum_col('本年同期(箱数)')

                        pinned_total = {
                            '省区': '合计',
                            '今日箱数': _t_cur,
                            '今日门店数': _sum_col('今日门店数'),
                            '今日同期(箱数)': _t_last,
                            '今日同比(箱)': ((_t_cur - _t_last) / _t_last) if _t_last > 0 else None,
                            '本月箱数': _m_cur,
                            '本月门店数': _sum_col('本月门店数'),
                            '本月同期(箱数)': _m_last,
                            '本月同比(箱)': ((_m_cur - _m_last) / _m_last) if _m_last > 0 else None,
                            '本年箱数': _y_cur,
                            '本年门店数': _sum_col('本年门店数'),
                            '本年同期(箱数)': _y_last,
                            '本年同比(箱)': ((_y_cur - _y_last) / _y_last) if _y_last > 0 else None,
                        }

                        gridOptions = {
                            'pinnedBottomRowData': [pinned_total],
                            'columnDefs': col_defs,
                            'defaultColDef': {
                                'resizable': True,
                                'sortable': True,
                                'filter': True,
                                'wrapHeaderText': True,
                                'autoHeaderHeight': True,
                                'minWidth': 70,
                                'flex': 1,
                                'cellStyle': {'textAlign': 'center', 'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center'},
                                'headerClass': 'ag-header-center',
                            },
                            'rowHeight': 40,
                            'headerHeight': 60,
                            'groupHeaderHeight': 60,
                            'animateRows': True,
                            'suppressCellFocus': True,
                            'enableCellTextSelection': True,
                            'suppressDragLeaveHidesColumns': True,
                            'sideBar': {
                                "toolPanels": [
                                    {
                                        "id": "columns",
                                        "labelDefault": "列",
                                        "iconKey": "columns",
                                        "toolPanel": "agColumnsToolPanel",
                                        "toolPanelParams": {
                                            "suppressRowGroups": True,
                                            "suppressValues": True,
                                            "suppressPivots": True,
                                            "suppressPivotMode": True
                                        }
                                    }
                                ],
                                "defaultToolPanel": None
                            },
                        }

                        AgGrid(
                            prov_show,
                            gridOptions=gridOptions,
                            height=520,
                            width='100%',
                            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                            update_mode=GridUpdateMode.NO_UPDATE,
                            fit_columns_on_grid_load=True,
                            allow_unsafe_jscode=True,
                            theme='streamlit',
                            key="outbound_prov_table"
                        )

                    st.markdown("</div>", unsafe_allow_html=True)

                    # === TAB 7: PERFORMANCE ===
            with tab7:
                st.markdown("""
                <style>
                  .perf-wrap {display:flex; flex-direction:column; gap:16px;}
                  .perf-kpis {display:grid; grid-template-columns: repeat(4, 1fr); gap:14px;}
                  .perf-card {background:#F3E5F5; border:1px solid rgba(156,39,176,0.18); border-radius:12px; padding:14px 16px; box-shadow:0 6px 20px rgba(18,12,28,0.06);}
                  .perf-k {font-size:13px; color:rgba(27,21,48,0.72);}
                  .perf-v {font-size:20px; font-weight:800; color:#9C27B0; margin-top:8px;}
                  .perf-sub {display:flex; justify-content:space-between; align-items:center; margin-top:8px; font-size:12px; color:rgba(27,21,48,0.72);}
                  .perf-up {color:#2FBF71; font-weight:700;}
                  .perf-down {color:#E5484D; font-weight:700;}
                  .perf-mid {color:#FFB000; font-weight:700;}
                  .stDataFrame td { vertical-align: middle !important; }
                  @media (max-width: 1100px) {.perf-kpis {grid-template-columns: repeat(2, 1fr);} }
                </style>
                """, unsafe_allow_html=True)

                if df_perf_raw is None or df_perf_raw.empty:
                    st.warning("⚠️ 未检测到发货业绩数据 (Sheet4)。请确认Excel包含Sheet4且数据完整。")
                    with st.expander("🛠️ 调试信息", expanded=False):
                        for log in debug_logs: st.text(log)
                else:
                    df_perf = df_perf_raw.copy()
                    
                    # --- 1. Data Prep ---
                    # Load Targets from Sheet5 (C=Prov, D=Cat, E=Month, F=Task)
                    df_target = None
                    if df_target_raw is not None and len(df_target_raw.columns) >= 6:
                        try:
                            # Use iloc to be safe about column names
                            df_target = df_target_raw.iloc[:, [2, 3, 4, 5]].copy()
                            df_target.columns = ['省区', '品类', '月份', '任务量']
                            df_target['任务量'] = pd.to_numeric(df_target['任务量'], errors='coerce').fillna(0)
                            df_target['月份'] = pd.to_numeric(df_target['月份'], errors='coerce').fillna(0).astype(int)
                            df_target['省区'] = df_target['省区'].astype(str).str.strip()
                            df_target['品类'] = df_target['品类'].astype(str).str.strip()
                        except Exception as e:
                            st.error(f"任务表解析失败: {e}")
                            df_target = None
                    
                    # Data Cleaning
                    df_track = df_perf.copy()
                    df_track['年份'] = pd.to_numeric(df_track['年份'], errors='coerce').fillna(0).astype(int)
                    df_track['月份'] = pd.to_numeric(df_track['月份'], errors='coerce').fillna(0).astype(int)
                    
                    # Fix: Check if '发货金额' exists, if not, try to use '发货箱数' or create empty
                    if '发货金额' not in df_track.columns:
                            if '发货箱数' in df_track.columns:
                                df_track['发货金额'] = df_track['发货箱数'] # Fallback
                            else:
                                df_track['发货金额'] = 0.0
                    
                    df_track['发货金额'] = pd.to_numeric(df_track['发货金额'], errors='coerce').fillna(0.0)
                    
                    for c in ['省区', '经销商名称', '归类', '发货仓', '大分类', '月分析']:
                        if c in df_track.columns:
                            df_track[c] = df_track[c].fillna('').astype(str).str.strip()
                    
                    # Determine Year
                    years = sorted([y for y in df_track['年份'].unique() if y > 2000])
                    cur_year = 2026 if 2026 in years else (max(years) if years else 2025)
                    last_year = cur_year - 1
                    
                    # --- 2. Filters ---
                    with st.expander("🎛️ 筛选控制面板", expanded=False):
                        f1, f2, f3, f4, f5 = st.columns(5)
                        
                        # Province
                        prov_opts = ['全部'] + sorted([x for x in df_track['省区'].unique() if x])
                        with f1:
                            sel_prov = st.selectbox("省区", prov_opts, key="t26_prov")
                        
                        # Filter Step 1
                        df_f = df_track if sel_prov == '全部' else df_track[df_track['省区'] == sel_prov]
                        
                        # Distributor
                        dist_opts = ['全部'] + sorted([x for x in df_f['经销商名称'].unique() if x])
                        with f2:
                            sel_dist = st.selectbox("经销商", dist_opts, key="t26_dist")
                        if sel_dist != '全部':
                            df_f = df_f[df_f['经销商名称'] == sel_dist]
                            
                        if '大分类' in df_track.columns:
                            cat_col_S = '大分类'
                        elif '月分析' in df_track.columns:
                            cat_col_S = '月分析'
                            st.warning("⚠️ 未找到'Sheet4 S列大分类'字段名“大分类”，已使用“月分析”列作为替代。请确认源数据列名。")
                        else:
                            cat_col_S = '发货仓'
                            st.error("❌ 数据源中未找到Sheet4 S列“大分类”/“月分析”列，已临时使用“发货仓”列作为大分类筛选。")

                        if cat_col_S in df_f.columns:
                            df_f[cat_col_S] = df_f[cat_col_S].fillna('').astype(str).str.strip()

                        if cat_col_S in df_track.columns:
                            df_track[cat_col_S] = df_track[cat_col_S].fillna('').astype(str).str.strip()

                        cat_check_value = "益益成人粉"
                        cat_exists_all = False
                        cat_exists_filtered = False
                        if cat_col_S in df_track.columns:
                            cat_exists_all = bool((df_track[cat_col_S] == cat_check_value).any())
                        if cat_col_S in df_f.columns:
                            cat_exists_filtered = bool((df_f[cat_col_S] == cat_check_value).any())

                        if cat_exists_all and (not cat_exists_filtered):
                            st.warning(f"⚠️ 源数据“大分类”包含“{cat_check_value}”，但在当前省区/经销商筛选下无数据。请调整筛选查看。")

                        with st.expander("🔎 大分类数据校验", expanded=False):
                            if cat_col_S not in df_track.columns:
                                st.error(f"未找到用于大分类的字段：{cat_col_S}")
                            else:
                                s_all = df_track[cat_col_S]
                                s_all_nonempty = s_all[s_all != ""]
                                st.write(f"大分类字段：{cat_col_S}")
                                st.write(f"唯一类目数：{int(s_all_nonempty.nunique())}")
                                st.write(f"空值占比：{fmt_pct_ratio(float((s_all == '').mean()))}")
                                st.write(f"是否包含“{cat_check_value}”：{'是' if cat_exists_all else '否'}")
                                top_counts = s_all_nonempty.value_counts().head(12).reset_index()
                                top_counts.columns = ["类目", "行数"]
                                show_aggrid_table(top_counts, height=300, key="verify_table")

                        wh_opts = ['全部'] + sorted([x for x in df_f.get(cat_col_S, pd.Series(dtype=str)).unique() if x])
                        with f3:
                            sel_wh = st.selectbox(f"大类 ({cat_col_S})", wh_opts, key="t26_wh")
                        
                        if sel_wh != '全部':
                            df_f = df_f[df_f.get(cat_col_S, pd.Series(dtype=str)) == sel_wh]
                            
                        # Small Category (Group) - Multi Select
                        grp_opts = sorted([x for x in df_f['归类'].unique() if x])
                        with f4:
                            sel_grp = st.multiselect("小类 (归类)", grp_opts, default=[], key="t26_grp")
                        if sel_grp:
                            df_f = df_f[df_f['归类'].isin(sel_grp)]
                            
                        # Month Selection (Single)
                        avail_months = sorted(df_f[df_f['年份'] == cur_year]['月份'].unique())
                        def_month = int(avail_months[-1]) if avail_months else 1
                        with f5:
                            sel_month = st.selectbox("统计月份", list(range(1, 13)), index=def_month-1, key="t26_month")
                    
                    # --- 3. Calculations ---
                    # Actuals
                    act_cur_year = df_f[df_f['年份'] == cur_year]['发货金额'].sum()
                    act_last_year = df_f[df_f['年份'] == last_year]['发货金额'].sum()
                    
                    act_cur_month = df_f[(df_f['年份'] == cur_year) & (df_f['月份'] == sel_month)]['发货金额'].sum()
                    act_last_month = df_f[(df_f['年份'] == last_year) & (df_f['月份'] == sel_month)]['发货金额'].sum()
                    
                    # Targets
                    target_cur_year = 0.0
                    target_cur_month = 0.0
                    if df_target is not None:
                        # Apply filters to target (Province, Category)
                        # Note: Distributor filter can't apply to Target usually, unless target is by dist. 
                        # User said Sheet5 has Province/Category.
                        df_t_f = df_target.copy()
                        if sel_prov != '全部':
                            df_t_f = df_t_f[df_t_f['省区'] == sel_prov]
                        # Category mapping? Sheet5 '品类' vs Sheet4 '归类'/'发货仓'.
                        # User said D col is Category. Assuming it matches '归类' or needs mapping.
                        # For now, we sum all if no specific match logic provided or if '全部'.
                        # If user selected specific categories, we try to filter.
                        # BUT, without exact mapping, filtering Targets by Category is risky. 
                        # We'll calculate Total Target for selected Province.
                        
                        target_cur_year = df_t_f['任务量'].sum()
                        target_cur_month = df_t_f[df_t_f['月份'] == sel_month]['任务量'].sum()
                    
                    # Rates & YoY
                    rate_year = (act_cur_year / target_cur_year) if target_cur_year > 0 else None
                    rate_month = (act_cur_month / target_cur_month) if target_cur_month > 0 else None
                    
                    yoy_year = (act_cur_year - act_last_year) / act_last_year if act_last_year > 0 else None
                    yoy_month = (act_cur_month - act_last_month) / act_last_month if act_last_month > 0 else None
                    
                    # --- 4. KPI Cards ---
                    def _fmt_wan(x): return fmt_num((x or 0) / 10000)
                    def _fmt_pct(x): return fmt_pct_ratio(x) if x is not None else "—"
                    def _color_pct(x): return "perf-up" if x and x>0 else "perf-down"
                    def _arrow(x): return "↑" if x and x>0 else ("↓" if x and x<0 else "")

                    def _render_card(title, icon, val_wan, target_wan, rate, yoy_val_wan, yoy_pct):
                        trend_cls = "trend-up" if yoy_pct and yoy_pct > 0 else ("trend-down" if yoy_pct and yoy_pct < 0 else "trend-neutral")
                        arrow = _arrow(yoy_pct)
                        rate_txt = _fmt_pct(rate)
                        yoy_txt = _fmt_pct(yoy_pct)
                        pct_val = min(max(rate * 100 if rate else 0, 0), 100)
                        prog_color = "#28A745" if rate and rate >= 1.0 else ("#FFC107" if rate and rate >= 0.8 else "#DC3545")

                        st.markdown(f"""
                        <div class="out-kpi-card">
                            <div class="out-kpi-bar"></div>
                            <div class="out-kpi-head">
                                <div class="out-kpi-ico">{icon}</div>
                                <div class="out-kpi-title">{title}</div>
                            </div>
                            <div class="out-kpi-val">¥ {val_wan}万</div>
                            <div class="out-kpi-sub2" style="margin-top:8px;">
                                <span>达成率</span>
                                <span style="font-weight:800; color:{prog_color}">{rate_txt}</span>
                            </div>
                            <div class="out-kpi-progress" style="margin-top:6px;">
                                <div class="out-kpi-progress-bar" style="background:{prog_color}; width:{pct_val}%;"></div>
                            </div>
                            <div class="out-kpi-sub2" style="margin-top:10px;">
                                <span>目标</span>
                                <span>{target_wan}万</span>
                            </div>
                            <div class="out-kpi-sub2">
                                <span>同期</span>
                                <span>{yoy_val_wan}万</span>
                            </div>
                            <div class="out-kpi-sub2">
                                <span>同比</span>
                                <span class="{trend_cls}">{arrow} {yoy_txt}</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                    # --- TABS: KPI, Category, Province ---
                    tab_perf_kpi, tab_perf_cat, tab_perf_prov = st.tabs(["📊 核心业绩指标", "📦 分品类", "🗺️ 分省区"])

                    with tab_perf_kpi:
                        k1, k2 = st.columns(2)
                        
                        with k1:
                            _render_card("本月业绩", "📅", _fmt_wan(act_cur_month), _fmt_wan(target_cur_month), rate_month, _fmt_wan(act_last_month), yoy_month)
                        with k2:
                            _render_card("年度累计业绩", "🏆", _fmt_wan(act_cur_year), _fmt_wan(target_cur_year), rate_year, _fmt_wan(act_last_year), yoy_year)
                    
                    with tab_perf_cat:
                        # --- NEW: Category Performance Cards ---
                        
                        # Prepare Category Data
                        # Using cat_col_S ('大分类' or '月分析' or '发货仓')
                        
                        # 1. Monthly Category Data
                        cat_cur_m = df_f[(df_f['年份'] == cur_year) & (df_f['月份'] == sel_month)].groupby(cat_col_S)['发货金额'].sum().reset_index().rename(columns={'发货金额': '本月'})
                        cat_last_m = df_f[(df_f['年份'] == last_year) & (df_f['月份'] == sel_month)].groupby(cat_col_S)['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期'})
                        
                        cat_m_final = pd.merge(cat_cur_m, cat_last_m, on=cat_col_S, how='outer').fillna(0)
                        cat_m_final['本月(万)'] = cat_m_final['本月'] / 10000
                        cat_m_final['同期(万)'] = cat_m_final['同期'] / 10000
                        cat_m_final['同比'] = np.where(cat_m_final['本月'] > 0, (cat_m_final['本月'] - cat_m_final['同期']) / cat_m_final['本月'], None)
                        cat_m_final = cat_m_final.sort_values('本月', ascending=False)

                        # 2. Yearly Category Data
                        cat_cur_y = df_f[df_f['年份'] == cur_year].groupby(cat_col_S)['发货金额'].sum().reset_index().rename(columns={'发货金额': '本年'})
                        cat_last_y = df_f[df_f['年份'] == last_year].groupby(cat_col_S)['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期'})
                        
                        cat_y_final = pd.merge(cat_cur_y, cat_last_y, on=cat_col_S, how='outer').fillna(0)
                        cat_y_final['本年(万)'] = cat_y_final['本年'] / 10000
                        cat_y_final['同期(万)'] = cat_y_final['同期'] / 10000
                        cat_y_final['同比'] = np.where(cat_y_final['本年'] > 0, (cat_y_final['本年'] - cat_y_final['同期']) / cat_y_final['本年'], None)
                        cat_y_final = cat_y_final.sort_values('本年', ascending=False)

                        # Render 2 Columns for Tables
                        c_cat_m, c_cat_y = st.columns(2)

                        with c_cat_m:
                            st.markdown(
                                """
                                <div style="background-color: #F8F9FA; border-radius: 8px; padding: 16px; border: 1px solid #E9ECEF; box-shadow: 0 2px 4px rgba(0,0,0,0.05); height: 100%;">
                                    <div style="font-size: 14px; color: #6C757D; margin-bottom: 12px; font-weight: 500;">📅 本月分品类业绩</div>
                                """, 
                                unsafe_allow_html=True
                            )
                            # Replaced with AgGrid
                            show_aggrid_table(
                                cat_m_final[[cat_col_S, '本月(万)', '同期(万)', '同比']],
                                height=250,
                                key="ag_cat_m"
                            )
                            
                            # Donut Chart for Month
                            if not cat_m_final.empty and cat_m_final['本月(万)'].sum() > 0:
                                total_m = cat_m_final['本月(万)'].sum()
                                cat_m_final['legend_label'] = cat_m_final.apply(
                                    lambda r: f"{r[cat_col_S]}   {r['本月(万)']:.1f}万   {r['本月(万)']/total_m:.1%}", axis=1
                                )
                                
                                fig_m = go.Figure(data=[go.Pie(
                                    labels=cat_m_final['legend_label'],
                                    values=cat_m_final['本月(万)'],
                                    hole=0.6,
                                    marker=dict(colors=px.colors.qualitative.Pastel),
                                    textinfo='none',
                                    domain={'x': [0.4, 1.0]}
                                )])
                                fig_m.update_layout(
                                    showlegend=True,
                                    legend=dict(
                                        yanchor="middle", y=0.5,
                                        xanchor="left", x=0,
                                        font=dict(size=12, color="#333333")
                                    ),
                                    margin=dict(t=10, b=10, l=0, r=0), 
                                    height=250
                                )
                                st.plotly_chart(fig_m, use_container_width=True, key="perf_cat_month_donut")
                            else:
                                st.info("暂无数据")
                                
                            st.markdown("</div>", unsafe_allow_html=True)

                        with c_cat_y:
                            st.markdown(
                                """
                                <div style="background-color: #F8F9FA; border-radius: 8px; padding: 16px; border: 1px solid #E9ECEF; box-shadow: 0 2px 4px rgba(0,0,0,0.05); height: 100%;">
                                    <div style="font-size: 14px; color: #6C757D; margin-bottom: 12px; font-weight: 500;">🏆 年度分品类业绩</div>
                                """, 
                                unsafe_allow_html=True
                            )
                            # Replaced with AgGrid
                            show_aggrid_table(
                                cat_y_final[[cat_col_S, '本年(万)', '同期(万)', '同比']],
                                height=250,
                                key="ag_cat_y"
                            )
                            
                            # Donut Chart for Year
                            if not cat_y_final.empty and cat_y_final['本年(万)'].sum() > 0:
                                total_y = cat_y_final['本年(万)'].sum()
                                cat_y_final['legend_label'] = cat_y_final.apply(
                                    lambda r: f"{r[cat_col_S]}   {r['本年(万)']:.1f}万   {r['本年(万)']/total_y:.1%}", axis=1
                                )
                                
                                fig_y = go.Figure(data=[go.Pie(
                                    labels=cat_y_final['legend_label'],
                                    values=cat_y_final['本年(万)'],
                                    hole=0.6,
                                    marker=dict(colors=px.colors.qualitative.Pastel),
                                    textinfo='none',
                                    domain={'x': [0.4, 1.0]}
                                )])
                                fig_y.update_layout(
                                    showlegend=True,
                                    legend=dict(
                                        yanchor="middle", y=0.5,
                                        xanchor="left", x=0,
                                        font=dict(size=12, color="#333333")
                                    ),
                                    margin=dict(t=10, b=10, l=0, r=0), 
                                    height=250
                                )
                                st.plotly_chart(fig_y, use_container_width=True, key="perf_cat_year_donut")
                            else:
                                st.info("暂无数据")
                                
                            st.markdown("</div>", unsafe_allow_html=True)

                    with tab_perf_prov:
                        # --- 5. Province Table (Detailed) ---
                        
                        # Prepare Data
                        # Group by Province
                        # 1. Actuals (Cur Month)
                        df_m_cur = df_f[(df_f['年份'] == cur_year) & (df_f['月份'] == sel_month)]
                        prov_cur = df_m_cur.groupby('省区')['发货金额'].sum().reset_index().rename(columns={'发货金额': '本月业绩'})
                        
                        # 2. Actuals (Same Period)
                        df_m_last = df_f[(df_f['年份'] == last_year) & (df_f['月份'] == sel_month)]
                        prov_last = df_m_last.groupby('省区')['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期业绩'})
                        
                        # 3. Targets (Month)
                        if df_target is not None:
                            t_m = df_target[df_target['月份'] == sel_month]
                            prov_target = t_m.groupby('省区')['任务量'].sum().reset_index().rename(columns={'任务量': '本月任务'})
                        else:
                            prov_target = pd.DataFrame(columns=['省区', '本月任务'])
                            
                        # Merge All
                        prov_final = pd.merge(prov_cur, prov_target, on='省区', how='outer')
                        prov_final = pd.merge(prov_final, prov_last, on='省区', how='outer').fillna(0)
                        
                        # Filter out rows with 0
                        prov_final = prov_final[(prov_final['本月业绩']!=0) | (prov_final['本月任务']!=0) | (prov_final['同期业绩']!=0)]
                        
                        # Metrics
                        prov_final['达成率'] = prov_final.apply(lambda x: (x['本月业绩'] / x['本月任务']) if x['本月任务'] > 0 else 0, axis=1)
                        prov_final['同比增长'] = prov_final.apply(lambda x: ((x['本月业绩'] - x['同期业绩']) / x['同期业绩']) if x['同期业绩'] > 0 else 0, axis=1)
                        
                        # Sort
                        prov_final = prov_final.sort_values('本月业绩', ascending=False)
                        
                        # Format for Display
                        prov_final['本月业绩(万)'] = prov_final['本月业绩'] / 10000
                        prov_final['本月任务(万)'] = prov_final['本月任务'] / 10000
                        prov_final['同期业绩(万)'] = prov_final['同期业绩'] / 10000
                        
                        # Display Columns
                        disp_df = prov_final[['省区', '本月业绩(万)', '本月任务(万)', '达成率', '同期业绩(万)', '同比增长']].copy()
                        
                        # Interactive Table
                        st.caption("👇 点击表格行可下钻查看详细数据")
                        
                        # AgGrid for Province Performance
                        ag_prov = show_aggrid_table(
                            disp_df, 
                            key="perf_prov_ag",
                            on_row_selected=True
                        )
                        
                        # Drill Down
                        # Check if selected_rows exists and is not empty
                        selected_rows = ag_prov.get('selected_rows') if ag_prov else None
                        
                        if selected_rows is not None and len(selected_rows) > 0:
                            # AgGrid return structure might differ based on version
                            # Sometimes it returns a DataFrame, sometimes a list of dicts
                            if isinstance(selected_rows, pd.DataFrame):
                                first_row = selected_rows.iloc[0]
                            else:
                                first_row = selected_rows[0]
                                
                            # Handle if it returns a DataFrame row or dict
                            sel_prov_drill = first_row.get('省区') if isinstance(first_row, dict) else first_row['省区']
                            
                            # Drill Down Tabs
                            st.markdown("---")
                            st.subheader(f"📍 {sel_prov_drill} - 明细数据")
                            
                            tab_dist, tab_cat = st.tabs(["🏢 经销商明细", "📦 品类明细"])
                            
                            # Filter data for selected province
                            d_cur = df_f[(df_f['年份'] == cur_year) & (df_f['月份'] == sel_month) & (df_f['省区'] == sel_prov_drill)]
                            d_last = df_f[(df_f['年份'] == last_year) & (df_f['月份'] == sel_month) & (df_f['省区'] == sel_prov_drill)]

                            # --- Tab 1: Distributor Drill Down ---
                            with tab_dist:
                                st.caption(f"正在查看：{sel_prov_drill} > 经销商明细")
                                d_cur_g = d_cur.groupby('经销商名称')['发货金额'].sum().reset_index().rename(columns={'发货金额': '本月'})
                                d_last_g = d_last.groupby('经销商名称')['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期'})
                                
                                d_final = pd.merge(d_cur_g, d_last_g, on='经销商名称', how='outer').fillna(0)
                                d_final['同比增长'] = d_final.apply(lambda x: ((x['本月'] - x['同期']) / x['同期']) if x['同期'] > 0 else 0, axis=1)
                                d_final = d_final.sort_values('本月', ascending=False)
                                
                                d_final['本月(万)'] = d_final['本月'] / 10000
                                d_final['同期(万)'] = d_final['同期'] / 10000
                                
                                ag_dist = show_aggrid_table(
                                    d_final[['经销商名称', '本月(万)', '同期(万)', '同比增长']],
                                    key="perf_dist_ag",
                                    on_row_selected=True
                                )
                                
                                selected_rows_dist = ag_dist.get('selected_rows') if ag_dist else None
                                
                                if selected_rows_dist is not None and len(selected_rows_dist) > 0:
                                    if isinstance(selected_rows_dist, pd.DataFrame):
                                        first_row_dist = selected_rows_dist.iloc[0]
                                    else:
                                        first_row_dist = selected_rows_dist[0]
                                        
                                    sel_dist_drill = first_row_dist.get('经销商名称') if isinstance(first_row_dist, dict) else first_row_dist['经销商名称']
                                    st.info(f"📍 正在查看 {sel_prov_drill} > {sel_dist_drill} 的大分类明细")
                                    
                                    if '大分类' in d_cur.columns:
                                        cat_col_S = '大分类'
                                    elif '月分析' in d_cur.columns:
                                        cat_col_S = '月分析'
                                    else:
                                        cat_col_S = '发货仓'
                                    
                                    # Filter data for selected dist
                                    bc_cur = d_cur[d_cur['经销商名称'] == sel_dist_drill]
                                    bc_last = d_last[d_last['经销商名称'] == sel_dist_drill]
                                    
                                    bc_cur_g = bc_cur.groupby(cat_col_S)['发货金额'].sum().reset_index().rename(columns={'发货金额': '本月'})
                                    bc_last_g = bc_last.groupby(cat_col_S)['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期'})
                                    
                                    bc_final = pd.merge(bc_cur_g, bc_last_g, on=cat_col_S, how='outer').fillna(0)
                                    bc_final['同比增长'] = bc_final.apply(lambda x: ((x['本月'] - x['同期']) / x['同期']) if x['同期'] > 0 else 0, axis=1)
                                    bc_final = bc_final.sort_values('本月', ascending=False)
                                    
                                    bc_final['本月(万)'] = bc_final['本月'] / 10000
                                    bc_final['同期(万)'] = bc_final['同期'] / 10000
                                    
                                    ag_bc = show_aggrid_table(
                                        bc_final[[cat_col_S, '本月(万)', '同期(万)', '同比增长']],
                                        key="perf_bc_table_dist_ag",
                                        on_row_selected=True
                                    )
                                    
                                    selected_rows_bc = ag_bc.get('selected_rows') if ag_bc else None
                                    
                                    if selected_rows_bc is not None and len(selected_rows_bc) > 0:
                                        if isinstance(selected_rows_bc, pd.DataFrame):
                                            first_row_bc = selected_rows_bc.iloc[0]
                                        else:
                                            first_row_bc = selected_rows_bc[0]
                                            
                                        sel_bc_drill = first_row_bc.get(cat_col_S) if isinstance(first_row_bc, dict) else first_row_bc[cat_col_S]
                                        st.info(f"📍 正在查看 {sel_prov_drill} > {sel_dist_drill} > {sel_bc_drill} 的小分类(归类)明细")
                                        
                                        # Level 4: Small Category (Group) for Selected Big Cat
                                        sc_cur = bc_cur[bc_cur[cat_col_S] == sel_bc_drill]
                                        sc_last = bc_last[bc_last[cat_col_S] == sel_bc_drill]
                                        
                                        sc_cur_g = sc_cur.groupby('归类')['发货金额'].sum().reset_index().rename(columns={'发货金额': '本月'})
                                        sc_last_g = sc_last.groupby('归类')['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期'})
                                        
                                        sc_final = pd.merge(sc_cur_g, sc_last_g, on='归类', how='outer').fillna(0)
                                        sc_final['同比增长'] = sc_final.apply(lambda x: ((x['本月'] - x['同期']) / x['同期']) if x['同期'] > 0 else 0, axis=1)
                                        sc_final = sc_final.sort_values('本月', ascending=False)
                                        
                                        sc_final['本月(万)'] = sc_final['本月'] / 10000
                                        sc_final['同期(万)'] = sc_final['同期'] / 10000
                                        
                                        show_aggrid_table(
                                            sc_final[['归类', '本月(万)', '同期(万)', '同比增长']],
                                            key="perf_sc_table_dist_ag"
                                        )

                            with tab_cat:
                                st.caption(f"正在查看：{sel_prov_drill} > 品类明细 (按大分类聚合)")
                                if '大分类' in d_cur.columns:
                                    agg_col = '大分类'
                                elif '月分析' in d_cur.columns:
                                    agg_col = '月分析'
                                else:
                                    agg_col = '发货仓'
                                
                                c_cur_g = d_cur.groupby(agg_col)['发货金额'].sum().reset_index().rename(columns={'发货金额': '本月'})
                                c_last_g = d_last.groupby(agg_col)['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期'})
                                
                                c_final = pd.merge(c_cur_g, c_last_g, on=agg_col, how='outer').fillna(0)
                                c_final['同比增长'] = c_final.apply(lambda x: ((x['本月'] - x['同期']) / x['同期']) if x['同期'] > 0 else 0, axis=1)
                                c_final = c_final.sort_values('本月', ascending=False)
                                
                                c_final['本月(万)'] = c_final['本月'] / 10000
                                c_final['同期(万)'] = c_final['同期'] / 10000
                                
                                ag_cat = show_aggrid_table(
                                    c_final[[agg_col, '本月(万)', '同期(万)', '同比增长']],
                                    key="perf_cat_table_ag",
                                    on_row_selected=True
                                )
                                
                                selected_rows_cat = ag_cat.get('selected_rows') if ag_cat else None
                                
                                if selected_rows_cat is not None and len(selected_rows_cat) > 0:
                                    if isinstance(selected_rows_cat, pd.DataFrame):
                                        first_row_cat = selected_rows_cat.iloc[0]
                                    else:
                                        first_row_cat = selected_rows_cat[0]
                                        
                                    sel_cat_drill = first_row_cat.get(agg_col) if isinstance(first_row_cat, dict) else first_row_cat[agg_col]
                                    st.info(f"📍 正在查看 {sel_prov_drill} > {sel_cat_drill} 的小分类(归类)明细")
                                    
                                    # Level 3: Small Category (Group) for Selected Big Cat (Province Level)
                                    sc_cur = d_cur[d_cur[agg_col] == sel_cat_drill]
                                    sc_last = d_last[d_last[agg_col] == sel_cat_drill]
                                    
                                    sc_cur_g = sc_cur.groupby('归类')['发货金额'].sum().reset_index().rename(columns={'发货金额': '本月'})
                                    sc_last_g = sc_last.groupby('归类')['发货金额'].sum().reset_index().rename(columns={'发货金额': '同期'})
                                    
                                    sc_final = pd.merge(sc_cur_g, sc_last_g, on='归类', how='outer').fillna(0)
                                    sc_final['同比增长'] = sc_final.apply(lambda x: ((x['本月'] - x['同期']) / x['同期']) if x['同期'] > 0 else 0, axis=1)
                                    sc_final = sc_final.sort_values('本月', ascending=False)
                                    
                                    sc_final['本月(万)'] = sc_final['本月'] / 10000
                                    sc_final['同期(万)'] = sc_final['同期'] / 10000
                                    
                                    # Dynamic height
                                    n_rows_sc2 = len(sc_final)
                                    calc_height_sc2 = (n_rows_sc2 + 1) * 35 + 10
                                    final_height_sc2 = max(150, min(calc_height_sc2, 2000))
                                    
                                    show_aggrid_table(
                                        sc_final[['归类', '本月(万)', '同期(万)', '同比增长']],
                                        height=final_height_sc2,
                                        key="perf_sc_table_cat_ag"
                                    )

else:
    st.info("请在左侧上传数据文件以开始分析。")

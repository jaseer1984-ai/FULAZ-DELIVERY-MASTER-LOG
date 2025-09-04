# -*- coding: utf-8 -*-
# Enhanced Excel → Streamlit Dashboard for FULAZ Delivery MasterLog
# - Robust .xlsx/.xls handling without caching unserializable objects
# - Date filters extracted from first row (header dates)
# - ALL CAPS formatting for professional report presentation
# - Top-positioned comprehensive filter section
# - Enhanced metric cards and analytics tabs
# Author: AI Assistant | Enhanced for FULAZ Professional Dashboard

import io
from typing import List, Optional, Tuple
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ---------------------------- Config ----------------------------
st.set_page_config(
    page_title="FULAZ DELIVERY MASTERLOG DASHBOARD",
    layout="wide",
    page_icon="🏗️",
    initial_sidebar_state="expanded"
)

# ---------------------------- Small formatters ----------------------------
def fmt_num(x, dec=2):
    try:
        return f"{float(x):,.{dec}f}".upper()
    except Exception:
        return "--"

def fmt_pct(x):
    try:
        return f"{float(x):.1f}%"
    except Exception:
        return "--"

# ---------------------------- Cache-safe upload bytes ----------------------------
@st.cache_data(show_spinner=False)
def _get_upload_bytes(uploaded_file) -> Tuple[bytes, str]:
    """
    Return (file_bytes, filename). This is pickle-serializable and safe for st.cache_data.
    """
    if hasattr(uploaded_file, "getvalue"):
        return uploaded_file.getvalue(), getattr(uploaded_file, "name", "") or ""
    # Fallback for file-like
    return uploaded_file.read(), getattr(uploaded_file, "name", "") or ""

# ---------------------------- Excel helpers (no caching of ExcelFile) ----------------------------
def _open_excel_from_bytes(file_bytes: bytes, filename: str) -> pd.ExcelFile:
    """
    Create a pd.ExcelFile from bytes using correct engine.
    Not cached because ExcelFile isn't pickle-serializable.
    """
    is_xls = filename.lower().endswith(".xls")
    engine = "xlrd" if is_xls else "openpyxl"
    return pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)

@st.cache_data(show_spinner=False)
def load_excel(uploaded_file, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    Load Excel into a DataFrame with header-row auto-detect.
    Returns a DataFrame (pickle-serializable) so caching is safe.
    """
    file_bytes, filename = _get_upload_bytes(uploaded_file)
    xobj = _open_excel_from_bytes(file_bytes, filename)

    def read_sheet(xobj: pd.ExcelFile, sheet_nm: Optional[str]) -> pd.DataFrame:
        target = sheet_nm or xobj.sheet_names[0]
        # Try header at Excel rows 3,2,1
        for hdr in (2, 1, 0):
            try:
                df_try = xobj.parse(target, header=hdr)
                df_try.columns = [str(c).strip().upper() for c in df_try.columns]
                good_cols = sum(not str(c).lower().startswith("unnamed") for c in df_try.columns)
                if good_cols >= 5:
                    return df_try
            except Exception:
                pass
        # Fallback
        df = xobj.parse(target)
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df

    return read_sheet(xobj, sheet_name)

def extract_header_dates(uploaded_file, sheet_name: Optional[str] = None) -> List[tuple]:
    """
    Extract dates from first row only. Builds ExcelFile locally (not cached) from cached bytes.
    """
    try:
        file_bytes, filename = _get_upload_bytes(uploaded_file)
        xobj = _open_excel_from_bytes(file_bytes, filename)
        target = sheet_name or xobj.sheet_names[0]
        first = xobj.parse(target, header=None, nrows=1).iloc[0].tolist()
        out = []
        for i, cell in enumerate(first):
            if pd.notna(cell):
                if isinstance(cell, datetime):
                    out.append((i, cell.date()))
                else:
                    dt = pd.to_datetime(cell, errors="coerce")
                    if pd.notna(dt):
                        out.append((i, dt.date()))
        return out
    except Exception as e:
        st.warning(f"Could not extract header dates: {e}")
        return []

def detect_truck_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if str(c).strip().upper().startswith("TRUCK")]

def numericify(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def find_date_cols(df: pd.DataFrame) -> List[str]:
    """Enhanced date column detection."""
    named = [c for c in df.columns if "date" in str(c).lower()]
    dtyped = list(df.select_dtypes(include=["datetime64[ns]", "datetime64[ns, UTC]"]).columns)

    for c in df.columns:
        if c not in named and c not in dtyped:
            s = df[c].dropna()
            if len(s) > 0:
                try:
                    sample = s.head(min(20, len(s)))
                    parsed = pd.to_datetime(sample, errors="coerce")
                    if parsed.notna().sum() >= max(1, int(0.5 * len(sample))):
                        named.append(c)
                except Exception:
                    pass

    seen, out = set(), []
    for c in named + dtyped:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out

def coerce_dates(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

# ---------------------------- CSS ----------------------------
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        text-transform: uppercase;
        letter-spacing: 2px;
    }
    .filter-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
    }
    .filter-header {
        color: white;
        font-size: 1.8rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 1rem;
        text-transform: uppercase;
    }
    .metric-card {
        background: linear-gradient(145deg, #ffffff, #f0f2f6);
        padding: 1.5rem;
        border-radius: 15px;
        border-left: 6px solid #1f77b4;
        margin: 0.5rem 0;
        box-shadow: 0 8px 25px rgba(0,0,0,0.08);
        transition: all 0.3s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 35px rgba(0,0,0,0.12);
    }
    .metric-title {
        font-size: 0.9rem;
        font-weight: 600;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 0.5rem;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 0.3rem;
    }
    .metric-delta {
        font-size: 0.8rem;
        color: #28a745;
        font-weight: 500;
    }
    .section-header {
        font-size: 1.8rem;
        font-weight: bold;
        color: #2c3e50;
        margin: 2rem 0 1rem 0;
        border-bottom: 3px solid #1f77b4;
        padding-bottom: 0.5rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .info-box {
        background: linear-gradient(90deg, #e3f2fd, #ffffff);
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #2196f3;
        margin: 1rem 0;
        font-weight: 500;
        text-transform: uppercase;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------- Sidebar upload ----------------------------
st.sidebar.markdown("### 📄 FILE UPLOAD")
uploaded = st.sidebar.file_uploader("UPLOAD FULAZ DELIVERY MASTERLOG", type=["xlsx", "xls"])

if not uploaded:
    st.markdown('<div class="main-header">🏗️ FULAZ DELIVERY MASTERLOG DASHBOARD</div>', unsafe_allow_html=True)
    st.info("👈 PLEASE UPLOAD YOUR **FULAZ DELIVERY MASTERLOG** EXCEL FILE TO GET STARTED.")
    st.markdown("""
    ### EXPECTED FILE STRUCTURE:
    - **CUSTOMER NAME**, **PROJECT NAME**, **ZONE/LOCATION**
    - **ITEM NAME**, **DELIVERED WEIGHT**, **DELIVERED QTY**
    - **PROGRESS %**, **TRUCK COLUMNS** (TRUCK1, TRUCK2, ...)
    - **DATE COLUMNS** (AUTOMATICALLY DETECTED)
    """)
    st.stop()

# Peek sheet names (build ExcelFile locally from cached bytes)
try:
    _bytes, _name = _get_upload_bytes(uploaded)
    _peek = _open_excel_from_bytes(_bytes, _name)
    sheet_options = _peek.sheet_names
except Exception:
    sheet_options = ["Sheet1"]
sheet_name = st.sidebar.selectbox("📋 SELECT SHEET", options=sheet_options, index=0)

# ---------------------------- Load & Prepare Data ----------------------------
with st.spinner("LOADING AND PROCESSING DATA..."):
    header_dates = extract_header_dates(uploaded, sheet_name=sheet_name)
    data = load_excel(uploaded, sheet_name=sheet_name)
    if data.empty:
        st.error("❌ FAILED TO LOAD DATA FROM EXCEL FILE")
        st.info("Please check that your file is a valid Excel (.xlsx or .xls) and try again.")
        st.stop()
    data.columns = [c.strip().upper() for c in data.columns]

# Key columns
COL_CUSTOMER = "CUSTOMER NAME"
COL_PROJECT = "PROJECT NAME"
COL_PROJECT_NUM = "PROJECT NUMBER"
COL_ZONE = "ZONE / LOCATION"
COL_ITEM = "ITEM NAME"
COL_ITEM_DESC = "ITEM DESCRIPTION"
COL_WEIGHT = "DELIVERED WEIGHT"
COL_QTY = "DELIVERED QTY"
COL_PROGRESS = "PROGRESS %"
COL_CONTRACTED_WEIGHT = "CONTRACTED WEIGHT"
COL_CONTRACTED_QTY = "CONTRACTED QTY"
COL_BALANCE_WEIGHT = "BALANCE WEIGHT"
COL_BALANCE_QTY = "BALANCE QTY"

# Dates & numerics
date_cols = find_date_cols(data)
data = coerce_dates(data, date_cols)

truck_cols = detect_truck_cols(data)
numeric_cols_to_cast = [
    c for c in [COL_WEIGHT, COL_QTY, COL_PROGRESS, COL_CONTRACTED_WEIGHT,
                COL_CONTRACTED_QTY, COL_BALANCE_WEIGHT, COL_BALANCE_QTY]
    if c in data.columns
] + truck_cols
data = numericify(data, numeric_cols_to_cast)

# Normalize progress 0–1 if needed
if COL_PROGRESS in data.columns:
    max_progress = data[COL_PROGRESS].max()
    if pd.notna(max_progress) and max_progress > 1:
        data[COL_PROGRESS] = data[COL_PROGRESS] / 100.0

# ---------------------------- Main Header ----------------------------
st.markdown('<div class="main-header">🏗️ FULAZ DELIVERY MASTERLOG DASHBOARD</div>', unsafe_allow_html=True)

# ---------------------------- FILTERS ----------------------------
st.markdown('<div class="filter-container">', unsafe_allow_html=True)
st.markdown('<div class="filter-header">🔍 COMPREHENSIVE FILTERS</div>', unsafe_allow_html=True)

active_date_col = None  # safety init

# Row 1: Date filters
st.markdown("#### 📅 DATE RANGE FILTERS")
date_col1, date_col2, date_col3, date_col4 = st.columns(4)

with date_col1:
    if header_dates:
        st.markdown("**HEADER DATES AVAILABLE**")
        header_date_options = ["ALL DATES"] + [f"{d.strftime('%Y-%m-%d')}" for _col, d in header_dates]
        selected_header_dates = st.multiselect(
            "SELECT HEADER DATES",
            header_date_options,
            default=["ALL DATES"]
        )
        if "ALL DATES" not in selected_header_dates:
            selected_date_values = []
            for _col, d in header_dates:
                if d.strftime('%Y-%m-%d') in selected_header_dates:
                    selected_date_values.append(d)
            if selected_date_values:
                st.success(f"FILTERED BY {len(selected_date_values)} HEADER DATES")
                # (Optional) add per-header-date column logic here if needed.

with date_col2:
    if date_cols:
        active_date_col = st.selectbox("📅 DATE COLUMN", options=["NONE"] + date_cols, index=0)

with date_col3, date_col4:
    if date_cols and active_date_col and active_date_col != "NONE":
        dseries = pd.to_datetime(data[active_date_col], errors="coerce")
        valid = dseries.dropna()
        if len(valid) > 0:
            min_d, max_d = valid.min().date(), valid.max().date()
            start_date = st.date_input("START DATE", value=min_d, min_value=min_d, max_value=max_d, key="start_date_input")
            end_date = st.date_input("END DATE", value=max_d, min_value=min_d, max_value=max_d, key="end_date_input")
            if start_date and end_date:
                mask = (dseries.dt.date >= start_date) & (dseries.dt.date <= end_date)
                data = data[mask]
                st.success(f"DATE FILTERED: {start_date} TO {end_date}")

# Row 2: Business filters
st.markdown("#### 🏢 BUSINESS DIMENSION FILTERS")
filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)

with filter_col1:
    if COL_CUSTOMER in data.columns:
        customers = ["ALL"] + sorted(data[COL_CUSTOMER].dropna().astype(str).unique())
        selected_customers = st.multiselect("🏢 CUSTOMERS", customers, default=["ALL"])
        if "ALL" not in selected_customers:
            data = data[data[COL_CUSTOMER].astype(str).isin(selected_customers)]

with filter_col2:
    if COL_ZONE in data.columns:
        zones = ["ALL"] + sorted(data[COL_ZONE].dropna().astype(str).unique())
        selected_zones = st.multiselect("🗺️ ZONES/LOCATIONS", zones, default=["ALL"])
        if "ALL" not in selected_zones:
            data = data[data[COL_ZONE].astype(str).isin(selected_zones)]

with filter_col3:
    if COL_PROJECT in data.columns:
        projects = ["ALL"] + sorted(data[COL_PROJECT].dropna().astype(str).unique())
        selected_projects = st.multiselect("📋 PROJECTS", projects, default=["ALL"])
        if "ALL" not in selected_projects:
            data = data[data[COL_PROJECT].astype(str).isin(selected_projects)]

with filter_col4:
    if COL_ITEM in data.columns:
        items = ["ALL"] + sorted(data[COL_ITEM].dropna().astype(str).unique())
        selected_items = st.multiselect("🔧 ITEM TYPES", items, default=["ALL"])
        if "ALL" not in selected_items:
            data = data[data[COL_ITEM].astype(str).isin(selected_items)]

unique_customers = len(data[COL_CUSTOMER].unique()) if COL_CUSTOMER in data.columns else 0
unique_zones = len(data[COL_ZONE].unique()) if COL_ZONE in data.columns else 0
unique_projects = len(data[COL_PROJECT].unique()) if COL_PROJECT in data.columns else 0

st.markdown(f"""
<div class="info-box">
📊 <strong>FILTERED DATASET SUMMARY:</strong> {len(data):,} RECORDS | {unique_customers} CUSTOMERS | {unique_zones} ZONES | {unique_projects} PROJECTS
</div>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------- KPIs ----------------------------
st.markdown('<div class="section-header">📈 KEY PERFORMANCE INDICATORS</div>', unsafe_allow_html=True)

total_delivered_weight = float(data[COL_WEIGHT].sum(skipna=True)) if COL_WEIGHT in data.columns else 0.0
total_delivered_qty = float(data[COL_QTY].sum(skipna=True)) if COL_QTY in data.columns else 0.0
total_contracted_weight = float(data[COL_CONTRACTED_WEIGHT].sum(skipna=True)) if COL_CONTRACTED_WEIGHT in data.columns else 0.0
total_contracted_qty = float(data[COL_CONTRACTED_QTY].sum(skipna=True)) if COL_CONTRACTED_QTY in data.columns else 0.0
avg_progress = float(data[COL_PROGRESS].mean(skipna=True) * 100) if COL_PROGRESS in data.columns else 0.0

weight_completion = (total_delivered_weight / total_contracted_weight * 100) if total_contracted_weight > 0 else 0.0
qty_completion = (total_delivered_qty / total_contracted_qty * 100) if total_contracted_qty > 0 else 0.0

active_trucks = 0
if truck_cols:
    truck_data = data[truck_cols].fillna(0)
    active_trucks = (truck_data > 0).any().sum()

kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
with kpi_col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🏗️ TOTAL DELIVERED WEIGHT</div>
        <div class="metric-value">{fmt_num(total_delivered_weight)} KG</div>
        <div class="metric-delta">{fmt_pct(weight_completion)} OF CONTRACTED</div>
    </div>
    """, unsafe_allow_html=True)
with kpi_col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">📦 TOTAL DELIVERED QUANTITY</div>
        <div class="metric-value">{fmt_num(total_delivered_qty, 0)}</div>
        <div class="metric-delta">{fmt_pct(qty_completion)} OF CONTRACTED</div>
    </div>
    """, unsafe_allow_html=True)
with kpi_col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">📊 AVERAGE PROGRESS</div>
        <div class="metric-value">{fmt_pct(avg_progress)}</div>
        <div class="metric-delta">ACROSS {len(data):,} ITEMS</div>
    </div>
    """, unsafe_allow_html=True)
with kpi_col4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🚛 ACTIVE TRUCKS</div>
        <div class="metric-value">{active_trucks}</div>
        <div class="metric-delta">OUT OF {len(truck_cols)} TOTAL</div>
    </div>
    """, unsafe_allow_html=True)

kpi_col5, kpi_col6, kpi_col7, kpi_col8 = st.columns(4)
with kpi_col5:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🏢 ACTIVE CUSTOMERS</div>
        <div class="metric-value">{unique_customers}</div>
        <div class="metric-delta">CUSTOMER PORTFOLIO</div>
    </div>
    """, unsafe_allow_html=True)
with kpi_col6:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">📋 ACTIVE PROJECTS</div>
        <div class="metric-value">{unique_projects}</div>
        <div class="metric-delta">PROJECT PIPELINE</div>
    </div>
    """, unsafe_allow_html=True)
with kpi_col7:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🗺️ ACTIVE ZONES</div>
        <div class="metric-value">{unique_zones}</div>
        <div class="metric-delta">GEOGRAPHIC COVERAGE</div>
    </div>
    """, unsafe_allow_html=True)
with kpi_col8:
    unique_items = len(data[COL_ITEM].unique()) if COL_ITEM in data.columns else 0
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🔧 ITEM TYPES</div>
        <div class="metric-value">{unique_items}</div>
        <div class="metric-delta">PRODUCT VARIETY</div>
    </div>
    """, unsafe_allow_html=True)

efficiency_col1, efficiency_col2, efficiency_col3, efficiency_col4 = st.columns(4)
avg_weight_per_delivery = (total_delivered_weight / len(data)) if len(data) > 0 else 0.0
total_balance_weight = float(data[COL_BALANCE_WEIGHT].sum(skipna=True)) if COL_BALANCE_WEIGHT in data.columns else 0.0
avg_truck_load = (total_delivered_qty / active_trucks) if active_trucks > 0 else 0.0
completion_projects = len(data[data[COL_PROGRESS] >= 0.95]) if COL_PROGRESS in data.columns else 0

with efficiency_col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">⚡ AVG WEIGHT/DELIVERY</div>
        <div class="metric-value">{fmt_num(avg_weight_per_delivery)} KG</div>
        <div class="metric-delta">DELIVERY EFFICIENCY</div>
    </div>
    """, unsafe_allow_html=True)
with efficiency_col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">⏳ REMAINING BALANCE</div>
        <div class="metric-value">{fmt_num(total_balance_weight)} KG</div>
        <div class="metric-delta">PENDING DELIVERY</div>
    </div>
    """, unsafe_allow_html=True)
with efficiency_col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🚛 AVG TRUCK LOAD</div>
        <div class="metric-value">{fmt_num(avg_truck_load, 1)}</div>
        <div class="metric-delta">UNITS PER TRUCK</div>
    </div>
    """, unsafe_allow_html=True)
with efficiency_col4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">✅ NEAR COMPLETION</div>
        <div class="metric-value">{completion_projects}</div>
        <div class="metric-delta">ITEMS >95% COMPLETE</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ---------------------------- Analytics Tabs ----------------------------
st.markdown('<div class="section-header">📊 ANALYTICS & INSIGHTS</div>', unsafe_allow_html=True)
tab1, tab2, tab3, tab4, tab5 = st.tabs(["🚛 TRUCK ANALYSIS", "🗺️ ZONE PERFORMANCE", "📋 PROJECT PROGRESS", "📈 TRENDS", "🔧 ITEM ANALYSIS"])

with tab1:
    st.subheader("TRUCK UTILIZATION ANALYSIS")
    if truck_cols:
        truck_data = data[truck_cols].fillna(0)
        truck_totals = truck_data.sum().sort_values(ascending=False)
        truck_utilization = truck_totals[truck_totals > 0]

        col1, col2 = st.columns(2)
        with col1:
            if len(truck_utilization) > 0:
                top_n = min(20, len(truck_utilization))
                top_trucks = truck_utilization.head(top_n).reset_index()
                top_trucks.columns = ["TRUCK", "TOTAL_QTY"]
                fig_trucks = px.bar(top_trucks, x="TRUCK", y="TOTAL_QTY",
                                    title=f"TOP {top_n} TRUCKS BY QUANTITY DELIVERED",
                                    color="TOTAL_QTY", color_continuous_scale="viridis")
                fig_trucks.update_layout(height=400, showlegend=False, title_font_size=16, title_font_color="#1f77b4")
                st.plotly_chart(fig_trucks, use_container_width=True)
        with col2:
            if len(truck_utilization) > 0:
                fig_hist = px.histogram(truck_utilization.values, nbins=20,
                                        title="TRUCK UTILIZATION DISTRIBUTION",
                                        labels={"value": "QUANTITY DELIVERED", "count": "NUMBER OF TRUCKS"})
                fig_hist.update_layout(height=400, title_font_size=16, title_font_color="#1f77b4")
                st.plotly_chart(fig_hist, use_container_width=True)

        st.subheader("TRUCK EFFICIENCY METRICS")
        col3, col4, col5 = st.columns(3)
        with col3:
            avg_utilization = truck_utilization.mean() if len(truck_utilization) > 0 else 0
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">📊 AVERAGE TRUCK LOAD</div>
                <div class="metric-value">{fmt_num(avg_utilization, 1)}</div>
            </div>
            """, unsafe_allow_html=True)
        with col4:
            max_utilization = truck_utilization.max() if len(truck_utilization) > 0 else 0
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">🔝 MAX TRUCK LOAD</div>
                <div class="metric-value">{fmt_num(max_utilization, 1)}</div>
            </div>
            """, unsafe_allow_html=True)
        with col5:
            utilization_rate = (len(truck_utilization) / len(truck_cols) * 100) if len(truck_cols) > 0 else 0
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">📈 UTILIZATION RATE</div>
                <div class="metric-value">{utilization_rate:.1f}%</div>
            </div>
            """, unsafe_allow_html=True)

with tab2:
    st.subheader("ZONE PERFORMANCE ANALYSIS")
    if COL_ZONE in data.columns and COL_WEIGHT in data.columns:
        zone_analysis = data.groupby(COL_ZONE, dropna=False).agg({
            COL_WEIGHT: 'sum',
            COL_QTY: 'sum',
            COL_PROGRESS: 'mean',
            COL_CUSTOMER: 'nunique'
        }).round(2)
        zone_analysis.columns = ['TOTAL_WEIGHT', 'TOTAL_QTY', 'AVG_PROGRESS', 'UNIQUE_CUSTOMERS']
        zone_analysis = zone_analysis.sort_values('TOTAL_WEIGHT', ascending=False)

        col1, col2 = st.columns(2)
        with col1:
            fig_zone_weight = px.bar(zone_analysis.reset_index(), x=COL_ZONE, y='TOTAL_WEIGHT',
                                     title="TOTAL DELIVERED WEIGHT BY ZONE",
                                     color='TOTAL_WEIGHT', color_continuous_scale="blues")
            fig_zone_weight.update_layout(height=400, title_font_size=16, title_font_color="#1f77b4")
            st.plotly_chart(fig_zone_weight, use_container_width=True)
        with col2:
            fig_zone_progress = px.bar(zone_analysis.reset_index(), x=COL_ZONE, y='AVG_PROGRESS',
                                       title="AVERAGE PROGRESS BY ZONE",
                                       color='AVG_PROGRESS', color_continuous_scale="greens")
            fig_zone_progress.update_layout(height=400, title_font_size=16, title_font_color="#1f77b4")
            st.plotly_chart(fig_zone_progress, use_container_width=True)

        st.subheader("ZONE PERFORMANCE SUMMARY")
        zone_analysis['AVG_PROGRESS'] = zone_analysis['AVG_PROGRESS'] * 100
        styled_zone_df = zone_analysis.style.format({
            'TOTAL_WEIGHT': '{:,.2f}',
            'TOTAL_QTY': '{:,.0f}',
            'AVG_PROGRESS': '{:.1f}%'
        })
        st.dataframe(styled_zone_df, height=300)

with tab3:
    st.subheader("PROJECT PROGRESS TRACKING")
    if COL_PROJECT in data.columns and COL_PROGRESS in data.columns:
        project_progress = data.groupby(COL_PROJECT, dropna=False).agg({
            COL_PROGRESS: 'mean',
            COL_WEIGHT: 'sum',
            COL_QTY: 'sum',
            COL_CUSTOMER: 'first',
            COL_ZONE: lambda x: ', '.join(pd.Series(x).astype(str).unique())
        }).round(3)
        project_progress.columns = ['AVG_PROGRESS', 'TOTAL_WEIGHT', 'TOTAL_QTY', 'CUSTOMER', 'ZONES']
        project_progress = project_progress.sort_values('AVG_PROGRESS', ascending=True)

        col1, col2 = st.columns(2)
        with col1:
            fig_project = px.bar(project_progress.reset_index().head(15),
                                 x='AVG_PROGRESS', y=COL_PROJECT, orientation='h',
                                 title="PROJECT PROGRESS STATUS (BOTTOM 15)",
                                 color='AVG_PROGRESS', color_continuous_scale="reds")
            fig_project.update_layout(height=500, title_font_size=16, title_font_color="#1f77b4")
            st.plotly_chart(fig_project, use_container_width=True)
        with col2:
            progress_ranges = pd.cut(project_progress['AVG_PROGRESS'],
                                     bins=[0, 0.25, 0.5, 0.75, 1.0],
                                     labels=['0-25%', '26-50%', '51-75%', '76-100%'])
            counts = progress_ranges.value_counts()
            fig_pie = px.pie(values=counts.values, names=counts.index,
                             title="PROJECT COMPLETION DISTRIBUTION")
            fig_pie.update_layout(height=400, title_font_size=16, title_font_color="#1f77b4")
            st.plotly_chart(fig_pie, use_container_width=True)

        st.subheader("PROJECT DETAILS")
        project_progress['AVG_PROGRESS'] = project_progress['AVG_PROGRESS'] * 100
        styled_project_df = project_progress.style.format({
            'AVG_PROGRESS': '{:.1f}%',
            'TOTAL_WEIGHT': '{:,.2f}',
            'TOTAL_QTY': '{:,.0f}'
        })
        st.dataframe(styled_project_df, height=300)

with tab4:
    st.subheader("DELIVERY TRENDS OVER TIME")
    if date_cols and active_date_col and active_date_col != "NONE":
        data_with_dates = data.dropna(subset=[active_date_col]).copy()
        data_with_dates['DATE'] = pd.to_datetime(data_with_dates[active_date_col])
        if len(data_with_dates) > 0:
            daily_stats = data_with_dates.groupby(data_with_dates['DATE'].dt.date).agg({
                COL_WEIGHT: 'sum',
                COL_QTY: 'sum'
            }).reset_index()
            daily_stats.columns = ['DATE', 'DAILY_WEIGHT', 'DAILY_QTY']

            col1, col2 = st.columns(2)
            with col1:
                fig_weight_trend = px.line(daily_stats, x='DATE', y='DAILY_WEIGHT',
                                           title="DAILY DELIVERED WEIGHT TREND", markers=True)
                fig_weight_trend.update_layout(height=400, title_font_size=16, title_font_color="#1f77b4")
                st.plotly_chart(fig_weight_trend, use_container_width=True)
            with col2:
                fig_qty_trend = px.line(daily_stats, x='DATE', y='DAILY_QTY',
                                        title="DAILY DELIVERED QUANTITY TREND", markers=True)
                fig_qty_trend.update_layout(height=400, title_font_size=16, title_font_color="#1f77b4")
                st.plotly_chart(fig_qty_trend, use_container_width=True)

            daily_stats['CUMULATIVE_WEIGHT'] = daily_stats['DAILY_WEIGHT'].cumsum()
            daily_stats['CUMULATIVE_QTY'] = daily_stats['DAILY_QTY'].cumsum()

            fig_cumulative = make_subplots(
                rows=2, cols=1,
                subplot_titles=('CUMULATIVE WEIGHT DELIVERED', 'CUMULATIVE QUANTITY DELIVERED'),
                vertical_spacing=0.1
            )
            fig_cumulative.add_trace(
                go.Scatter(x=daily_stats['DATE'], y=daily_stats['CUMULATIVE_WEIGHT'],
                           mode='lines+markers', name='WEIGHT (KG)'),
                row=1, col=1
            )
            fig_cumulative.add_trace(
                go.Scatter(x=daily_stats['DATE'], y=daily_stats['CUMULATIVE_QTY'],
                           mode='lines+markers', name='QUANTITY'),
                row=2, col=1
            )
            fig_cumulative.update_layout(height=600, title_text="CUMULATIVE DELIVERY PROGRESS",
                                         title_font_size=18, title_font_color="#1f77b4")
            st.plotly_chart(fig_cumulative, use_container_width=True)
        else:
            st.info("No valid dates found in the selected date column after filtering.")
    else:
        st.info("NO DATE COLUMNS FOUND FOR TREND ANALYSIS. PLEASE SELECT A DATE COLUMN IN THE FILTERS SECTION.")

with tab5:
    st.subheader("ITEM ANALYSIS")
    if COL_ITEM in data.columns:
        item_analysis = data.groupby(COL_ITEM, dropna=False).agg({
            COL_WEIGHT: 'sum',
            COL_QTY: 'sum',
            COL_PROGRESS: 'mean',
            COL_PROJECT: 'nunique'
        }).round(2)
        item_analysis.columns = ['TOTAL_WEIGHT', 'TOTAL_QTY', 'AVG_PROGRESS', 'PROJECT_COUNT']
        item_analysis = item_analysis.sort_values('TOTAL_WEIGHT', ascending=False)

        col1, col2 = st.columns(2)
        with col1:
            fig_items = px.bar(item_analysis.head(15).reset_index(),
                               x=COL_ITEM, y='TOTAL_WEIGHT',
                               title="TOP 15 ITEMS BY WEIGHT DELIVERED",
                               color='TOTAL_WEIGHT', color_continuous_scale="viridis")
            fig_items.update_xaxes(tickangle=45)
            fig_items.update_layout(height=500, title_font_size=16, title_font_color="#1f77b4")
            st.plotly_chart(fig_items, use_container_width=True)
        with col2:
            fig_item_progress = px.scatter(item_analysis.reset_index(),
                                           x='TOTAL_WEIGHT', y='AVG_PROGRESS',
                                           size='TOTAL_QTY', hover_name=COL_ITEM,
                                           title="ITEM WEIGHT VS PROGRESS (BUBBLE SIZE = QUANTITY)",
                                           color='PROJECT_COUNT', color_continuous_scale="plasma")
            fig_item_progress.update_layout(height=500, title_font_size=16, title_font_color="#1f77b4")
            st.plotly_chart(fig_item_progress, use_container_width=True)

        st.subheader("ITEM PERFORMANCE SUMMARY")
        item_analysis['AVG_PROGRESS'] = item_analysis['AVG_PROGRESS'] * 100
        styled_item_df = item_analysis.style.format({
            'TOTAL_WEIGHT': '{:,.2f}',
            'TOTAL_QTY': '{:,.0f}',
            'AVG_PROGRESS': '{:.1f}%'
        })
        st.dataframe(styled_item_df, height=300)

st.markdown("---")

# ---------------------------- Pivot Analysis ----------------------------
st.markdown('<div class="section-header">🔄 ADVANCED PIVOT ANALYSIS</div>', unsafe_allow_html=True)

pivot_col1, pivot_col2, pivot_col3, pivot_col4 = st.columns(4)
with pivot_col1:
    available_dims = [c for c in [COL_CUSTOMER, COL_PROJECT, COL_ZONE, COL_ITEM] if c in data.columns]
    pivot_rows = st.multiselect("📊 ROWS (GROUP BY)", available_dims, default=[COL_ZONE] if COL_ZONE in available_dims else [])
with pivot_col2:
    remaining_dims = [c for c in available_dims if c not in pivot_rows]
    pivot_cols = st.multiselect("📈 COLUMNS", remaining_dims, default=[])
with pivot_col3:
    value_options = [c for c in [COL_WEIGHT, COL_QTY, COL_PROGRESS] if c in data.columns]
    pivot_value = st.selectbox("📋 VALUES", value_options if value_options else [""], index=0)
with pivot_col4:
    pivot_agg = st.selectbox("🔢 AGGREGATION", ['SUM', 'MEAN', 'COUNT', 'MIN', 'MAX'], index=0)

if pivot_rows or pivot_cols:
    try:
        agg_func = pivot_agg.lower()
        if pivot_value and pivot_value in data.columns:
            pivot_table = pd.pivot_table(
                data,
                index=pivot_rows if pivot_rows else None,
                columns=pivot_cols if pivot_cols else None,
                values=pivot_value,
                aggfunc=agg_func,
                fill_value=0,
                dropna=False
            )
        else:
            pivot_table = pd.pivot_table(
                data,
                index=pivot_rows if pivot_rows else None,
                columns=pivot_cols if pivot_cols else None,
                aggfunc='size',
                fill_value=0
            )

        st.subheader("PIVOT TABLE RESULTS")
        if pivot_value in [COL_WEIGHT, COL_QTY]:
            formatted_pivot = pivot_table.style.format('{:,.2f}')
        elif pivot_value == COL_PROGRESS:
            formatted_pivot = pivot_table.style.format('{:.2%}')
        else:
            formatted_pivot = pivot_table.style.format('{:,.0f}')
        st.dataframe(formatted_pivot, height=400)

        pivot_csv = pivot_table.to_csv()
        st.download_button(
            "📥 DOWNLOAD PIVOT TABLE",
            pivot_csv,
            "FULAZ_PIVOT_ANALYSIS.CSV",
            "text/csv",
            key="pivot_download"
        )
    except Exception as e:
        st.error(f"ERROR CREATING PIVOT TABLE: {str(e)}")
else:
    st.info("SELECT ROWS OR COLUMNS TO CREATE A PIVOT TABLE")

st.markdown("---")

# ---------------------------- Export Section ----------------------------
st.markdown('<div class="section-header">📤 DATA EXPORT & REPORTING</div>', unsafe_allow_html=True)

export_col1, export_col2, export_col3, export_col4 = st.columns(4)
with export_col1:
    filtered_csv = data.to_csv(index=False)
    st.download_button(
        "📥 DOWNLOAD FILTERED DATA",
        filtered_csv,
        "FULAZ_FILTERED_DELIVERY_DATA.CSV",
        "text/csv",
        key="filtered_download",
        help="DOWNLOAD THE CURRENTLY FILTERED DATASET"
    )

with export_col2:
    if st.button("📊 GENERATE SUMMARY REPORT", key="summary_btn"):
        summary_data = {
            'METRIC': [
                'TOTAL RECORDS',
                'TOTAL DELIVERED WEIGHT (KG)',
                'TOTAL DELIVERED QUANTITY',
                'AVERAGE PROGRESS (%)',
                'WEIGHT COMPLETION RATE (%)',
                'QUANTITY COMPLETION RATE (%)',
                'ACTIVE CUSTOMERS',
                'ACTIVE PROJECTS',
                'ACTIVE ZONES',
                'ACTIVE TRUCKS',
                'AVERAGE WEIGHT PER DELIVERY'
            ],
            'VALUE': [
                f"{len(data):,}",
                f"{total_delivered_weight:,.2f}",
                f"{total_delivered_qty:,.0f}",
                f"{avg_progress:.1f}%",
                f"{weight_completion:.1f}%",
                f"{qty_completion:.1f}%",
                f"{unique_customers}",
                f"{unique_projects}",
                f"{unique_zones}",
                f"{active_trucks}",
                f"{avg_weight_per_delivery:,.2f}"
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_csv = summary_df.to_csv(index=False)
        st.download_button(
            "📥 DOWNLOAD SUMMARY",
            summary_csv,
            "FULAZ_DELIVERY_SUMMARY_REPORT.CSV",
            "text/csv",
            key="summary_download"
        )

with export_col3:
    if truck_cols and st.button("🚛 EXPORT TRUCK DATA", key="truck_btn"):
        tdata = data[truck_cols].fillna(0)
        truck_summary = pd.DataFrame({
            'TRUCK': truck_cols,
            'TOTAL_DELIVERIES': [tdata[col].sum() for col in truck_cols],
            'ACTIVE_DAYS': [(tdata[col] > 0).sum() for col in truck_cols],
            'UTILIZATION_RATE': [(tdata[col] > 0).mean() * 100 for col in truck_cols]
        })
        truck_summary = truck_summary.sort_values('TOTAL_DELIVERIES', ascending=False)
        truck_csv = truck_summary.to_csv(index=False)
        st.download_button(
            "📥 DOWNLOAD TRUCK ANALYSIS",
            truck_csv,
            "FULAZ_TRUCK_UTILIZATION_ANALYSIS.CSV",
            "text/csv",
            key="truck_download"
        )

with export_col4:
    if COL_ZONE in data.columns and st.button("🗺️ EXPORT ZONE DATA", key="zone_btn"):
        zone_export = data.groupby(COL_ZONE, dropna=False).agg({
            COL_WEIGHT: ['sum', 'mean'],
            COL_QTY: ['sum', 'mean'],
            COL_PROGRESS: 'mean',
            COL_CUSTOMER: 'nunique',
            COL_PROJECT: 'nunique'
        }).round(2)
        zone_export.columns = ['TOTAL_WEIGHT', 'AVG_WEIGHT', 'TOTAL_QTY', 'AVG_QTY', 'AVG_PROGRESS', 'CUSTOMERS', 'PROJECTS']
        zone_csv = zone_export.to_csv()
        st.download_button(
            "📥 DOWNLOAD ZONE ANALYSIS",
            zone_csv,
            "FULAZ_ZONE_PERFORMANCE_ANALYSIS.CSV",
            "text/csv",
            key="zone_download"
        )

# ---------------------------- Data Preview ----------------------------
with st.expander("🔍 FILTERED DATA PREVIEW", expanded=False):
    st.markdown("### CURRENT FILTERED DATASET")
    if len(data) > 0:
        preview_data = data.head(100)
        st.dataframe(preview_data, height=400, use_container_width=True)
        st.markdown(f"**SHOWING FIRST 100 ROWS OF {len(data):,} TOTAL FILTERED RECORDS**")
    else:
        st.warning("NO DATA AVAILABLE AFTER APPLYING FILTERS")

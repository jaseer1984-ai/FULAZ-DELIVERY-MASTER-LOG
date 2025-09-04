# -*- coding: utf-8 -*-
# Enhanced Excel → Streamlit Dashboard for FULAZ Delivery MasterLog
# - Comprehensive KPI section with progress tracking
# - Advanced filtering with customer and project filters
# - Interactive time-series analysis
# - Detailed truck utilization analytics
# - Geographic zone analysis
# - Project completion tracking
# - Export capabilities with multiple formats
# Author: AI Assistant | Enhanced for FULAZ data structure

import io
from typing import List, Optional
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ---------------------------- Config & Helpers ----------------------------
st.set_page_config(
    page_title="FULAZ Delivery MasterLog Dashboard", 
    layout="wide", 
    page_icon="🏗️",
    initial_sidebar_state="expanded"
)

def fmt_num(x, dec=2):
    try:
        return f"{float(x):,.{dec}f}"
    except Exception:
        return "--"

def fmt_pct(x):
    try:
        return f"{float(x):.1f}%"
    except Exception:
        return "--"

@st.cache_data(show_spinner=False)
def load_excel(file, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Enhanced Excel loader with better error handling"""
    def read_xls(excel_file: pd.ExcelFile, sheet_nm: Optional[str]) -> pd.DataFrame:
        target_sheet = sheet_nm or excel_file.sheet_names[0]
        for hdr in (2, 1, 0):  # Try header at rows 3, 2, then 1
            try:
                df_try = excel_file.parse(target_sheet, header=hdr)
                df_try.columns = [str(c).strip() for c in df_try.columns]
                # Check for meaningful columns
                good_cols = sum(not str(c).lower().startswith("unnamed") for c in df_try.columns)
                if good_cols >= 5:
                    return df_try
            except Exception:
                continue
        # Fallback
        df_fallback = excel_file.parse(target_sheet)
        df_fallback.columns = [str(c).strip() for c in df_fallback.columns]
        return df_fallback

    # Handle different file types
    if isinstance(file, (str, bytes)):
        path_str = file if isinstance(file, str) else ""
        engine = "xlrd" if path_str.lower().endswith(".xls") else "openpyxl"
        xls = pd.ExcelFile(file, engine=engine)
        return read_xls(xls, sheet_name)

    # Handle uploaded file
    file_bytes = file.read()
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="xlrd")
    return read_xls(xls, sheet_name)

def detect_truck_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if str(c).strip().lower().startswith("truck")]

def numericify(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def find_date_cols(df: pd.DataFrame) -> List[str]:
    """Enhanced date column detection"""
    named = [c for c in df.columns if "date" in str(c).lower()]
    dtyped = list(df.select_dtypes(include=["datetime64[ns]", "datetime64[ns, UTC]"]).columns)
    
    # Check for columns that might contain dates
    for c in df.columns:
        if c not in named and c not in dtyped:
            s = df[c].dropna()
            if len(s) > 0:
                try:
                    # Try to parse a sample of values
                    sample_size = min(20, len(s))
                    sample = s.head(sample_size)
                    parsed = pd.to_datetime(sample, errors="coerce")
                    if parsed.notna().sum() >= sample_size * 0.5:
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

# ---------------------------- Custom CSS ----------------------------
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .kpi-container {
        background: linear-gradient(90deg, #f0f2f6 0%, #ffffff 100%);
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
        margin: 0.5rem 0;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #2c3e50;
        margin: 1rem 0;
        border-bottom: 2px solid #1f77b4;
        padding-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------- Sidebar: Upload + Configuration ----------------------------
st.sidebar.markdown("### 📄 File Upload")
uploaded = st.sidebar.file_uploader("Upload FULAZ Delivery MasterLog", type=["xlsx", "xls"])

if not uploaded:
    st.markdown('<div class="main-header">🏗️ FULAZ Delivery MasterLog Dashboard</div>', unsafe_allow_html=True)
    st.info("👈 Please upload your **FULAZ Delivery MasterLog** Excel file to get started.")
    st.markdown("""
    ### Expected File Structure:
    - **Customer Name**, **Project Name**, **Zone/Location**
    - **Item Name**, **Delivered Weight**, **Delivered Qty**
    - **Progress %**, **Truck columns** (Truck1, Truck2, ...)
    - **Date columns** (automatically detected)
    """)
    st.stop()

# Sheet selection
try:
    _peek = pd.ExcelFile(io.BytesIO(uploaded.getvalue()), engine="openpyxl")
    sheet_options = _peek.sheet_names
except Exception:
    _peek = pd.ExcelFile(io.BytesIO(uploaded.getvalue()), engine="xlrd")
    sheet_options = _peek.sheet_names

sheet_name = st.sidebar.selectbox("📋 Select Sheet", options=sheet_options, index=0)

# ---------------------------- Load & Prepare Data ----------------------------
with st.spinner("Loading and processing data..."):
    data = load_excel(uploaded, sheet_name=sheet_name)
    data.columns = [c.strip() for c in data.columns]

# Define key columns based on FULAZ structure
COL_CUSTOMER = "Customer Name"
COL_PROJECT = "Project Name" 
COL_PROJECT_NUM = "Project Number"
COL_ZONE = "Zone / Location"
COL_ITEM = "Item Name"
COL_ITEM_DESC = "Item Description"
COL_WEIGHT = "Delivered Weight"
COL_QTY = "Delivered Qty"
COL_PROGRESS = "Progress %"
COL_CONTRACTED_WEIGHT = "Contracted Weight"
COL_CONTRACTED_QTY = "Contracted Qty"
COL_BALANCE_WEIGHT = "Balance Weight"
COL_BALANCE_QTY = "Balance Qty"

# Process dates, trucks, and numerics
date_cols = find_date_cols(data)
data = coerce_dates(data, date_cols)

truck_cols = detect_truck_cols(data)
numeric_cols_to_cast = [
    c for c in [COL_WEIGHT, COL_QTY, COL_PROGRESS, COL_CONTRACTED_WEIGHT, 
               COL_CONTRACTED_QTY, COL_BALANCE_WEIGHT, COL_BALANCE_QTY] 
    if c in data.columns
] + truck_cols

data = numericify(data, numeric_cols_to_cast)

# Clean progress percentage (convert to 0-1 range if needed)
if COL_PROGRESS in data.columns:
    # If progress values are > 1, assume they're in percentage form
    max_progress = data[COL_PROGRESS].max()
    if not pd.isna(max_progress) and max_progress > 1:
        data[COL_PROGRESS] = data[COL_PROGRESS] / 100

# ---------------------------- Main Dashboard ----------------------------
st.markdown('<div class="main-header">🏗️ FULAZ Delivery MasterLog Dashboard</div>', unsafe_allow_html=True)

# ---------------------------- Enhanced Filters Section ----------------------------
st.markdown('<div class="section-header">🔍 Smart Filters</div>', unsafe_allow_html=True)

# Create filter columns
filter_col1, filter_col2, filter_col3 = st.columns(3)

with filter_col1:
    # Date filter
    active_date_col = None
    if date_cols:
        active_date_col = st.selectbox("📅 Date Column", options=date_cols, index=0)
        if active_date_col:
            dseries = pd.to_datetime(data[active_date_col], errors="coerce")
            valid_dates = dseries.dropna()
            if len(valid_dates) > 0:
                min_d, max_d = valid_dates.min().date(), valid_dates.max().date()
                date_range = st.date_input(
                    "Date Range",
                    value=(min_d, max_d),
                    min_value=min_d,
                    max_value=max_d
                )
                if isinstance(date_range, tuple) and len(date_range) == 2:
                    start_d, end_d = date_range
                    mask_date = (dseries.dt.date >= start_d) & (dseries.dt.date <= end_d)
                    data = data[mask_date]

with filter_col2:
    # Customer filter
    if COL_CUSTOMER in data.columns:
        customers = ["All"] + sorted(data[COL_CUSTOMER].dropna().astype(str).unique())
        selected_customers = st.multiselect("🏢 Customer", customers, default=["All"])
        if "All" not in selected_customers:
            data = data[data[COL_CUSTOMER].astype(str).isin(selected_customers)]
    
    # Zone filter
    if COL_ZONE in data.columns:
        zones = ["All"] + sorted(data[COL_ZONE].dropna().astype(str).unique())
        selected_zones = st.multiselect("🗺️ Zone/Location", zones, default=["All"])
        if "All" not in selected_zones:
            data = data[data[COL_ZONE].astype(str).isin(selected_zones)]

with filter_col3:
    # Project filter
    if COL_PROJECT in data.columns:
        projects = ["All"] + sorted(data[COL_PROJECT].dropna().astype(str).unique())
        selected_projects = st.multiselect("📋 Project", projects, default=["All"])
        if "All" not in selected_projects:
            data = data[data[COL_PROJECT].astype(str).isin(selected_projects)]
    
    # Item filter
    if COL_ITEM in data.columns:
        items = ["All"] + sorted(data[COL_ITEM].dropna().astype(str).unique())
        selected_items = st.multiselect("🔧 Item Type", items, default=["All"])
        if "All" not in selected_items:
            data = data[data[COL_ITEM].astype(str).isin(selected_items)]

# Filter summary
st.info(f"📊 **Filtered Dataset**: {len(data):,} records from {len(data[COL_CUSTOMER].unique()) if COL_CUSTOMER in data.columns else 'N/A'} customers across {len(data[COL_ZONE].unique()) if COL_ZONE in data.columns else 'N/A'} zones")

# ---------------------------- Enhanced KPIs Section ----------------------------
st.markdown('<div class="section-header">📈 Key Performance Indicators</div>', unsafe_allow_html=True)

# Calculate comprehensive KPIs
total_delivered_weight = float(data[COL_WEIGHT].sum(skipna=True)) if COL_WEIGHT in data.columns else 0
total_delivered_qty = float(data[COL_QTY].sum(skipna=True)) if COL_QTY in data.columns else 0
total_contracted_weight = float(data[COL_CONTRACTED_WEIGHT].sum(skipna=True)) if COL_CONTRACTED_WEIGHT in data.columns else 0
total_contracted_qty = float(data[COL_CONTRACTED_QTY].sum(skipna=True)) if COL_CONTRACTED_QTY in data.columns else 0
avg_progress = float(data[COL_PROGRESS].mean(skipna=True) * 100) if COL_PROGRESS in data.columns else 0

# Completion rates
weight_completion = (total_delivered_weight / total_contracted_weight * 100) if total_contracted_weight > 0 else 0
qty_completion = (total_delivered_qty / total_contracted_qty * 100) if total_contracted_qty > 0 else 0

# Active trucks calculation
active_trucks = 0
if truck_cols:
    truck_data = data[truck_cols].fillna(0)
    active_trucks = (truck_data > 0).any().sum()

# Display KPIs in columns
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

with kpi_col1:
    st.markdown('<div class="kpi-container">', unsafe_allow_html=True)
    st.metric(
        "🏗️ Delivered Weight",
        f"{fmt_num(total_delivered_weight)} kg",
        delta=f"{fmt_pct(weight_completion)} of contracted"
    )
    st.markdown('</div>', unsafe_allow_html=True)

with kpi_col2:
    st.markdown('<div class="kpi-container">', unsafe_allow_html=True)
    st.metric(
        "📦 Delivered Quantity", 
        fmt_num(total_delivered_qty, 0),
        delta=f"{fmt_pct(qty_completion)} of contracted"
    )
    st.markdown('</div>', unsafe_allow_html=True)

with kpi_col3:
    st.markdown('<div class="kpi-container">', unsafe_allow_html=True)
    st.metric(
        "📊 Average Progress",
        f"{fmt_pct(avg_progress)}",
        delta=f"Across {len(data)} items"
    )
    st.markdown('</div>', unsafe_allow_html=True)

with kpi_col4:
    st.markdown('<div class="kpi-container">', unsafe_allow_html=True)
    st.metric(
        "🚛 Active Trucks",
        f"{active_trucks}",
        delta=f"Out of {len(truck_cols)} total"
    )
    st.markdown('</div>', unsafe_allow_html=True)

# Additional KPIs row
kpi_col5, kpi_col6, kpi_col7, kpi_col8 = st.columns(4)

with kpi_col5:
    unique_customers = len(data[COL_CUSTOMER].unique()) if COL_CUSTOMER in data.columns else 0
    st.metric("🏢 Active Customers", f"{unique_customers}")

with kpi_col6:
    unique_projects = len(data[COL_PROJECT].unique()) if COL_PROJECT in data.columns else 0
    st.metric("📋 Active Projects", f"{unique_projects}")

with kpi_col7:
    unique_zones = len(data[COL_ZONE].unique()) if COL_ZONE in data.columns else 0
    st.metric("🗺️ Active Zones", f"{unique_zones}")

with kpi_col8:
    unique_items = len(data[COL_ITEM].unique()) if COL_ITEM in data.columns else 0
    st.metric("🔧 Item Types", f"{unique_items}")

st.markdown("---")

# ---------------------------- Enhanced Visualizations ----------------------------
st.markdown('<div class="section-header">📊 Analytics & Insights</div>', unsafe_allow_html=True)

# Create tabs for different analysis views
tab1, tab2, tab3, tab4, tab5 = st.tabs(["🚛 Truck Analysis", "🗺️ Zone Performance", "📋 Project Progress", "📈 Trends", "🔧 Item Analysis"])

with tab1:
    st.subheader("Truck Utilization Analysis")
    
    if truck_cols:
        # Calculate truck utilization
        truck_data = data[truck_cols].fillna(0)
        truck_totals = truck_data.sum().sort_values(ascending=False)
        truck_utilization = truck_totals[truck_totals > 0]
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Top trucks chart
            if len(truck_utilization) > 0:
                top_n = min(20, len(truck_utilization))
                top_trucks = truck_utilization.head(top_n).reset_index()
                top_trucks.columns = ["Truck", "Total_Qty"]
                
                fig_trucks = px.bar(
                    top_trucks, 
                    x="Truck", 
                    y="Total_Qty",
                    title=f"Top {top_n} Trucks by Quantity Delivered",
                    color="Total_Qty",
                    color_continuous_scale="viridis"
                )
                fig_trucks.update_layout(height=400, showlegend=False)
                st.plotly_chart(fig_trucks, use_container_width=True)
        
        with col2:
            # Truck utilization distribution
            if len(truck_utilization) > 0:
                fig_hist = px.histogram(
                    truck_utilization.values,
                    nbins=20,
                    title="Truck Utilization Distribution",
                    labels={"value": "Quantity Delivered", "count": "Number of Trucks"}
                )
                fig_hist.update_layout(height=400)
                st.plotly_chart(fig_hist, use_container_width=True)
        
        # Truck efficiency metrics
        st.subheader("Truck Efficiency Metrics")
        col3, col4, col5 = st.columns(3)
        
        with col3:
            avg_utilization = truck_utilization.mean() if len(truck_utilization) > 0 else 0
            st.metric("Average Truck Load", fmt_num(avg_utilization, 1))
        
        with col4:
            max_utilization = truck_utilization.max() if len(truck_utilization) > 0 else 0
            st.metric("Max Truck Load", fmt_num(max_utilization, 1))
        
        with col5:
            utilization_rate = (len(truck_utilization) / len(truck_cols) * 100) if len(truck_cols) > 0 else 0
            st.metric("Truck Utilization Rate", f"{utilization_rate:.1f}%")

with tab2:
    st.subheader("Zone Performance Analysis")
    
    if COL_ZONE in data.columns and COL_WEIGHT in data.columns:
        zone_analysis = data.groupby(COL_ZONE, dropna=False).agg({
            COL_WEIGHT: 'sum',
            COL_QTY: 'sum',
            COL_PROGRESS: 'mean',
            COL_CUSTOMER: 'nunique'
        }).round(2)
        
        zone_analysis.columns = ['Total_Weight', 'Total_Qty', 'Avg_Progress', 'Unique_Customers']
        zone_analysis = zone_analysis.sort_values('Total_Weight', ascending=False)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Weight by zone
            fig_zone_weight = px.bar(
                zone_analysis.reset_index(),
                x=COL_ZONE,
                y='Total_Weight',
                title="Total Delivered Weight by Zone",
                color='Total_Weight',
                color_continuous_scale="blues"
            )
            fig_zone_weight.update_layout(height=400)
            st.plotly_chart(fig_zone_weight, use_container_width=True)
        
        with col2:
            # Progress by zone
            fig_zone_progress = px.bar(
                zone_analysis.reset_index(),
                x=COL_ZONE,
                y='Avg_Progress',
                title="Average Progress by Zone",
                color='Avg_Progress',
                color_continuous_scale="greens"
            )
            fig_zone_progress.update_layout(height=400)
            st.plotly_chart(fig_zone_progress, use_container_width=True)
        
        # Zone performance table
        st.subheader("Zone Performance Summary")
        zone_analysis['Avg_Progress'] = zone_analysis['Avg_Progress'] * 100
        st.dataframe(zone_analysis.style.format({
            'Total_Weight': '{:,.2f}',
            'Total_Qty': '{:,.0f}',
            'Avg_Progress': '{:.1f}%'
        }))

with tab3:
    st.subheader("Project Progress Tracking")
    
    if COL_PROJECT in data.columns and COL_PROGRESS in data.columns:
        project_progress = data.groupby(COL_PROJECT, dropna=False).agg({
            COL_PROGRESS: 'mean',
            COL_WEIGHT: 'sum',
            COL_QTY: 'sum',
            COL_CUSTOMER: 'first',
            COL_ZONE: lambda x: ', '.join(x.unique())
        }).round(3)
        
        project_progress.columns = ['Avg_Progress', 'Total_Weight', 'Total_Qty', 'Customer', 'Zones']
        project_progress = project_progress.sort_values('Avg_Progress', ascending=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Project progress chart
            fig_project = px.bar(
                project_progress.reset_index().head(15),
                x='Avg_Progress',
                y=COL_PROJECT,
                orientation='h',
                title="Project Progress Status (Bottom 15)",
                color='Avg_Progress',
                color_continuous_scale="reds"
            )
            fig_project.update_layout(height=500)
            st.plotly_chart(fig_project, use_container_width=True)
        
        with col2:
            # Project completion status
            progress_ranges = pd.cut(
                project_progress['Avg_Progress'], 
                bins=[0, 0.25, 0.5, 0.75, 1.0], 
                labels=['0-25%', '26-50%', '51-75%', '76-100%']
            )
            progress_counts = progress_ranges.value_counts()
            
            fig_pie = px.pie(
                values=progress_counts.values,
                names=progress_counts.index,
                title="Project Completion Distribution"
            )
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)
        
        # Project details table
        st.subheader("Project Details")
        project_progress['Avg_Progress'] = project_progress['Avg_Progress'] * 100
        st.dataframe(
            project_progress.style.format({
                'Avg_Progress': '{:.1f}%',
                'Total_Weight': '{:,.2f}',
                'Total_Qty': '{:,.0f}'
            }),
            height=300
        )

with tab4:
    st.subheader("Delivery Trends Over Time")
    
    if date_cols and active_date_col:
        # Create time series data
        data_with_dates = data.dropna(subset=[active_date_col])
        data_with_dates['Date'] = pd.to_datetime(data_with_dates[active_date_col])
        
        if len(data_with_dates) > 0:
            # Daily aggregation
            daily_stats = data_with_dates.groupby(data_with_dates['Date'].dt.date).agg({
                COL_WEIGHT: 'sum',
                COL_QTY: 'sum'
            }).reset_index()
            daily_stats.columns = ['Date', 'Daily_Weight', 'Daily_Qty']
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Weight trend
                fig_weight_trend = px.line(
                    daily_stats,
                    x='Date',
                    y='Daily_Weight',
                    title="Daily Delivered Weight Trend",
                    markers=True
                )
                fig_weight_trend.update_layout(height=400)
                st.plotly_chart(fig_weight_trend, use_container_width=True)
            
            with col2:
                # Quantity trend
                fig_qty_trend = px.line(
                    daily_stats,
                    x='Date',
                    y='Daily_Qty',
                    title="Daily Delivered Quantity Trend",
                    markers=True,
                    color_discrete_sequence=['orange']
                )
                fig_qty_trend.update_layout(height=400)
                st.plotly_chart(fig_qty_trend, use_container_width=True)
            
            # Cumulative progress
            daily_stats['Cumulative_Weight'] = daily_stats['Daily_Weight'].cumsum()
            daily_stats['Cumulative_Qty'] = daily_stats['Daily_Qty'].cumsum()
            
            fig_cumulative = make_subplots(
                rows=2, cols=1,
                subplot_titles=('Cumulative Weight Delivered', 'Cumulative Quantity Delivered'),
                vertical_spacing=0.1
            )
            
            fig_cumulative.add_trace(
                go.Scatter(
                    x=daily_stats['Date'],
                    y=daily_stats['Cumulative_Weight'],
                    mode='lines+markers',
                    name='Weight (kg)',
                    line=dict(color='blue')
                ),
                row=1, col=1
            )
            
            fig_cumulative.add_trace(
                go.Scatter(
                    x=daily_stats['Date'],
                    y=daily_stats['Cumulative_Qty'],
                    mode='lines+markers',
                    name='Quantity',
                    line=dict(color='red')
                ),
                row=2, col=1
            )
            
            fig_cumulative.update_layout(height=600, title_text="Cumulative Delivery Progress")
            st.plotly_chart(fig_cumulative, use_container_width=True)
    else:
        st.info("No date columns found for trend analysis.")

with tab5:
    st.subheader("Item Analysis")
    
    if COL_ITEM in data.columns:
        item_analysis = data.groupby(COL_ITEM, dropna=False).agg({
            COL_WEIGHT: 'sum',
            COL_QTY: 'sum',
            COL_PROGRESS: 'mean',
            COL_PROJECT: 'nunique'
        }).round(2)
        
        item_analysis.columns = ['Total_Weight', 'Total_Qty', 'Avg_Progress', 'Project_Count']
        item_analysis = item_analysis.sort_values('Total_Weight', ascending=False)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Top items by weight
            fig_items = px.bar(
                item_analysis.head(15).reset_index(),
                x=COL_ITEM,
                y='Total_Weight',
                title="Top 15 Items by Weight Delivered",
                color='Total_Weight',
                color_continuous_scale="viridis"
            )
            fig_items.update_xaxes(tickangle=45)
            fig_items.update_layout(height=500)
            st.plotly_chart(fig_items, use_container_width=True)
        
        with col2:
            # Item progress analysis
            fig_item_progress = px.scatter(
                item_analysis.reset_index(),
                x='Total_Weight',
                y='Avg_Progress',
                size='Total_Qty',
                hover_name=COL_ITEM,
                title="Item Weight vs Progress (Bubble size = Quantity)",
                color='Project_Count',
                color_continuous_scale="plasma"
            )
            fig_item_progress.update_layout(height=500)
            st.plotly_chart(fig_item_progress, use_container_width=True)
        
        # Item performance table
        st.subheader("Item Performance Summary")
        item_analysis['Avg_Progress'] = item_analysis['Avg_Progress'] * 100
        st.dataframe(
            item_analysis.style.format({
                'Total_Weight': '{:,.2f}',
                'Total_Qty': '{:,.0f}',
                'Avg_Progress': '{:.1f}%'
            }),
            height=300
        )

st.markdown("---")

# ---------------------------- Advanced Pivot Analysis ----------------------------
st.markdown('<div class="section-header">🔄 Advanced Pivot Analysis</div>', unsafe_allow_html=True)

# Enhanced pivot builder
pivot_col1, pivot_col2, pivot_col3 = st.columns(3)

with pivot_col1:
    available_dims = [c for c in [COL_CUSTOMER, COL_PROJECT, COL_ZONE, COL_ITEM] if c in data.columns]
    pivot_rows = st.multiselect("📊 Rows (Group by)", available_dims, default=[COL_ZONE] if COL_ZONE in available_dims else [])

with pivot_col2:
    remaining_dims = [c for c in available_dims if c not in pivot_rows]
    pivot_cols = st.multiselect("📈 Columns", remaining_dims, default=[])

with pivot_col3:
    value_options = [c for c in [COL_WEIGHT, COL_QTY, COL_PROGRESS] if c in data.columns]
    pivot_value = st.selectbox("📋 Values", value_options, index=0 if value_options else 0)
    pivot_agg = st.selectbox("🔢 Aggregation", ['sum', 'mean', 'count', 'min', 'max'], index=0)

# Create pivot table
if pivot_rows or pivot_cols:
    try:
        if pivot_value and pivot_value in data.columns:
            pivot_table = pd.pivot_table(
                data,
                index=pivot_rows if pivot_rows else None,
                columns=pivot_cols if pivot_cols else None,
                values=pivot_value,
                aggfunc=pivot_agg,
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
        
        st.subheader("Pivot Table Results")
        st.dataframe(pivot_table, height=400)
        
        # Download pivot table
        pivot_csv = pivot_table.to_csv()
        st.download_button(
            "📥 Download Pivot Table",
            pivot_csv,
            "fulaz_pivot_analysis.csv",
            "text/csv"
        )
        
    except Exception as e:
        st.error(f"Error creating pivot table: {str(e)}")

st.markdown("---")

# ---------------------------- Data Export Section ----------------------------
st.markdown('<div class="section-header">📤 Data Export & Summary</div>', unsafe_allow_html=True)

export_col1, export_col2, export_col3 = st.columns(3)

with export_col1:
    # Filtered data download
    filtered_csv = data.to_csv(index=False)
    st.download_button(
        "📥 Download Filtered Data (CSV)",
        filtered_csv,
        "fulaz_filtered_delivery_data.csv",
        "text/csv"
    )

with export_col2:
    # Summary report
    if st.button("📊 Generate Summary Report"):
        summary_data = {
            'Metric': [
                'Total Records',
                'Total Delivered Weight (kg)',
                'Total Delivered Quantity',
                'Average Progress (%)',
                'Active Customers',
                'Active Projects', 
                'Active Zones',
                'Active Trucks'
            ],
            'Value': [
                f"{len(data):,}",
                f"{total_delivered_weight:,.2f}",
                f"{total_delivered_qty:,.0f}",
                f"{avg_progress:.1f}%",
                f"{unique_customers}",
                f"{unique_projects}",
                f"{unique_zones}",
                f"{active_trucks}"
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_csv = summary_df.to_csv(index=False)
        st.download_button(
            "📥 Download Summary Report",
            summary_csv,
            "fulaz_delivery_summary.csv",
            "text/csv"
        )

with export_col3:
    # Truck utilization export
    if truck_cols and st.button("🚛 Export Truck Data"):
        truck_data = data[truck_cols].fillna(0)
        truck_summary = pd.DataFrame({
            'Truck': truck_cols,
            'Total_Deliveries': [truck_data[col].sum() for col in truck_cols],
            'Active_Days': [(truck_data[col] > 0).sum() for col in truck_cols]
        })
        truck_csv = truck_summary.to_csv(index=False)
        st.download_button(
            "📥 Download Truck Analysis",
            truck_csv,
            "fulaz_truck_utilization.csv",
            "text/csv"
        )

# ---------------------------- Data Preview ----------------------------
with st.expander("🔍 Filtered Data Preview", expanded=False):
    st.dataframe(data, height=400, use_container_width=True)
    st.caption(f"Showing {len(data):,} records after applying filters")

# ---------------------------- Footer ----------------------------
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <strong>🏗️ FULAZ Delivery MasterLog Dashboard</strong><br>
    Enhanced analytics for construction delivery management<br>
    <em>Built with Streamlit & Plotly | Last updated: September 2025</em>
</div>
""", unsafe_allow_html=True)

# Performance tips
with st.expander("💡 Dashboard Tips", expanded=False):
    st.markdown("""
    ### Getting the Most from Your Dashboard:
    
    **🔍 Filtering:**
    - Use date ranges to focus on specific time periods
    - Combine multiple filters for detailed analysis
    - "All" selections show complete dataset
    
    **📊 Analytics Tabs:**
    - **Truck Analysis**: Monitor fleet utilization and efficiency
    - **Zone Performance**: Compare regional delivery performance  
    - **Project Progress**: Track individual project completion
    - **Trends**: Analyze delivery patterns over time
    - **Item Analysis**: Understand material-specific metrics
    
    **📤 Export Options:**
    - Download filtered data for external analysis
    - Generate summary reports for stakeholders
    - Export pivot tables for custom analysis
    
    **⚡ Performance:**
    - Dashboard updates automatically when filters change
    - Large datasets may take a few seconds to process
    - Use date filters to improve loading speed
    """)

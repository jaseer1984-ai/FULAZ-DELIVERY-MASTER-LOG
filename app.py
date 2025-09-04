# -*- coding: utf-8 -*-
# FULAZ Delivery MasterLog Dashboard - MINIMAL WORKING VERSION
# NO xlrd dependency - uses only openpyxl for .xlsx files
# All caps formatting for professional presentation

import io
import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------------------- Config ----------------------------
st.set_page_config(
    page_title="FULAZ DELIVERY MASTERLOG DASHBOARD", 
    layout="wide", 
    page_icon="🏗️"
)

def format_number(x, decimals=2):
    """Format numbers with commas and uppercase styling"""
    try:
        return f"{float(x):,.{decimals}f}".upper()
    except:
        return "--"

def load_xlsx_file(uploaded_file, sheet_name=None):
    """Load Excel file using only openpyxl - no caching to avoid conflicts"""
    try:
        # Read file bytes
        file_bytes = uploaded_file.getvalue()
        
        # Try different header rows
        for header_row in [2, 1, 0]:  # Rows 3, 2, 1
            try:
                df = pd.read_excel(
                    io.BytesIO(file_bytes),
                    sheet_name=sheet_name or 0,
                    header=header_row,
                    engine='openpyxl'
                )
                
                # Convert columns to uppercase
                df.columns = [str(col).strip().upper() for col in df.columns]
                
                # Check for meaningful data
                if len(df) > 0 and len(df.columns) > 5:
                    return df
                    
            except Exception:
                continue
        
        # If headers fail, try without header
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name or 0,
            header=None,
            engine='openpyxl'
        )
        
        if len(df) > 0:
            # Use first row as headers
            df.columns = [str(col).strip().upper() for col in df.iloc[0]]
            df = df.drop(df.index[0]).reset_index(drop=True)
            return df
            
        return pd.DataFrame()
        
    except Exception as e:
        st.error(f"ERROR LOADING FILE: {str(e)}")
        return pd.DataFrame()

# ---------------------------- CSS Styling ----------------------------
st.markdown("""
<style>
    .main-title {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        text-transform: uppercase;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(145deg, #ffffff, #f0f2f6);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
        margin: 0.5rem 0;
        text-align: center;
    }
    .metric-title {
        font-size: 0.9rem;
        color: #666;
        font-weight: bold;
        text-transform: uppercase;
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: bold;
        color: #1f77b4;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------- Main Interface ----------------------------
st.markdown('<div class="main-title">🏗️ FULAZ DELIVERY MASTERLOG DASHBOARD</div>', unsafe_allow_html=True)

# File upload
uploaded = st.file_uploader(
    "📄 UPLOAD FULAZ DELIVERY MASTERLOG (.xlsx only)", 
    type=["xlsx"],
    help="Only .xlsx Excel files are supported"
)

if not uploaded:
    st.info("👆 PLEASE UPLOAD YOUR FULAZ DELIVERY MASTERLOG EXCEL FILE")
    st.markdown("""
    ### REQUIREMENTS:
    - File must be in .xlsx format
    - Should contain columns: CUSTOMER NAME, ZONE/LOCATION, ITEM NAME
    - Should have DELIVERED WEIGHT, DELIVERED QTY, PROGRESS %
    - May contain TRUCK columns (TRUCK1, TRUCK2, etc.)
    """)
    st.stop()

# Load data
with st.spinner("LOADING DATA..."):
    try:
        # Get sheet names
        file_bytes = uploaded.getvalue()
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        sheet_names = excel_file.sheet_names
        excel_file.close()
        
        # Sheet selection
        if len(sheet_names) > 1:
            sheet_name = st.selectbox("SELECT SHEET:", sheet_names)
        else:
            sheet_name = sheet_names[0]
        
        # Load the data
        data = load_xlsx_file(uploaded, sheet_name)
        
        if data.empty:
            st.error("❌ NO DATA LOADED - PLEASE CHECK YOUR FILE FORMAT")
            st.stop()
            
        st.success(f"✅ DATA LOADED: {len(data):,} ROWS, {len(data.columns)} COLUMNS")
        
    except Exception as e:
        st.error(f"❌ ERROR: {str(e)}")
        st.info("💡 Please ensure your file is a valid .xlsx Excel file")
        st.stop()

# ---------------------------- Column Detection ----------------------------
# Define expected columns (in uppercase)
COL_CUSTOMER = "CUSTOMER NAME"
COL_ZONE = "ZONE / LOCATION" 
COL_ITEM = "ITEM NAME"
COL_WEIGHT = "DELIVERED WEIGHT"
COL_QTY = "DELIVERED QTY"
COL_PROGRESS = "PROGRESS %"
COL_PROJECT = "PROJECT NAME"

# Detect truck columns
truck_cols = [col for col in data.columns if col.startswith("TRUCK")]

# Convert numeric columns
numeric_cols = [COL_WEIGHT, COL_QTY, COL_PROGRESS] + truck_cols
for col in numeric_cols:
    if col in data.columns:
        data[col] = pd.to_numeric(data[col], errors='coerce')

# Fix progress percentage (convert to 0-1 if needed)
if COL_PROGRESS in data.columns:
    max_progress = data[COL_PROGRESS].max()
    if max_progress > 1:
        data[COL_PROGRESS] = data[COL_PROGRESS] / 100

# ---------------------------- Filters ----------------------------
st.markdown("### 🔍 FILTERS")

col1, col2, col3 = st.columns(3)

with col1:
    if COL_CUSTOMER in data.columns:
        customers = ["ALL"] + sorted(data[COL_CUSTOMER].dropna().unique())
        selected_customers = st.multiselect("🏢 CUSTOMERS:", customers, default=["ALL"])
        if "ALL" not in selected_customers:
            data = data[data[COL_CUSTOMER].isin(selected_customers)]

with col2:
    if COL_ZONE in data.columns:
        zones = ["ALL"] + sorted(data[COL_ZONE].dropna().unique())
        selected_zones = st.multiselect("🗺️ ZONES:", zones, default=["ALL"])
        if "ALL" not in selected_zones:
            data = data[data[COL_ZONE].isin(selected_zones)]

with col3:
    if COL_PROJECT in data.columns:
        projects = ["ALL"] + sorted(data[COL_PROJECT].dropna().unique())
        selected_projects = st.multiselect("📋 PROJECTS:", projects, default=["ALL"])
        if "ALL" not in selected_projects:
            data = data[data[COL_PROJECT].isin(selected_projects)]

st.info(f"📊 FILTERED DATA: {len(data):,} RECORDS")

# ---------------------------- KPIs ----------------------------
st.markdown("### 📈 KEY PERFORMANCE INDICATORS")

# Calculate KPIs
total_weight = data[COL_WEIGHT].sum() if COL_WEIGHT in data.columns else 0
total_qty = data[COL_QTY].sum() if COL_QTY in data.columns else 0
avg_progress = data[COL_PROGRESS].mean() * 100 if COL_PROGRESS in data.columns else 0
active_trucks = len([col for col in truck_cols if data[col].sum() > 0]) if truck_cols else 0

# Display KPIs
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

with kpi_col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🏗️ TOTAL WEIGHT</div>
        <div class="metric-value">{format_number(total_weight)} KG</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">📦 TOTAL QUANTITY</div>
        <div class="metric-value">{format_number(total_qty, 0)}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">📊 AVG PROGRESS</div>
        <div class="metric-value">{format_number(avg_progress, 1)}%</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_col4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🚛 ACTIVE TRUCKS</div>
        <div class="metric-value">{active_trucks}</div>
    </div>
    """, unsafe_allow_html=True)

# ---------------------------- Charts ----------------------------
st.markdown("### 📊 ANALYTICS")

# Chart tabs
tab1, tab2, tab3 = st.tabs(["🚛 TRUCKS", "🗺️ ZONES", "🔧 ITEMS"])

with tab1:
    if truck_cols:
        truck_data = data[truck_cols].sum().sort_values(ascending=False)
        truck_active = truck_data[truck_data > 0].head(15)
        
        if len(truck_active) > 0:
            fig_trucks = px.bar(
                x=truck_active.index,
                y=truck_active.values,
                title="TOP 15 ACTIVE TRUCKS",
                labels={'x': 'TRUCK', 'y': 'TOTAL DELIVERIES'}
            )
            fig_trucks.update_layout(height=400)
            st.plotly_chart(fig_trucks, use_container_width=True)
    else:
        st.info("NO TRUCK COLUMNS FOUND")

with tab2:
    if COL_ZONE in data.columns and COL_WEIGHT in data.columns:
        zone_weight = data.groupby(COL_ZONE)[COL_WEIGHT].sum().sort_values(ascending=False)
        
        if len(zone_weight) > 0:
            fig_zones = px.bar(
                x=zone_weight.index,
                y=zone_weight.values,
                title="DELIVERED WEIGHT BY ZONE",
                labels={'x': 'ZONE', 'y': 'WEIGHT (KG)'}
            )
            fig_zones.update_layout(height=400)
            st.plotly_chart(fig_zones, use_container_width=True)

with tab3:
    if COL_ITEM in data.columns and COL_WEIGHT in data.columns:
        item_weight = data.groupby(COL_ITEM)[COL_WEIGHT].sum().sort_values(ascending=False).head(10)
        
        if len(item_weight) > 0:
            fig_items = px.bar(
                x=item_weight.index,
                y=item_weight.values,
                title="TOP 10 ITEMS BY WEIGHT",
                labels={'x': 'ITEM', 'y': 'WEIGHT (KG)'}
            )
            fig_items.update_layout(height=400)
            fig_items.update_xaxes(tickangle=45)
            st.plotly_chart(fig_items, use_container_width=True)

# ---------------------------- Data Export ----------------------------
st.markdown("### 📤 EXPORT DATA")

col1, col2 = st.columns(2)

with col1:
    # Export filtered data
    csv_data = data.to_csv(index=False)
    st.download_button(
        "📥 DOWNLOAD FILTERED DATA",
        csv_data,
        "FULAZ_FILTERED_DATA.csv",
        "text/csv"
    )

with col2:
    # Export summary
    summary = pd.DataFrame({
        'METRIC': ['TOTAL RECORDS', 'TOTAL WEIGHT (KG)', 'TOTAL QUANTITY', 'AVERAGE PROGRESS (%)', 'ACTIVE TRUCKS'],
        'VALUE': [len(data), f"{total_weight:,.2f}", f"{total_qty:,.0f}", f"{avg_progress:.1f}%", active_trucks]
    })
    summary_csv = summary.to_csv(index=False)
    st.download_button(
        "📥 DOWNLOAD SUMMARY",
        summary_csv,
        "FULAZ_SUMMARY_REPORT.csv",
        "text/csv"
    )

# ---------------------------- Data Preview ----------------------------
with st.expander("🔍 DATA PREVIEW"):
    st.dataframe(data.head(100), height=300)
    st.caption(f"SHOWING FIRST 100 ROWS OF {len(data):,} TOTAL RECORDS")

# ---------------------------- Footer ----------------------------
st.markdown("---")
st.markdown("""
<div style='text-align: center; padding: 20px; background: #f8f9fa; border-radius: 10px;'>
    <strong>🏗️ FULAZ DELIVERY MASTERLOG DASHBOARD</strong><br>
    <em>MINIMAL VERSION - PROFESSIONAL ANALYTICS PLATFORM</em>
</div>
""", unsafe_allow_html=True)

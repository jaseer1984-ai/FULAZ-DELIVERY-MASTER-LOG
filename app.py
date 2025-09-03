# -*- coding: utf-8 -*-
# Excel → Streamlit Dashboard for Delivery MasterLog (Trucks, Weight, Progress)
# - Robust Excel loader (.xlsx via openpyxl, .xls via xlrd)
# - Auto-detects header row (tries row 3, then 2, then 1 in Excel terms)
# - KPIs: Delivered Weight, Delivered Qty, Avg Progress %
# - Charts: Top Trucks, Weight by Zone/Location, Weight by Item Name
# - Pivot builder + filtered download
# Author: AI Assistant | Last updated: 2025-09-03

import io
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# ---------------------------- Config & Helpers ----------------------------
st.set_page_config(page_title="Delivery MasterLog — Dashboard", layout="wide", page_icon="📦")

def fmt_num(x, dec=2):
    try:
        return f"{float(x):,.{dec}f}"
    except Exception:
        return "--"

@st.cache_data(show_spinner=False)
def load_excel(file, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    Robust Excel loader:
    - Supports UploadedFile or local path
    - Uses openpyxl for .xlsx, xlrd for legacy .xls
    - Attempts headers at Excel rows: 3 (header=2), then 2 (header=1), then 1 (header=0)
    - Returns a dataframe with stripped column names
    """
    def read_xls(excel_file: pd.ExcelFile, sheet_name: Optional[str]) -> pd.DataFrame:
        target_sheet = sheet_name or excel_file.sheet_names[0]
        # Try header row 3 → 2 → 1 (0-indexed: 2,1,0)
        for hdr in (2, 1, 0):
            try:
                df_try = excel_file.parse(target_sheet, header=hdr)
                df_try.columns = [str(c).strip() for c in df_try.columns]
                # Accept if at least a few non-"Unnamed" columns are present
                good_cols = sum(not str(c).lower().startswith("unnamed") for c in df_try.columns)
                if good_cols >= 3:
                    return df_try
            except Exception:
                pass
        # Fallback: default parse
        df_fallback = excel_file.parse(target_sheet)
        df_fallback.columns = [str(c).strip() for c in df_fallback.columns]
        return df_fallback

    # If path-like string/bytes
    if isinstance(file, (str, bytes)):
        path_str = file if isinstance(file, str) else ""
        if path_str.lower().endswith(".xls"):
            xls = pd.ExcelFile(file, engine="xlrd")
        else:
            xls = pd.ExcelFile(file, engine="openpyxl")
        return read_xls(xls, sheet_name)

    # Streamlit UploadedFile
    file_bytes = file.read()
    # Try .xlsx first, then .xls
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


# ---------------------------- Sidebar Controls ----------------------------
st.sidebar.title("⚙️ Controls")

uploaded = st.sidebar.file_uploader("Upload Delivery MasterLog (.xlsx / .xls)", type=["xlsx", "xls"])

if not uploaded:
    st.sidebar.info("Upload your **Delivery MasterLog in FULAZ.xlsx** file to get started.")
    st.stop()

# Optional: choose sheet
try:
    # Peek sheet names safely
    _peek = pd.ExcelFile(io.BytesIO(uploaded.getvalue()), engine="openpyxl")
    sheet_options = _peek.sheet_names
except Exception:
    _peek = pd.ExcelFile(io.BytesIO(uploaded.getvalue()), engine="xlrd")
    sheet_options = _peek.sheet_names

sheet_name = st.sidebar.selectbox("Sheet", options=sheet_options, index=0)

# ---------------------------- Load & Prepare ----------------------------
data = load_excel(uploaded, sheet_name=sheet_name)
data.columns = [c.strip() for c in data.columns]

# Heuristic important columns (names observed in your file)
COL_ZONE = "Zone / Location"
COL_ITEM = "Item Name"
COL_WEIGHT = "Delivered Weight"
COL_QTY = "Delivered Qty"
COL_PROGRESS = "Progress %"

truck_cols = detect_truck_cols(data)
numeric_cols_to_cast = [COL_WEIGHT, COL_QTY, COL_PROGRESS] + truck_cols
data = numericify(data, numeric_cols_to_cast)

# ---------------------------- KPIs ----------------------------
total_weight = float(data[COL_WEIGHT].sum(skipna=True)) if COL_WEIGHT in data.columns else np.nan
total_qty = float(data[COL_QTY].sum(skipna=True)) if COL_QTY in data.columns else np.nan
avg_progress_pct = float(data[COL_PROGRESS].mean(skipna=True) * 100.0) if COL_PROGRESS in data.columns else np.nan

st.title("📦 Delivery MasterLog — Dashboard")
st.caption("Auto-reads your Excel, computes KPIs, and shows trucks & delivery insights. Designed for **FULAZ** log structure.")

c1, c2, c3 = st.columns(3)
c1.metric("Total Delivered Weight (kg)", "N/A" if np.isnan(total_weight) else fmt_num(total_weight, 2))
c2.metric("Total Delivered Qty", "N/A" if np.isnan(total_qty) else fmt_num(total_qty, 0))
c3.metric("Average Progress %", "N/A" if np.isnan(avg_progress_pct) else fmt_num(avg_progress_pct, 1) + "%")

st.markdown("---")

# ---------------------------- Filters ----------------------------
# Build filter lists dynamically from likely categorical columns
cat_candidates = []
for col in [COL_ZONE, COL_ITEM]:
    if col in data.columns and data[col].dtype == object:
        cat_candidates.append(col)

filters = {}
for col in cat_candidates:
    vals = ["(All)"] + sorted([v for v in data[col].dropna().astype(str).unique()][:2000])
    sel = st.multiselect(f"Filter — {col}", vals, default="(All)")
    if "(All)" not in sel:
        filters[col] = sel

fdata = data.copy()
for col, sel in filters.items():
    fdata = fdata[fdata[col].astype(str).isin(sel)]

# ---------------------------- Charts ----------------------------
# Chart 1: Top Trucks by Quantity (summing across Truck columns)
if truck_cols:
    truck_totals = fdata[truck_cols].sum().sort_values(ascending=False)
    top_n = st.slider("Top N Trucks", min_value=5, max_value=min(30, len(truck_totals)), value=min(15, len(truck_totals)))
    top_trucks = truck_totals.head(top_n).reset_index()
    top_trucks.columns = ["Truck", "Qty"]

    fig_trucks = px.bar(top_trucks, x="Truck", y="Qty", title=f"Top {top_n} Trucks by Quantity")
    fig_trucks.update_layout(margin=dict(l=4, r=4, t=50, b=0), height=360, showlegend=False)
    st.plotly_chart(fig_trucks, use_container_width=True, config={"displayModeBar": False})
else:
    st.info("No columns starting with **Truck** were found to build the Top Trucks chart.")

# Chart 2: Delivered Weight by Zone / Location
if COL_ZONE in fdata.columns and COL_WEIGHT in fdata.columns:
    zone_totals = (
        fdata.groupby(COL_ZONE, dropna=False)[COL_WEIGHT]
        .sum()
        .reset_index()
        .sort_values(COL_WEIGHT, ascending=False)
        .head(15)
    )
    fig_zone = px.bar(zone_totals, x=COL_ZONE, y=COL_WEIGHT, title="Delivered Weight by Zone / Location (Top 15)")
    fig_zone.update_layout(margin=dict(l=4, r=4, t=50, b=0), height=360, showlegend=False)
    st.plotly_chart(fig_zone, use_container_width=True, config={"displayModeBar": False})

# Chart 3: Delivered Weight by Item Name
if COL_ITEM in fdata.columns and COL_WEIGHT in fdata.columns:
    item_totals = (
        fdata.groupby(COL_ITEM, dropna=False)[COL_WEIGHT]
        .sum()
        .reset_index()
        .sort_values(COL_WEIGHT, ascending=False)
        .head(15)
    )
    fig_item = px.bar(item_totals, x=COL_ITEM, y=COL_WEIGHT, title="Delivered Weight by Item Name (Top 15)")
    fig_item.update_layout(margin=dict(l=4, r=4, t=50, b=0), height=360, showlegend=False)
    st.plotly_chart(fig_item, use_container_width=True, config={"displayModeBar": False})

st.markdown("---")

# ---------------------------- Pivot Builder ----------------------------
st.subheader("Pivot")
# Find additional categorical columns for pivoting (avoid truck columns)
more_cat_cols = [
    c for c in fdata.columns
    if c not in truck_cols and fdata[c].dtype == object and fdata[c].nunique(dropna=True) <= 100
]

left, right = st.columns([2, 1])
with left:
    rows = st.multiselect("Rows", options=more_cat_cols, default=([COL_ZONE] if COL_ZONE in more_cat_cols else []))
    cols = st.multiselect("Columns", options=[c for c in more_cat_cols if c not in rows], default=[])
with right:
    value_col = st.selectbox("Value", options=[COL_WEIGHT, COL_QTY] + ([COL_PROGRESS] if COL_PROGRESS in fdata.columns else []))
    aggfunc = st.selectbox("Aggregation", options=["sum", "mean", "count"], index=0)

vals = None if value_col not in fdata.columns else value_col
if vals and aggfunc in ["sum", "mean"] and pd.api.types.is_numeric_dtype(fdata[vals]):
    pvt = pd.pivot_table(
        fdata, index=rows or None, columns=cols or None, values=vals,
        aggfunc=aggfunc, fill_value=0, dropna=False, observed=False
    )
else:
    pvt = pd.pivot_table(
        fdata, index=rows or None, columns=cols or None, values=vals if vals in fdata.columns else None,
        aggfunc="count", fill_value=0, dropna=False, observed=False
    )
st.dataframe(pvt, use_container_width=True)

# ---------------------------- Data & Download ----------------------------
with st.expander("Filtered data (preview)"):
    st.dataframe(fdata, use_container_width=True, height=320)

csv = fdata.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download filtered CSV", data=csv, file_name="filtered_delivery_data.csv", mime="text/csv")

st.caption("Tip: Use the filters above to narrow by Zone/Location or Item. The app sums all **Truck** columns to rank trucks by quantity.")

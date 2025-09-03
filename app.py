
# -*- coding: utf-8 -*-
# Streamlit Excel Dashboard — Auto-adaptive (wide-to-long, filters, KPIs, charts)
# Author: AI Assistant
# Last updated: 2025-09-03

import io
from datetime import datetime
from typing import List

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# ---------------------- App Config ----------------------
st.set_page_config(page_title="Excel → Streamlit Dashboard", layout="wide", page_icon="📊")

@st.cache_data(show_spinner=False)
def load_excel(file, sheet_name=None) -> pd.DataFrame:
    if isinstance(file, (str, bytes)):
        # path or bytes
        xls = pd.ExcelFile(file)
    else:
        # UploadedFile
        file_bytes = file.read()
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = sheet_name or xls.sheet_names[0]
    df = xls.parse(sheet)
    # Normalize columns
    df.columns = [str(c).strip() for c in df.columns]
    return df

def parseable_date(s: pd.Series) -> bool:
    sample = s.dropna().astype(str).head(60)
    if sample.empty:
        return False
    parsed = pd.to_datetime(sample, errors="coerce", dayfirst=True, infer_datetime_format=True)
    return parsed.notna().mean() >= 0.6

def detect_structure(df: pd.DataFrame):
    """Return a tuple (mode, info) where mode in {'tidy','wide'}.
    - tidy: there is a date column already
    - wide: many date-like COLUMN NAMES; we'll melt them
    """
    # 1) Look for an existing date column
    date_cols = [c for c in df.columns if parseable_date(df[c])]
    # 2) Or date-like column NAMES (e.g., '2025-08-20', '25/8/2025', etc.)
    name_dates = []
    for c in df.columns:
        try:
            pd.to_datetime([c], errors="raise", dayfirst=True)
            name_dates.append(c)
        except Exception:
            pass

    mode = "tidy" if date_cols else ("wide" if len(name_dates) >= max(3, int(0.05*len(df.columns))) else "tidy")
    info = {"date_cols": date_cols, "name_date_cols": name_dates}
    return mode, info

def to_tidy(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure we return a DataFrame with a 'Date' column and one or more numeric value columns."""
    mode, info = detect_structure(df)
    tidy = df.copy()
    if mode == "tidy":
        # Choose the best date column or let user pick later
        date_col = None
        for c in info["date_cols"]:
            if parseable_date(df[c]):
                date_col = c
                break
        if date_col:
            tidy["Date"] = pd.to_datetime(tidy[date_col], errors="coerce", dayfirst=True, infer_datetime_format=True)
        else:
            # No strong date column; create a row index-based pseudo-date
            tidy["Date"] = pd.Series(pd.NaT, index=tidy.index)
    else:
        # WIDE: melt date-like column NAMES
        date_cols = info["name_date_cols"]
        id_vars = [c for c in df.columns if c not in date_cols]
        long_df = df.melt(id_vars=id_vars, value_vars=date_cols, var_name="Date", value_name="Value")
        # Convert Date column from header name
        long_df["Date"] = pd.to_datetime(long_df["Date"], errors="coerce", dayfirst=True, infer_datetime_format=True)
        tidy = long_df

    # Coerce numerics
    for c in tidy.columns:
        if c not in ["Date"]:
            # try convert; ignore if becomes all NaN
            converted = pd.to_numeric(tidy[c], errors="coerce")
            if converted.notna().sum() >= 5:
                tidy[c] = converted

    return tidy

def numeric_candidates(df: pd.DataFrame) -> List[str]:
    nums = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c]) and c.lower() != "year"]
    # dedupe and keep a reasonable number
    return nums[:20]

def cat_candidates(df: pd.DataFrame) -> List[str]:
    cats = [c for c in df.columns if df[c].dtype == object and df[c].nunique(dropna=True) <= 100]
    return cats[:20]

# ---------------------- Sidebar ----------------------
st.sidebar.title("⚙️ Controls")
data_src = st.sidebar.radio("Data source", ["Use sample upload", "Use built-in demo"], index=0)

uploaded = st.sidebar.file_uploader("Upload Excel (.xlsx, .xls)", type=["xlsx", "xls"])

default_path = None
if data_src == "Use sample upload" and uploaded is not None:
    raw = load_excel(uploaded)
elif data_src == "Use sample upload" and uploaded is None:
    st.sidebar.info("Upload your Excel file to begin.")
    raw = None
else:
    # Built-in demo (small synthetic)
    demo = pd.DataFrame({
        "Date": pd.date_range("2025-01-01", periods=30, freq="D"),
        "Branch": np.random.choice(["JED","RYD","DAM"], size=30),
        "Truck": np.random.choice(["T-01","T-02","T-03"], size=30),
        "Deliveries": np.random.randint(20, 80, size=30),
        "Weight_Tons": np.random.uniform(5, 20, size=30).round(2),
        "OnTime": np.random.choice(["Yes","No"], size=30, p=[0.8,0.2]),
    })
    raw = demo

if raw is None or raw.empty:
    st.stop()

tidy = to_tidy(raw)

st.sidebar.markdown("---")
# Assist selection if multiple candidates
date_col = "Date"
if "Date" not in tidy.columns:
    # fallback: let user pick any column
    possible = [c for c in tidy.columns if parseable_date(tidy[c])]
    date_col = st.sidebar.selectbox("Select date column", options=possible or tidy.columns.tolist())
else:
    # confirm
    date_col = "Date"

# Ensure Date dtype
tidy[date_col] = pd.to_datetime(tidy[date_col], errors="coerce")

# Candidate fields
num_cols = numeric_candidates(tidy)
cat_cols = cat_candidates(tidy)

# User selections
metric = st.sidebar.selectbox("Primary metric", options=(["Value"] if "Value" in tidy.columns else []) + num_cols or tidy.columns.tolist())
group = st.sidebar.selectbox("Group by (category)", options=(["Branch","Truck","Status"] + cat_cols))
st.sidebar.markdown("---")
date_min = pd.to_datetime(tidy[date_col].min())
date_max = pd.to_datetime(tidy[date_col].max())
if pd.isna(date_min) or pd.isna(date_max):
    date_min = pd.Timestamp("2025-01-01")
    date_max = pd.Timestamp("2025-12-31")
date_range = st.sidebar.date_input("Date range", value=(date_min.date(), date_max.date()))
if isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = [pd.to_datetime(d) for d in date_range]
else:
    start, end = date_min, date_max

# Optional filters for up to three categorical columns
f1 = st.sidebar.selectbox("Filter 1 (optional)", options=["(none)"] + cat_cols, index=0)
f2 = st.sidebar.selectbox("Filter 2 (optional)", options=["(none)"] + cat_cols, index=0)
f3 = st.sidebar.selectbox("Filter 3 (optional)", options=["(none)"] + cat_cols, index=0)

# Prepare filtered data
df = tidy.copy()
df = df[(df[date_col] >= start) & (df[date_col] <= end)]
for f in (f1, f2, f3):
    if f and f != "(none)":
        vals = ["(All)"] + sorted(list(df[f].dropna().unique()))
        sel = st.sidebar.multiselect(f"Select {f}", options=vals, default="(All)")
        if "(All)" not in sel:
            df = df[df[f].isin(sel)]

st.sidebar.markdown("---")
dl = st.sidebar.toggle("Show raw data", value=False)

# ---------------------- Header ----------------------
st.title("📦 Excel → Streamlit Dashboard")
st.caption("Auto-detects dates and converts wide sheets (date-as-header) into a tidy model for filtering and charts.")

# ---------------------- KPIs ----------------------
kpi1 = float(df[metric].sum(skipna=True)) if metric in df.columns and pd.api.types.is_numeric_dtype(df[metric]) else float(df.shape[0])
kpi2 = float(df[metric].mean(skipna=True)) if metric in df.columns and pd.api.types.is_numeric_dtype(df[metric]) else float('nan')
kpi3 = int(df[group].nunique(dropna=True)) if group in df.columns else 0

col1, col2, col3 = st.columns(3)
col1.metric(f"Total {metric}", f"{kpi1:,.2f}" if not np.isnan(kpi1) else "--")
col2.metric(f"Avg {metric}", f"{kpi2:,.2f}" if not np.isnan(kpi2) else "--")
col3.metric(f"Unique {group}", f"{kpi3:,}")

# ---------------------- Charts ----------------------
if date_col in df.columns and metric in df.columns and pd.api.types.is_numeric_dtype(df[metric]):
    trend = df.groupby(pd.Grouper(key=date_col, freq="D"), dropna=False)[metric].sum().reset_index()
    fig = px.line(trend, x=date_col, y=metric, title=f"{metric} over time")
    fig.update_layout(margin=dict(l=8,r=8,t=40,b=8), height=360, showlegend=False)
    st.plotly_chart(fig, use_container_width=True, theme="streamlit")

# Top by category
if group in df.columns:
    if metric in df.columns and pd.api.types.is_numeric_dtype(df[metric]):
        agg = df.groupby(group, dropna=False)[metric].sum().reset_index()
        ycol = metric
    else:
        agg = df[group].value_counts(dropna=False).reset_index().rename(columns={"count": "Records", "index": group})
        ycol = "Records"
    topn = agg.sort_values(ycol, ascending=False).head(15)
    fig2 = px.bar(topn, x=group, y=ycol, title=f"Top {group}")
    fig2.update_layout(margin=dict(l=8,r=8,t=40,b=8), height=360)
    st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

# Pivot table
st.subheader("Pivot")
left, right = st.columns([2,1])
with left:
    rows = st.multiselect("Rows", options=[c for c in cat_cols if c != group], default=([group] if group in cat_cols else []))
    cols = st.multiselect("Columns", options=[c for c in cat_cols if c not in rows], default=[])
with right:
    aggfunc = st.selectbox("Aggregation", options=["sum","mean","count"], index=0)

if metric in df.columns and aggfunc in ["sum","mean"] and pd.api.types.is_numeric_dtype(df[metric]):
    pvt = pd.pivot_table(df, index=rows or None, columns=cols or None, values=metric, aggfunc=aggfunc, fill_value=0, dropna=False, observed=False)
else:
    # count by default
    pvt = pd.pivot_table(df, index=rows or None, columns=cols or None, values=metric if metric in df.columns else None, aggfunc="count", fill_value=0, dropna=False, observed=False)

st.dataframe(pvt, use_container_width=True)

# ---------------------- Data & Download ----------------------
st.subheader("Filtered data")
st.dataframe(df, use_container_width=True, height=320) if dl else st.caption("Toggle 'Show raw data' in the sidebar to preview the full table.")
csv = df.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download filtered CSV", data=csv, file_name="filtered_data.csv", mime="text/csv")

st.caption("Tip: If your Excel has dates as column headers (e.g., '2025-08-20', '25/8/2025'), the app will automatically convert it to a tidy time series. Use the sidebar to pick your metric and grouping fields.")

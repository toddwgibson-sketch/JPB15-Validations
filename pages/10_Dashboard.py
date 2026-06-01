import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
from pathlib import Path

from utils.auth import require_login
require_login()

st.set_page_config(
    page_title="Validation Error Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================== BEAUTIFUL C-SUITE STYLING ======================
st.markdown("""
<style>
    .main-header {
        font-size: 2.4rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #475569;
        margin-bottom: 1.8rem;
    }
    .section-header {
        font-size: 1.35rem;
        font-weight: 600;
        color: #1e2937;
        margin-top: 1.8rem;
        margin-bottom: 0.6rem;
    }
    .kpi-card {
        background: white;
        border-radius: 16px;
        padding: 1.2rem 1.4rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.06);
        border: 1px solid #e2e8f0;
    }
    .metric-value {
        font-size: 1.9rem;
        font-weight: 700;
        color: #0f172a;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">📊 Validation Error Dashboard</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Executive View • Cross-Hall • Real-time Error Intelligence</p>', unsafe_allow_html=True)

# ====================== DATA ======================
DATA_FILE = Path(__file__).parent.parent / "data" / "validation_error_log.xlsx"

@st.cache_data(ttl=30)
def load_data():
    if not DATA_FILE.exists():
        return pd.DataFrame(columns=[
            "timestamp", "hall", "rack_type", "building", 
            "error_category", "count", "source_file", "processed_by"
        ])
    df = pd.read_excel(DATA_FILE)
    df['timestamp'] = pd.to_datetime(df['timestamp'])
    return df

df = load_data()

if df.empty:
    abs_path = DATA_FILE.resolve()
    st.warning("No error data logged yet. Process validation files using the tools in this app to populate this dashboard.")
    st.info("The central error log lives here (inside your GitHub repo):")
    st.code(str(abs_path))
    st.caption("Click the test button on the SYD20 QFAB page (or any other validation page) to create the first rows.")
    st.stop()

# ====================== FILTERS ======================
st.sidebar.header("Filters")

# Date range
min_date = df['timestamp'].min().date()
max_date = df['timestamp'].max().date()

date_range = st.sidebar.date_input(
    "Date Range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

selected_halls = st.sidebar.multiselect(
    "Hall", 
    options=sorted(df['hall'].unique()), 
    default=sorted(df['hall'].unique())
)

selected_types = st.sidebar.multiselect(
    "Rack Type", 
    options=sorted(df['rack_type'].unique()), 
    default=sorted(df['rack_type'].unique())
)

# Apply filters
filtered = df[
    (df['timestamp'].dt.date >= date_range[0]) &
    (df['timestamp'].dt.date <= date_range[1]) &
    (df['hall'].isin(selected_halls)) &
    (df['rack_type'].isin(selected_types))
]

# ====================== EXECUTIVE KPI CARDS ======================
st.markdown('<div class="section-header">Executive Snapshot</div>', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

total_errors = int(filtered['count'].sum())
unique_racks = filtered['building'].nunique()
active_halls = filtered['hall'].nunique()
avg_errors_per_rack = round(total_errors / unique_racks, 1) if unique_racks > 0 else 0

with col1:
    st.metric("Total Errors Logged", f"{total_errors:,}")
with col2:
    st.metric("Unique Racks/Bldgs", unique_racks)
with col3:
    st.metric("Halls Active", active_halls)
with col4:
    st.metric("Avg Errors per Rack", avg_errors_per_rack)

st.divider()

# ====================== MAIN CHARTS ======================
c1, c2 = st.columns(2)

with c1:
    st.markdown("**Errors by Hall**")
    hall_sum = filtered.groupby('hall')['count'].sum().reset_index()
    fig = px.bar(hall_sum, x='hall', y='count', color='hall', title=None)
    fig.update_layout(showlegend=False, height=380)
    st.plotly_chart(fig, use_container_width=True)

with c2:
    st.markdown("**Errors by Rack Type**")
    type_sum = filtered.groupby('rack_type')['count'].sum().reset_index()
    fig2 = px.pie(type_sum, names='rack_type', values='count', hole=0.45)
    fig2.update_layout(height=380)
    st.plotly_chart(fig2, use_container_width=True)

# ====================== TOP PROBLEMATIC RACKS ======================
st.markdown('<div class="section-header">Top Problematic Racks</div>', unsafe_allow_html=True)

rack_summary = (
    filtered.groupby(['building', 'hall', 'rack_type'])['count']
    .sum()
    .reset_index()
    .sort_values('count', ascending=False)
    .head(12)
)

fig3 = px.bar(
    rack_summary,
    x='building',
    y='count',
    color='hall',
    hover_data=['rack_type'],
    title="Highest Error Racks (Last Period)"
)
fig3.update_layout(xaxis_tickangle=-35, height=420)
st.plotly_chart(fig3, use_container_width=True)

# ====================== DETAILED TABLE ======================
st.markdown('<div class="section-header">Detailed Error Log</div>', unsafe_allow_html=True)

st.dataframe(
    filtered.sort_values('timestamp', ascending=False),
    use_container_width=True,
    hide_index=True,
    column_config={
        "timestamp": st.column_config.DatetimeColumn("Timestamp"),
        "count": st.column_config.NumberColumn("Errors", format="%d"),
    }
)

# ====================== EXECUTIVE EXPORT ======================
st.divider()

col1, col2 = st.columns([1, 2])

with col1:
    if st.button("📥 Export Executive Report (Excel)", use_container_width=True, type="primary"):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # Summary sheet
            summary = filtered.groupby(['hall', 'rack_type', 'building'])['count'].sum().reset_index()
            summary.to_excel(writer, index=False, sheet_name="By Rack & Hall")
            
            # Full log
            filtered.to_excel(writer, index=False, sheet_name="Full Log")
            
            # Executive summary sheet
            exec_df = pd.DataFrame({
                "Metric": ["Total Errors", "Unique Racks", "Halls Active", "Report Date"],
                "Value": [total_errors, unique_racks, active_halls, datetime.now().strftime("%Y-%m-%d")]
            })
            exec_df.to_excel(writer, index=False, sheet_name="Executive Summary")
        
        buffer.seek(0)
        st.download_button(
            "Download Report",
            data=buffer,
            file_name=f"Validation_Errors_Executive_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.caption("Data is automatically logged whenever processors run. This dashboard reads from `data/validation_error_log.xlsx`.")

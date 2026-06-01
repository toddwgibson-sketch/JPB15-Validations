import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
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
    st.caption("Run any validation tool (e.g. SYD20 QFAB). Errors are logged automatically to this file.")
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

# ====================== ERROR BREAKDOWN BY BUILDING (card view) ======================
st.markdown('<div class="section-header">Error Breakdown by Building</div>', unsafe_allow_html=True)

if not filtered.empty:
    building_order = sorted(filtered['building'].unique())
    num_bldgs = len(building_order)
    cols = st.columns(num_bldgs) if num_bldgs > 0 else [st.container()]

    # Consistent category order + nice labels
    category_order = ["Downlink", "Mismatch", "optics", "fec_ber"]
    cat_label = {
        "Downlink": "Downlink",
        "Mismatch": "Mismatch",
        "optics": "Optics",
        "fec_ber": "FEC BER"
    }

    for i, bldg in enumerate(building_order):
        with cols[i]:
            bldg_df = filtered[filtered['building'] == bldg]
            cat_counts = bldg_df.groupby('error_category')['count'].sum().to_dict()

            st.markdown(f"**{bldg}**")
            bldg_total = sum(cat_counts.values())
            st.caption(f"**Total errors in {bldg}: {int(bldg_total)}**")

            for cat in category_order:
                label = cat_label.get(cat, cat)
                val = int(cat_counts.get(cat, 0))
                st.metric(label, val)
else:
    st.info("No building data available yet.")

st.divider()

# ====================== QUICK BREAKDOWNS ======================
col1, col2 = st.columns(2)

with col1:
    st.markdown("**Errors by Hall**")
    if not filtered.empty:
        hall_sum = filtered.groupby('hall')['count'].sum().reset_index()
        fig_hall = px.bar(hall_sum, x='hall', y='count', color='hall',
                          color_discrete_sequence=px.colors.qualitative.Pastel,
                          text='count')
        fig_hall.update_layout(showlegend=False, height=280, margin=dict(t=10, b=10))
        st.plotly_chart(fig_hall, use_container_width=True)
    else:
        st.info("No data")

with col2:
    st.markdown("**Errors by Rack Type**")
    if not filtered.empty:
        type_sum = filtered.groupby('rack_type')['count'].sum().reset_index()
        fig_type = px.pie(type_sum, names='rack_type', values='count', hole=0.5,
                          color_discrete_sequence=px.colors.qualitative.Set3)
        fig_type.update_layout(height=280, margin=dict(t=10, b=10))
        st.plotly_chart(fig_type, use_container_width=True)
    else:
        st.info("No data")

st.divider()

# ====================== CATEGORY × BUILDING PIVOT (what you asked for) ======================
st.markdown('<div class="section-header">Errors by Category × Building</div>', unsafe_allow_html=True)

if not filtered.empty:
    # Create the clean pivot exactly like the summary you export from QFAB
    pivot = (
        filtered.pivot_table(
            index="error_category",
            columns="building",
            values="count",
            aggfunc="sum",
            fill_value=0
        )
        .astype(int)
    )
    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=False)
    pivot.loc["TOTAL"] = pivot.sum()

    st.dataframe(
        pivot,
        use_container_width=True,
        column_config={
            col: st.column_config.NumberColumn(col, format="%d") 
            for col in pivot.columns
        }
    )

    # Nicer Plotly bar chart for category totals
    cat_totals = pivot.drop("TOTAL")["Total"].reset_index()
    cat_totals.columns = ["Category", "Errors"]

    fig = px.bar(
        cat_totals,
        x="Category",
        y="Errors",
        title="Total Errors by Category",
        color="Category",
        color_discrete_sequence=px.colors.qualitative.Bold,
        text="Errors"
    )
    fig.update_traces(textposition="outside", textfont_size=13)
    fig.update_layout(
        height=340,
        showlegend=False,
        margin=dict(t=40, b=20),
        xaxis_title=None,
        yaxis_title="Number of Errors"
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No data in current filter.")

st.divider()

# ====================== TOP PROBLEMATIC RACKS ======================
st.markdown('<div class="section-header">Top Problematic Racks / Buildings</div>', unsafe_allow_html=True)

if not filtered.empty:
    rack_summary = (
        filtered.groupby(['building', 'hall', 'rack_type'])['count']
        .sum()
        .reset_index()
        .sort_values('count', ascending=False)
        .head(15)
    )

    fig_racks = px.bar(
        rack_summary,
        x='building',
        y='count',
        color='hall',
        hover_data=['rack_type'],
        title=None,
        color_discrete_sequence=px.colors.qualitative.Set2,
        text='count'
    )
    fig_racks.update_traces(textposition="outside")
    fig_racks.update_layout(
        height=420,
        xaxis_tickangle=-35,
        margin=dict(t=20, b=40),
        yaxis_title="Total Errors",
        xaxis_title=None,
        legend_title="Hall"
    )
    st.plotly_chart(fig_racks, use_container_width=True)
else:
    st.info("No data for top racks.")

st.divider()

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

st.caption("Errors are logged automatically from the validation tools. Data source: data/validation_error_log.xlsx inside the repo.")

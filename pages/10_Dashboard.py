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
        font-size: 50rem;
        font-weight: 700;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 1.05rem;
        margin-bottom: 1.8rem;
    }
    .section-header {
        font-size: 1.35rem;
        font-weight: 600;
        margin-top: 1.8rem;
        margin-bottom: 0.6rem;
    }
    .kpi-card {
        background: var(--secondary-background-color, white);
        border-radius: 16px;
        padding: 1.2rem 1.4rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.06);
        border: 1px solid var(--border-color, #e2e8f0);
    }
    .metric-value {
        font-size: 1.9rem;
        font-weight: 700;
        color: inherit;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">Dashboard</p>', unsafe_allow_html=True)
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


def get_latest_snapshot(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Return only the most recent row for each (hall, building, error_category).
    This makes the dashboard show *current* issues instead of accumulating history.
    Re-running the same blocks will overwrite the previous numbers for those blocks.
    """
    if dataframe.empty:
        return dataframe
    return (
        dataframe.sort_values('timestamp')
        .groupby(['hall', 'building', 'error_category'], as_index=False)
        .last()
    )

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

# Current snapshot only (latest entry per block + category).
# This is what powers the main views so re-running the same blocks overwrites old numbers.
current = get_latest_snapshot(filtered)

# ====================== EXECUTIVE KPI CARDS ======================
st.markdown('<div class="section-header">Executive Snapshot</div>', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

total_errors = int(current['count'].sum())
unique_racks = current['building'].nunique()
active_halls = current['hall'].nunique()
avg_errors_per_rack = round(total_errors / unique_racks, 1) if unique_racks > 0 else 0

with col1:
    st.metric("Total Errors Logged", f"{total_errors:,}")
with col2:
    st.metric("Unique Blocks", unique_racks)
with col3:
    st.metric("Halls Active", active_halls)
with col4:
    st.metric("Avg Errors per Rack", avg_errors_per_rack)

st.divider()

# ====================== ERROR BREAKDOWN BY BUILDING (widget cards) ======================
st.markdown('<div class="section-header">Error Breakdown by Block</div>', unsafe_allow_html=True)

# Nice consistent colors for the 4 categories (widget style)
CAT_COLORS = {
    "Downlink": "#e74c3c",
    "Mismatch": "#f39c12",
    "optics": "#3498db",
    "fec_ber": "#9b59b6"
}
CAT_LABELS = {
    "Downlink": "Downlink",
    "Mismatch": "Mismatch",
    "optics": "Optics",
    "fec_ber": "FEC BER"
}

if not current.empty:
    building_order = sorted(current['building'].unique())
    CARDS_PER_ROW = 5   # Adjust this if you want 4 or 6 per row

    category_order = list(CAT_LABELS.keys())

    # Build grid: 5 cards per row, then wrap to next row
    for start_idx in range(0, len(building_order), CARDS_PER_ROW):
        row_buildings = building_order[start_idx : start_idx + CARDS_PER_ROW]
        cols = st.columns(CARDS_PER_ROW)

        for i, bldg in enumerate(row_buildings):
            with cols[i]:
                bldg_df = current[current['building'] == bldg]
                cat_counts = bldg_df.groupby('error_category')['count'].sum().to_dict()

                bldg_total = int(sum(cat_counts.values()))

                # Card container (widget style)
                with st.container(border=True):
                    # Block header + total (compact)
                    st.markdown(f"<div style='font-size:1.05rem; font-weight:600; margin-bottom:2px'>{bldg}</div>", unsafe_allow_html=True)
                    st.markdown(f"<div style='font-size:1.9rem; font-weight:700; line-height:1.1; margin-bottom:6px'>{bldg_total}</div>", unsafe_allow_html=True)

                    # Mini horizontal stacked bar showing the split
                    bar_data = []
                    for cat in category_order:
                        val = int(cat_counts.get(cat, 0))
                        if val > 0:
                            bar_data.append({
                                "Category": CAT_LABELS[cat],
                                "Count": val,
                                "Color": CAT_COLORS[cat]
                            })

                    if bar_data:
                        bar_df = pd.DataFrame(bar_data)
                        fig = px.bar(
                            bar_df,
                            x="Count",
                            y=[""] * len(bar_df),
                            color="Category",
                            orientation="h",
                            color_discrete_map={d["Category"]: d["Color"] for d in bar_data},
                            height=48
                        )
                        fig.update_layout(
                            barmode="stack",
                            margin=dict(l=0, r=0, t=0, b=0),
                            xaxis_visible=False,
                            yaxis_visible=False,
                            showlegend=False,
                            height=48
                        )
                        fig.update_traces(marker_line_width=0)
                        st.plotly_chart(
                            fig, 
                            use_container_width=True, 
                            key=f"bldg_bar_{bldg}",           # ← unique key fixes the DuplicateElementId error
                            config={"displayModeBar": False}
                        )

                    # Compact category list with small font
                    st.markdown("<div style='margin-top:4px; font-size:0.82rem; line-height:1.25'>", unsafe_allow_html=True)
                    for cat in category_order:
                        label = CAT_LABELS[cat]
                        val = int(cat_counts.get(cat, 0))
                        color = CAT_COLORS[cat]
                        st.markdown(
                            f"<span style='color:{color}; font-weight:600'>■</span> {label}: <b>{val}</b>",
                            unsafe_allow_html=True
                        )
                    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.info("No block data available yet.")

st.divider()

# ====================== CATEGORY × BUILDING PIVOT (what you asked for) ======================
st.markdown('<div class="section-header">Errors by Category × Block</div>', unsafe_allow_html=True)

if not current.empty:
    # Create the clean pivot exactly like the summary you export from QFAB
    pivot = (
        current.pivot_table(
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
    st.info("No current data after applying latest-per-block logic.")

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
            # Summary sheet (current snapshot only)
            summary = current.groupby(['hall', 'rack_type', 'building'])['count'].sum().reset_index()
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

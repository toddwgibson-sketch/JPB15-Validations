import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

st.set_page_config(page_title="Rack Validation", layout="wide", page_icon="🛠")

# Make it look nice
st.markdown("""
<style>
    .big-font {font-size: 50px !important; font-weight: bold;}
    .reportview-container {background: #f0f2f6;}
</style>
""", unsafe_allow_html=True)

st.title("🛠 Rack Validation Dashboard")
st.markdown("**Live Tracking • Clean Reports • Always Updated**")

# Persistent Excel File
DATA_FILE = "rack_tracking.xlsx"

# Create file if it doesn't exist
if not os.path.exists(DATA_FILE):
    df = pd.DataFrame({
        'rack_id': ['RACK-001', 'RACK-002', 'RACK-003'],
        'errors': [12, 8, 25],
        'items_checked': [250, 180, 320],
        'error_rate': [4.8, 4.4, 7.8],
        'status': ['High', 'Medium', 'Critical'],
        'last_updated': ['2025-06-01', '2025-06-01', '2025-05-31'],
        'notes': ['Power issue', '', 'Multiple missing items']
    })
    df.to_excel(DATA_FILE, index=False)

# Load data
df = pd.read_excel(DATA_FILE)

# Sidebar
st.sidebar.header("Tools")
if st.sidebar.button("🔄 Refresh Dashboard"):
    st.rerun()

# Main Metrics
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Racks", len(df))
col2.metric("Total Errors", int(df['errors'].sum()))
col3.metric("Avg Error Rate", f"{df['error_rate'].mean():.1f}%")
col4.metric("Critical Racks", len(df[df['status'] == 'Critical']))

st.divider()

# Editable Table
st.subheader("📋 Edit Rack Data")
edited_df = st.data_editor(
    df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "rack_id": st.column_config.TextColumn("Rack ID", disabled=True),
        "errors": st.column_config.NumberColumn("Number of Errors", min_value=0),
        "items_checked": st.column_config.NumberColumn("Items Checked", min_value=0),
        "error_rate": st.column_config.NumberColumn("Error Rate %", disabled=True),
        "status": st.column_config.SelectboxColumn("Status", options=["Low", "Medium", "High", "Critical"]),
        "last_updated": st.column_config.TextColumn("Last Updated", disabled=True),
        "notes": st.column_config.TextColumn("Notes / Comments", width="medium")
    }
)

# Save changes automatically
if not edited_df.equals(df):
    edited_df['error_rate'] = round((edited_df['errors'] / edited_df['items_checked'].replace(0, 1)) * 100, 1)
    edited_df['last_updated'] = datetime.now().strftime("%Y-%m-%d")
    edited_df.to_excel(DATA_FILE, index=False)
    st.success("✅ Changes saved successfully!")

# Charts
st.divider()
c1, c2 = st.columns(2)

with c1:
    st.plotly_chart(px.bar(df, x='rack_id', y='errors', color='status', title="Errors by Rack"), 
                    use_container_width=True)

with c2:
    st.plotly_chart(px.pie(df, names='status', values='errors', title="Error Severity"), 
                    use_container_width=True)

# Rack Detail
st.subheader("🔍 View Rack Details")
selected = st.selectbox("Choose Rack", df['rack_id'])
st.dataframe(df[df['rack_id'] == selected], use_container_width=True)

st.caption(f"Data is saved in rack_tracking.xlsx • Last refreshed: {datetime.now().strftime('%H:%M')}")

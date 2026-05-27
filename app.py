import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="LV Portal Tools",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom Header with Logo ─────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 6])

with col_logo:
    logo_path = Path(__file__).parent / "assets" / "logo.png"
    if logo_path.exists():
        st.image(str(logo_path), width=90)
    else:
        st.markdown("### 🔧")

with col_title:
    st.markdown("""
    <h1 style='margin-bottom: 0; color: #1F4E79;'>LV Portal Validation Suite</h1>
    <p style='margin-top: 4px; color: #555; font-size: 1.05rem;'>
        Unified tools for processing LV Portal and Slack validation reports
    </p>
    """, unsafe_allow_html=True)

st.divider()

# ── Welcome Section ─────────────────────────────────────────────────────────
st.markdown("""
### Welcome 👋

This internal toolkit helps network engineers quickly process and analyse validation reports from the LV Portal and Slack reporting systems.

Select a tool from the **sidebar** to get started.
""")

st.divider()

# ── Tool Cards ──────────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3, gap="large")

with col1:
    st.markdown("""
    <div style="background-color:#F8FAFC; padding:20px; border-radius:12px; border: 1px solid #E2E8F0; height: 220px;">
        <h3 style="color:#1F4E79; margin-top:0;">🖥️ Compute Link Validator</h3>
        <p><b>Best for:</b> GPU / Compute to T0 link validation</p>
        <ul style="font-size:0.92rem; margin-bottom:0;">
            <li>Ghost host detection</li>
            <li>Compute Optics + FEC</li>
            <li>Summary dashboard</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div style="background-color:#F8FAFC; padding:20px; border-radius:12px; border: 1px solid #E2E8F0; height: 220px;">
        <h3 style="color:#1F4E79; margin-top:0;">🔗 QFAB / T1→T0 Validator</h3>
        <p><b>Best for:</b> Newer LV Portal T1-T0 exports</p>
        <ul style="font-size:0.92rem; margin-bottom:0;">
            <li>Per-channel Optics analysis</li>
            <li>DG19-style formatting</li>
            <li>Strong cutsheet fallback</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div style="background-color:#F8FAFC; padding:20px; border-radius:12px; border: 1px solid #E2E8F0; height: 220px;">
        <h3 style="color:#1F4E79; margin-top:0;">📊 Slack Report Highlighter</h3>
        <p><b>Best for:</b> Classic Slack validation reports</p>
        <ul style="font-size:0.92rem; margin-bottom:0;">
            <li>Full LLDP + Mismatch analysis</li>
            <li>Recurring issue detection</li>
            <li>History flags across reports</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ── Tips ────────────────────────────────────────────────────────────────────
with st.expander("💡 Quick Tips"):
    st.markdown("""
    - All tools support **multiple cutsheets** (very useful for mismatch resolution)
    - You can upload a **previous formatted report** in most tools to highlight recurring issues
    - Output files are downloaded directly — nothing is stored on the server
    - For best results, use the **latest cutsheet** available for the hall/site
    """)

st.caption("Internal tool • Network Validation Team")
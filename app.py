import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="staging env",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Header ───────────────────────────────────────────────────────────────────
col_logo, col_title = st.columns([0.8, 6])

with col_logo:
    logo_path = Path(__file__).parent / "assets" / "LOGO.jpg"
    if logo_path.exists():
        st.image(str(logo_path), width=155)
    else:
        st.markdown("<div style='font-size:42px; margin-top:-8px;'>🔧</div>", unsafe_allow_html=True)

with col_title:
    st.markdown("""
    <h1 style='margin:0; padding:0; color:#FFFFFF; font-size:2.1rem; font-weight:700;'>
        TEST - ignore the look and feel. function first
    </h1>
    <p style='margin:4px 0 0 0; color:#FFFFFF; font-size:1.05rem;'>
        can make it pretty later
    </p>
    """, unsafe_allow_html=True)

st.divider()

# ── Welcome ──────────────────────────────────────────────────────────────────
st.markdown("""
<div style="max-width:820px;">
<p style="color:#4A5568; font-size:1.02rem; line-height:1.55;">
abc
</p>
<p style="color:#4A5568; font-size:1.02rem;">
Select a tool from the <b>sidebar</b> to get started.
</p>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tool Cards ───────────────────────────────────────────────────────────────
st.markdown("### Available Tools")

col1, col2, col3 = st.columns(3, gap="large")

card_style = """
<style>
.card {
    background: white;
    border-radius: 14px;
    padding: 24px 22px;
    border: 1px solid #E8ECF0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    height: 100%;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 20px rgba(0,0,0,0.08);
}
.card h3 {
    margin: 0 0 12px 0;
    font-size: 1.15rem;
    color: #1E3A5F;
}
.card p {
    color: #5A6A7A;
    font-size: 0.95rem;
    margin: 0 0 14px 0;
}
.card ul {
    padding-left: 18px;
    margin: 0;
    color: #4A5568;
    font-size: 0.9rem;
    line-height: 1.5;
}
.card li {
    margin-bottom: 4px;
}
</style>
"""

st.markdown(card_style, unsafe_allow_html=True)

with col1:
    st.markdown("""
    <div class="card">
        <h3>🖥️ T0 to Host</h3>
        <p><b>Best for:</b>card template/p>
        <ul>
            <li>a</li>
            <li>b</li>
            <li>b</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="card">
         <h3>🖥️ T1 LVV</h3>
        <p><b>Best for:</b>card template/p>
        <ul>
            <li>a</li>
            <li>b</li>
            <li>b</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="card">
        <h3>📊 T1 Slack</h3>
         <p><b>Best for:</b>card template/p>
        <ul>
            <li>a</li>
            <li>b</li>
            <li>b</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)

# ── Tips ─────────────────────────────────────────────────────────────────────
with st.expander("FYI"):
    st.markdown("""
    - Output files are downloaded directly to your machine — nothing is stored on the server
    """)

st.caption("Internal tool • Network Validation Team")

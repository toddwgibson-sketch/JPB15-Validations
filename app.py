import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="staging env",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Header ───────────────────────────────────────────────────────────────────
def show_header(title: str, subtitle: str = ""):
    # Logo on top
    logo_path = Path(__file__).parent / "assets" / "LOGO.jpg"
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if logo_path.exists():
            st.image(str(logo_path), width=220)
        else:
            st.markdown(
                "<div style='font-size:48px; text-align:center; margin-bottom:8px;'>🔧</div>", 
                unsafe_allow_html=True
            )

    # Title + Subtitle
    st.markdown(f"""
        <div style="text-align: center; margin-top: 8px;">
            <h1 style='margin:0; padding:0; color:#FFFFFF; font-size:2.0rem; font-weight:700;'>
                {title}
            </h1>
            {f'<p style="margin:6px 0 0 0; color:#AAAAAA; font-size:1.05rem;">{subtitle}</p>' if subtitle else ''}
        </div>
    """, unsafe_allow_html=True)

    st.divider()


# ── Usage ─────────────────────────────────────────────────────────────
show_header(
    title="TEST - ignore the look and feel. function first",
    subtitle="can make it pretty later"
)

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

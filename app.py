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
            st.image(str(logo_path), width=880)
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
    title="TEST",
    subtitle="test"
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


# ── Tips ─────────────────────────────────────────────────────────────────────
with st.expander("FYI"):
    st.markdown("""
    - Output files are downloaded directly to your machine — nothing is stored on the server
    """)

st.caption("Internal tool • Quadconnect")

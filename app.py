import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="JPB15 Validations",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================== BASIC AUTH ======================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def login_form():
    st.title("🔐 Login")
    st.markdown("Please log in to access the tools.")

    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")

        if submitted:
            if username == "admin" and password == "admin":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid username or password")

# If not logged in, show login and stop the app
if not st.session_state.authenticated:
    login_form()
    st.stop()

# If we reach here, the user is logged in
with st.sidebar:
    st.success("Logged in as **admin**")
    if st.button("Logout"):
        st.session_state.authenticated = False
        st.rerun()
# =========================================================

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
    title="JPB15 Validations",
    subtitle="Internal Validation Tools"
)

# ── Welcome ──────────────────────────────────────────────────────────────────
st.markdown("""
<div style="max-width:820px;">
<p style="color:#4A5568; font-size:1.02rem; line-height:1.55;">
Select a tool from the <b>sidebar</b> to get started.
</p>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tips ─────────────────────────────────────────────────────────────────────
with st.expander("FYI"):
    st.markdown("""
    - Output files are downloaded directly to your machine — nothing is stored on the server
    """)

st.caption("Internal tool • Quadconnect • ")

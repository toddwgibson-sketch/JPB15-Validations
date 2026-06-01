import streamlit as st

def require_login():
    """
    Simple username/password gate for the entire app.
    Currently hardcoded to admin / admin.
    Call this at the very top of app.py and every page.
    """
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.title("🔐 Login Required")
        st.markdown("This is an internal tool. Please log in to continue.")

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login")

            if submitted:
                if username == "admin" and password == "admin":
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Invalid username or password")

        st.stop()

    # If we reach here, user is logged in
    # Add a logout button in the sidebar for convenience
    with st.sidebar:
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.rerun()
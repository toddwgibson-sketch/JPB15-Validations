#!/usr/bin/env python3
"""
LV Portal Formatter - Streamlit Version (JBP15 T0→Host)

This is a thin Streamlit wrapper around the original logic in:
lv_portal_formatter_HOPS_NEW_28_05 (1).py

It replaces the Tkinter file pickers and main() with a clean web UI.
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os
import io

# Add the original script's directory so we can import its logic functions
ORIGINAL_DIR = Path(__file__).parent
sys.path.insert(0, str(ORIGINAL_DIR))

try:
    # We import the original module (it may have Tkinter at top level, but we only use its pure functions)
    import lv_portal_formatter_HOPS_NEW_28_05__1_ as original_logic
except Exception as e:
    st.error(f"Failed to import the original logic module: {e}")
    st.stop()

st.set_page_config(page_title="LV Portal Formatter (JBP15)", page_icon="🔌", layout="wide")
st.title("LV Portal Validation Formatter — JBP15 T0→Host")
st.caption("Streamlit version • Same logic and output format as the original desktop tool")

with st.sidebar:
    st.header("How to use")
    st.markdown("""
    1. Upload one or more **GPU/Compute Cutsheets** (the special ones with OHR/FDF/NIC data).
    2. Upload the main **LV Portal Validation Export**.
    3. (Optional) Upload a previous formatted report for the History column.
    4. Click **Process**.
    """)
    st.info("Only T0/T1 compute links are processed (pure t0-host compute rows are filtered).")

# =============================================================================
# File Uploads
# =============================================================================
st.subheader("1. GPU / Compute Cutsheets")
gpu_files = st.file_uploader(
    "Upload one or more GPU/Compute cutsheets (Ctrl/Cmd + click for multiple)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="gpu_cutsheets"
)

st.subheader("2. LV Portal Validation Export")
portal_file = st.file_uploader(
    "LV Portal Validation Export (.xlsx)",
    type=["xlsx"],
    key="portal_report"
)

st.subheader("3. Previous Report (Optional)")
prev_file = st.file_uploader(
    "Previous formatted report (for History comparison)",
    type=["xlsx"],
    key="prev_report"
)

# =============================================================================
# Processing
# =============================================================================
if st.button("🚀 Process", type="primary", disabled=not (gpu_files and portal_file)):
    with st.spinner("Processing... This can take 1-3 minutes on large reports."):
        try:
            # Save uploads to temporary files so the original logic can consume them
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save GPU cutsheets
                gpu_paths = []
                for i, f in enumerate(gpu_files):
                    p = os.path.join(tmpdir, f"gpu_{i}.xlsx")
                    with open(p, "wb") as out:
                        out.write(f.getbuffer())
                    gpu_paths.append(p)

                # Save main report
                report_path = os.path.join(tmpdir, "report.xlsx")
                with open(report_path, "wb") as out:
                    out.write(portal_file.getbuffer())

                # Save previous report if provided
                prev_path = None
                if prev_file:
                    prev_path = os.path.join(tmpdir, "prev_report.xlsx")
                    with open(prev_path, "wb") as out:
                        out.write(prev_file.getbuffer())

                st.write("Files saved. Running core processing logic...")

                # ------------------------------------------------------------------
                # Call into the original logic
                # We temporarily override the file pickers so the original main()
                # can run without Tkinter dialogs.
                # ------------------------------------------------------------------
                original_pick = original_logic.pick_file
                original_pick_multi = original_logic.pick_multiple_files

                # Monkey-patch the pickers so the original main() uses our uploaded files
                def fake_pick_multi(title, *a, **k):
                    if "GPU" in title or "Compute" in title:
                        return gpu_paths
                    return []

                def fake_pick(title, *a, **k):
                    title_lower = title.lower()
                    if "lv portal" in title_lower or "validation" in title_lower:
                        return report_path
                    if "previous" in title_lower or "history" in title_lower:
                        return prev_path
                    return None

                original_logic.pick_multiple_files = fake_pick_multi
                original_logic.pick_file = fake_pick

                # Prevent the original main() from calling sys.exit() and killing Streamlit
                original_exit = sys.exit
                sys.exit = lambda *a, **k: None

                try:
                    original_logic.main()
                finally:
                    sys.exit = original_exit
                    original_logic.pick_multiple_files = original_pick_multi
                    original_logic.pick_file = original_pick

                # Find the output file that was just created
                base = os.path.splitext(report_path)[0]
                out_path = base + "_formatted.xlsx"

                if os.path.exists(out_path):
                    with open(out_path, "rb") as f:
                        result_bytes = f.read()

                    st.success("Processing complete!")

                    st.download_button(
                        label="📥 Download Formatted Report",
                        data=result_bytes,
                        file_name=os.path.basename(out_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error("Processing finished but no output file was found.")

        except Exception as e:
            st.error("Processing failed")
            st.exception(e)

st.caption("Converted to Streamlit • Core logic unchanged from the original desktop version")

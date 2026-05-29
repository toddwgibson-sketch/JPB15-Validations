#!/usr/bin/env python3
"""
LV Portal Formatter - Streamlit Page
Place this file in the pages/ folder.

This is the UI only. It imports the heavy logic from ../lv_portal_logic.py
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os

# Add parent directory to path so we can import the logic module
sys.path.append(str(Path(__file__).parent.parent))

try:
    import lv_portal_logic as logic
except ImportError as e:
    st.error(
        "Could not import lv_portal_logic.py\n\n"
        "Make sure the file 'lv_portal_logic.py' exists in the main project folder "
        "(one level above the pages/ folder).\n\n"
        f"Error: {e}"
    )
    st.stop()

st.set_page_config(page_title="LV Portal Formatter - JBP15", page_icon="🔌", layout="wide")
st.title("LV Portal Validation Formatter — JBP15 T0→Host")
st.caption("Streamlit version • Uses the original logic from lv_portal_logic.py")

with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    1. Upload one or more **GPU/Compute Cutsheets**.
    2. Upload the main **LV Portal Validation Export**.
    3. (Optional) Upload a previous report for the History column.
    4. Click **Process**.
    """)
    st.info("Only T0/T1 compute links are processed.")

# =============================================================================
# UI - All widgets outside any conditional to avoid key conflicts
# =============================================================================
st.subheader("1. GPU / Compute Cutsheets")
gpu_files = st.file_uploader(
    "Upload one or more GPU/Compute cutsheets",
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
    "Previous formatted report (for History)",
    type=["xlsx"],
    key="prev_report"
)

# =============================================================================
# Processing
# =============================================================================
if st.button("🚀 Process", type="primary", disabled=not (gpu_files and portal_file)):
    with st.spinner("Processing... This can take a while on large reports."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save GPU cutsheets
                gpu_paths = []
                for i, f in enumerate(gpu_files):
                    p = os.path.join(tmpdir, f"gpu_cutsheet_{i}.xlsx")
                    with open(p, "wb") as out:
                        out.write(f.getbuffer())
                    gpu_paths.append(p)

                # Save main report
                report_path = os.path.join(tmpdir, "lv_portal_report.xlsx")
                with open(report_path, "wb") as out:
                    out.write(portal_file.getbuffer())

                # Save previous report if provided
                prev_path = None
                if prev_file:
                    prev_path = os.path.join(tmpdir, "previous_report.xlsx")
                    with open(prev_path, "wb") as out:
                        out.write(prev_file.getbuffer())

                st.write("Files prepared. Running processing logic...")

                # Save original functions
                original_pick = logic.pick_file
                original_pick_multi = logic.pick_multiple_files

                # Patch the pickers
                def fake_pick_multi(title, *args, **kwargs):
                    if "gpu" in title.lower() or "compute" in title.lower():
                        return gpu_paths
                    return []

                def fake_pick(title, *args, **kwargs):
                    title_lower = title.lower()
                    if "lv portal" in title_lower or "validation" in title_lower:
                        return report_path
                    if "previous" in title_lower or "history" in title_lower:
                        return prev_path
                    return None

                logic.pick_multiple_files = fake_pick_multi
                logic.pick_file = fake_pick

                # Prevent sys.exit from killing the Streamlit session
                original_exit = sys.exit
                sys.exit = lambda *a, **k: None

                try:
                    # Run the original main() — it will save the output file
                    logic.main()
                finally:
                    # Restore everything
                    sys.exit = original_exit
                    logic.pick_file = original_pick
                    logic.pick_multiple_files = original_pick_multi

                # Look for the output file
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

st.caption("Streamlit wrapper • Core logic lives in lv_portal_logic.py (in the project root)")

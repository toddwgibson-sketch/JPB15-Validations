#!/usr/bin/env python3
"""
LV Portal Formatter - Streamlit Version (JBP15 T0→Host)

Clean Streamlit wrapper around the core logic.

INSTRUCTIONS:
1. Rename the original file from:
      lv_portal_formatter_HOPS_NEW_28_05 (1).py
   to:
      lv_portal_logic.py

2. Place both files in the same folder.

3. Run with:
      streamlit run lv_portal_formatter_streamlit.py
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os

# =============================================================================
# Import the core logic module
# (Make sure you renamed the original file to lv_portal_logic.py)
# =============================================================================
try:
    import lv_portal_logic as lv_logic
except ImportError as e:
    st.error(
        "Could not import the logic module.\n\n"
        "Please make sure you renamed the original file to:\n"
        "   lv_portal_logic.py\n\n"
        f"Original error: {e}"
    )
    st.stop()

st.set_page_config(page_title="LV Portal Formatter (JBP15)", page_icon="🔌", layout="wide")
st.title("LV Portal Validation Formatter — JBP15 T0→Host")
st.caption("Streamlit version • Reuses the original logic from lv_portal_logic.py")

with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    1. Upload one or more **GPU/Compute Cutsheets**.
    2. Upload the **LV Portal Validation Export**.
    3. (Optional) Upload a previous report for history comparison.
    4. Click **Process**.
    """)
    st.info("Only processes T0/T1 compute links (filters pure t0-host compute rows).")

# =============================================================================
# File Uploads
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
    with st.spinner("Processing... This can take 1–3 minutes on large reports."):
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

                st.write("Files prepared. Running original processing logic...")

                # ------------------------------------------------------------------
                # Monkey-patch the original pickers so main() uses our files
                # ------------------------------------------------------------------
                original_pick = lv_logic.pick_file
                original_pick_multi = lv_logic.pick_multiple_files

                def fake_pick_multi(title, *args, **kwargs):
                    title_lower = title.lower()
                    if "gpu" in title_lower or "compute" in title_lower:
                        return gpu_paths
                    return []

                def fake_pick(title, *args, **kwargs):
                    title_lower = title.lower()
                    if "lv portal" in title_lower or "validation" in title_lower:
                        return report_path
                    if "previous" in title_lower or "history" in title_lower:
                        return prev_path
                    return None

                lv_logic.pick_multiple_files = fake_pick_multi
                lv_logic.pick_file = fake_pick

                # Prevent the original main() from killing the Streamlit session
                original_exit = sys.exit
                sys.exit = lambda *a, **k: None

                try:
                    # Run the original main() — it will write the output file to disk
                    lv_logic.main()
                finally:
                    sys.exit = original_exit
                    lv_logic.pick_multiple_files = original_pick_multi
                    lv_logic.pick_file = original_pick

                # Find the output file that was created
                base = os.path.splitext(report_path)[0]
                out_path = base + "_formatted.xlsx"

                if os.path.exists(out_path):
                    with open(out_path, "rb") as f:
                        result_bytes = f.read()

                    st.success("Processing complete!")

                    st.download_button(
                        label="📥 Download Formatted LV Portal Report",
                        data=result_bytes,
                        file_name=os.path.basename(out_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error("Processing finished, but no output file was found.")

        except Exception as e:
            st.error("Processing failed")
            st.exception(e)

st.caption("Streamlit wrapper • Core logic unchanged from lv_portal_logic.py")
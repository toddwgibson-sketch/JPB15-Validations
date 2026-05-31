#!/usr/bin/env python3
"""
JPB19 HOPS / GPU Validation Formatter — Streamlit Version

Converted from the original HOPS CODE19 Tkinter script.
All core processing logic is preserved from the original.
"""

import streamlit as st
import os
import tempfile
import io
import contextlib
from pathlib import Path

try:
    from hops_code19 import build_combined_lookups, process_one
    HAS_ORIGINAL = True
except Exception:
    HAS_ORIGINAL = False


st.set_page_config(
    page_title="JPB19 HOPS Formatter",
    page_icon="🧠",
    layout="wide"
)

st.title("JPB19 HOPS / GPU Validation Formatter")
st.caption("Streamlit version of the HOPS CODE19 tool (logic unchanged)")

with st.expander("How it works", expanded=False):
    st.markdown("""
    - Upload the **Combined Cutsheet** (GPU/HOPS version)
    - Upload one or more **Validation Result** files
    - The tool applies the same processing as the original Tkinter version:
        - LLDP / Optic / FEC / Interface Down enrichment
        - CT-off detection
        - Pair highlighting
        - Rack-based output naming
    - Download the formatted files individually or as a ZIP
    """)

# === File Uploads ===
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Combined Cutsheet")
    cutsheet_file = st.file_uploader(
        "Combined GPU/HOPS cutsheet (.xlsx)",
        type=["xlsx", "xlsm"],
        key="hops_cutsheet"
    )

with col2:
    st.subheader("2. Validation Files")
    input_files = st.file_uploader(
        "One or more validation result files",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        key="hops_inputs"
    )

if not (cutsheet_file and input_files):
    st.info("Upload the combined cutsheet and at least one validation file to begin.")
    st.stop()

if st.button("🚀 Process Files", type="primary", use_container_width=True):
    if not HAS_ORIGINAL:
        st.error("Could not import the processing logic from 'hops_code19.py'.\n"
                 "Please make sure the original script is saved as 'hops_code19.py' in this folder.")
        st.stop()

    with st.status("Processing files...", expanded=True) as status:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save cutsheet
                cutsheet_path = os.path.join(tmpdir, cutsheet_file.name)
                with open(cutsheet_path, "wb") as f:
                    f.write(cutsheet_file.getbuffer())

                st.write("Loading combined cutsheet lookups...")
                combined = build_combined_lookups(cutsheet_path)

                results = []
                errors = []

                progress_bar = st.progress(0, text="Starting...")

                for idx, uf in enumerate(input_files):
                    progress = (idx) / len(input_files)
                    progress_bar.progress(progress, text=f"Processing {uf.name} ({idx+1}/{len(input_files)})")

                    in_path = os.path.join(tmpdir, uf.name)
                    with open(in_path, "wb") as f:
                        f.write(uf.getbuffer())

                    # Capture any print output from the original logic
                    log_output = io.StringIO()
                    with contextlib.redirect_stdout(log_output):
                        try:
                            out_path = process_one(in_path, combined)
                            with open(out_path, "rb") as f:
                                data = f.read()
                            results.append((uf.name, os.path.basename(out_path), data))
                        except Exception as e:
                            errors.append((uf.name, str(e)))

                progress_bar.progress(1.0, text="Finished")

                if results:
                    st.success(f"✅ Successfully processed {len(results)} file(s)")

                    st.subheader("Download Results")

                    # Individual downloads
                    for orig_name, final_name, data in results:
                        st.download_button(
                            label=f"📥 {final_name}",
                            data=data,
                            file_name=final_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=final_name
                        )

                    # ZIP option for multiple files
                    if len(results) > 1:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                            for _, fname, data in results:
                                zf.writestr(fname, data)
                        zip_buffer.seek(0)

                        st.download_button(
                            label="📦 Download All as ZIP",
                            data=zip_buffer.getvalue(),
                            file_name="HOPS_Formatted_Reports.zip",
                            mime="application/zip",
                            type="primary"
                        )

                if errors:
                    st.error(f"⚠️ {len(errors)} file(s) failed to process:")
                    for name, err in errors:
                        st.write(f"- **{name}**: {err}")

                status.update(label="Processing complete", state="complete")

        except Exception as e:
            status.update(label="Processing failed", state="error")
            st.error(f"An unexpected error occurred: {e}")
            st.exception(e)

st.caption("All processing happens locally on your machine. No data is sent to any server.")

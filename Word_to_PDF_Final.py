import streamlit as st
from docx2pdf import convert
import tempfile
import os
from pathlib import Path
import shutil
import time

st.set_page_config(page_title="DOCX to PDF Converter", layout="centered")

st.title("📄 DOCX to PDF Converter")
st.markdown("Drag and drop multiple **.docx** files below to convert them to PDF using Microsoft Word.")

uploaded_files = st.file_uploader("Upload DOCX files", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    st.info("Preparing to convert...")

    with st.spinner("Setting up files..."):
        temp_input_dir = Path(tempfile.mkdtemp())
        temp_output_dir = temp_input_dir / "converted_pdfs"
        temp_output_dir.mkdir(exist_ok=True)

        for file in uploaded_files:
            with open(temp_input_dir / file.name, "wb") as f:
                f.write(file.read())

    st.subheader("🔄 Conversion Progress")
    progress = st.progress(0)
    status = st.empty()

    docx_files = list(temp_input_dir.glob("*.docx"))
    for i, file in enumerate(docx_files, start=1):
        status.text(f"Converting: {file.name}")
        try:
            convert(str(file), str(temp_output_dir))
        except Exception as e:
            st.error(f"❌ Failed to convert {file.name}: {e}")
            continue
        progress.progress(i / len(docx_files))
        time.sleep(0.2)

    progress.empty()
    status.success("✅ Conversion complete!")

    st.subheader("📥 Download PDFs")
    for pdf_file in temp_output_dir.glob("*.pdf"):
        with open(pdf_file, "rb") as f:
            st.download_button(
                label=f"Download {pdf_file.name}",
                data=f,
                file_name=pdf_file.name,
                mime="application/pdf"
            )

    # Cleanup after session ends
    shutil.rmtree(temp_input_dir, ignore_errors=True)

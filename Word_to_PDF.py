import streamlit as st
from docx import Document
from docx2pdf import convert
import os
import shutil
from pathlib import Path
import tempfile
import time

st.set_page_config(page_title="DOCX to PDF Converter", layout="centered")

st.title("📄 DOCX to PDF Converter")
st.markdown("Drag and drop multiple **.docx** files below to convert them to PDF.")

uploaded_files = st.file_uploader("Upload DOCX files", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    with st.spinner("Preparing conversion..."):
        # Create temp directories
        temp_input_dir = Path(tempfile.mkdtemp())
        temp_output_dir = temp_input_dir / "output"
        temp_output_dir.mkdir(exist_ok=True)

        # Save uploaded files
        for file in uploaded_files:
            filepath = temp_input_dir / file.name
            with open(filepath, "wb") as f:
                f.write(file.read())

        st.success(f"Uploaded {len(uploaded_files)} file(s).")

        # Convert with progress
        st.subheader("Conversion Progress")
        progress_bar = st.progress(0)
        status_text = st.empty()

        docx_files = list(temp_input_dir.glob("*.docx"))
        total = len(docx_files)

        for i, file in enumerate(docx_files):
            status_text.text(f"Converting: {file.name}")
            convert(str(file), str(temp_output_dir))  # Conversion
            progress_bar.progress((i + 1) / total)
            time.sleep(0.2)  # simulate delay

        progress_bar.empty()
        status_text.text("✅ Conversion complete!")

        # Show download links
        st.subheader("Download PDFs:")
        for pdf_file in temp_output_dir.glob("*.pdf"):
            with open(pdf_file, "rb") as f:
                st.download_button(
                    label=f"📥 Download {pdf_file.name}",
                    data=f,
                    file_name=pdf_file.name,
                    mime="application/pdf"
                )

        # Clean-up option
        if st.button("Clean Temporary Files"):
            shutil.rmtree(temp_input_dir)
            st.success("Temporary files cleaned.")

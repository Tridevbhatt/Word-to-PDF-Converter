import streamlit as st
import pypandoc
import os
import random
import shutil
from pathlib import Path

# Set page config
st.set_page_config(page_title="DOCX to PDF Converter", layout="centered")

# Random background color
colors = ["#f0f8ff", "#ffe4e1", "#f5f5dc", "#e6e6fa", "#fafad2"]
background_color = random.choice(colors)

# Custom CSS for background and watermark
st.markdown(f"""
    <style>
        .stApp {{
            background-color: {background_color};
            position: relative;
        }}
        .watermark {{
            position: fixed;
            bottom: 20px;
            right: 20px;
            opacity: 0.1;
            font-size: 40px;
            z-index: 9999;
            color: #000000;
        }}
    </style>
    <div class="watermark">Trideb</div>
""", unsafe_allow_html=True)

st.title("📄 DOCX to PDF Converter")

uploaded_files = st.file_uploader("Upload DOCX files", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    output_dir = Path("converted_pdfs")
    output_dir.mkdir(exist_ok=True)

    st.info(f"Converting {len(uploaded_files)} DOCX file(s)...")
    progress_bar = st.progress(0)
    
    for i, uploaded_file in enumerate(uploaded_files):
        filename = Path(uploaded_file.name)
        docx_path = output_dir / filename
        with open(docx_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        pdf_path = output_dir / filename.with_suffix(".pdf")
        
        try:
            # Convert DOCX to PDF
            pypandoc.convert_file(str(docx_path), 'pdf', outputfile=str(pdf_path))
            st.success(f"✅ Converted: {filename.name}")
        except Exception as e:
            st.error(f"❌ Failed to convert {filename.name}: {e}")
        
        progress_bar.progress((i + 1) / len(uploaded_files))
    
    st.success("🎉 All files processed.")
    with st.expander("📂 Download Converted PDFs"):
        for pdf_file in output_dir.glob("*.pdf"):
            with open(pdf_file, "rb") as f:
                st.download_button(label=f"Download {pdf_file.name}", data=f, file_name=pdf_file.name)

    # Option to clean up
    if st.button("🧹 Clear All Converted Files"):
        shutil.rmtree(output_dir)
        st.experimental_rerun()
else:
    st.warning("Please upload at least one DOCX file.")

# streamlit_app.py

import streamlit as st
import os
import random
from docx2pdf import convert
import tempfile
import shutil

# -----------------------
# Utilities
# -----------------------

def generate_random_color():
    return "#{:06x}".format(random.randint(0, 0xFFFFFF))

def convert_uploaded_files(uploaded_files):
    temp_dir = tempfile.mkdtemp()
    output_dir = os.path.join(temp_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    # Save uploaded files
    for file in uploaded_files:
        file_path = os.path.join(temp_dir, file.name)
        with open(file_path, "wb") as f:
            f.write(file.read())
    
    # Convert .docx to PDF
    for filename in os.listdir(temp_dir):
        if filename.endswith(".docx"):
            docx_path = os.path.join(temp_dir, filename)
            try:
                convert(docx_path, output_dir)
                st.success(f"Converted {filename} to PDF!")
            except Exception as e:
                st.error(f"Failed to convert {filename}: {e}")
    
    # Create a ZIP archive for download
    zip_path = shutil.make_archive(os.path.join(temp_dir, "converted_pdfs"), 'zip', output_dir)
    return zip_path

# -----------------------
# Streamlit UI
# -----------------------

# Set page config
st.set_page_config(page_title="DOCX to PDF Converter", layout="centered")

# Random background color
bg_color = generate_random_color()
page_bg = f"""
<style>
body {{
    background-color: {bg_color};
    background-size: cover;
    position: relative;
}}
.watermark {{
    position: fixed;
    bottom: 10px;
    right: 10px;
    font-size: 70px;
    color: rgba(255, 255, 255, 0.1);
    transform: rotate(-30deg);
    z-index: 1;
    user-select: none;
}}
</style>
<div class="watermark">Trideb</div>
"""
st.markdown(page_bg, unsafe_allow_html=True)

# Title
st.title("📄 DOCX to PDF Converter")
st.write("Upload your `.docx` files below to convert them into PDFs.")

# File uploader
uploaded_files = st.file_uploader("Upload DOCX files", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Convert to PDF"):
        with st.spinner("Converting files..."):
            zip_path = convert_uploaded_files(uploaded_files)
        
        # Download button
        with open(zip_path, "rb") as f:
            st.download_button(
                label="📥 Download Converted PDFs (ZIP)",
                data=f,
                file_name="converted_pdfs.zip",
                mime="application/zip"
            )

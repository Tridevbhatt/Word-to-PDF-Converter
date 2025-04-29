import os
import streamlit as st
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from io import BytesIO
import random

# Random light background color
def random_light_color():
    return f"#{random.randint(200, 255):02x}{random.randint(200, 255):02x}{random.randint(200, 255):02x}"

# Convert .docx content to PDF using ReportLab
def convert_docx_to_pdf(docx_path, output_path):
    doc = Document(docx_path)
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    textobject = c.beginText(40, height - 50)
    textobject.setFont("Helvetica", 12)

    for para in doc.paragraphs:
        lines = para.text.split('\n')
        for line in lines:
            textobject.textLine(line)

    c.drawText(textobject)

    # Watermark
    c.setFont("Helvetica-Bold", 40)
    c.setFillGray(0.9, 0.2)  # Light gray watermark
    c.drawCentredString(width / 2, height / 2, "Trideb")

    c.save()

    with open(output_path, "wb") as f:
        f.write(buffer.getbuffer())

# Streamlit UI
st.set_page_config(page_title="DOCX to PDF Converter", layout="centered")
bg_color = random_light_color()
st.markdown(
    f"""
    <style>
        .stApp {{
            background-color: {bg_color};
        }}
        .watermark {{
            position: fixed;
            bottom: 10px;
            right: 20px;
            color: rgba(0, 0, 0, 0.1);
            font-size: 24px;
            font-weight: bold;
        }}
    </style>
    <div class="watermark">Trideb</div>
    """,
    unsafe_allow_html=True
)

st.title("📄 DOCX to PDF Batch Converter")

folder = st.text_input("Enter or paste the path to the folder containing .docx files:")

if st.button("Convert All"):
    if not folder or not os.path.isdir(folder):
        st.error("Please enter a valid folder path.")
    else:
        docx_files = [f for f in os.listdir(folder) if f.endswith(".docx")]
        total = len(docx_files)

        if total == 0:
            st.warning("No .docx files found in the folder.")
        else:
            progress = st.progress(0, text="Starting conversion...")
            for i, file in enumerate(docx_files):
                input_path = os.path.join(folder, file)
                output_path = os.path.join(folder, file.replace(".docx", ".pdf"))
                convert_docx_to_pdf(input_path, output_path)
                progress.progress((i + 1) / total, text=f"Converted {file}")

            st.success(f"✅ Successfully converted {total} file(s) to PDF.")

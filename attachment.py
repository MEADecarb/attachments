import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
import io
from PIL import Image
import zipfile
import os
import shutil

# Define functions
def set_font(paragraph, font_name, font_size):
    for run in paragraph.runs:
        run.font.name = font_name
        r = run._element
        rFonts = r.rPr.rFonts
        rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(font_size)

def set_document_font(doc, font_name='Times New Roman', font_size=12):
    for paragraph in doc.paragraphs:
        set_font(paragraph, font_name, font_size)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    set_font(paragraph, font_name, font_size)

def append_pdf_to_docx(pdf_paths, docx_path):
    try:
        doc = Document(docx_path)
    except:
        doc = Document()

    for pdf_path in pdf_paths:
        pdf_document = fitz.open(pdf_path)

        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()
            img_data = pix.tobytes()
            image = Image.open(io.BytesIO(img_data))
            image_stream = io.BytesIO()
            image.save(image_stream, format='PNG')
            doc.add_picture(image_stream, width=Inches(6))

            if page_num < len(pdf_document) - 1:
                doc.add_page_break()

    set_document_font(doc)
    doc.save(docx_path)

def process_zip(zip_path, pdf_paths, output_zip_path):
    extract_dir = "extracted_docs"
    os.makedirs(extract_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    for root, _, files in os.walk(extract_dir):
        for file in files:
            if file.endswith('.docx'):
                docx_path = os.path.join(root, file)
                append_pdf_to_docx(pdf_paths, docx_path)

    with zipfile.ZipFile(output_zip_path, 'w') as zip_ref:
        for root, _, files in os.walk(extract_dir):
            for file in files:
                full_path = os.path.join(root, file)
                zip_ref.write(full_path, os.path.relpath(full_path, extract_dir))

    shutil.rmtree(extract_dir)

# Streamlit app
st.title("Append PDFs to Word Documents")

# Upload zip file
zip_file = st.file_uploader("Upload ZIP file containing Word documents", type=["zip"])
# Upload multiple PDF files
pdf_files = st.file_uploader("Upload PDF files to append", type=["pdf"], accept_multiple_files=True)

if st.button("Process"):
    if zip_file and pdf_files:
        zip_path = zip_file.name
        output_zip_path = "output.zip"

        with open(zip_path, "wb") as f:
            f.write(zip_file.getbuffer())

        pdf_paths = []
        for pdf_file in pdf_files:
            pdf_path = pdf_file.name
            with open(pdf_path, "wb") as f:
                f.write(pdf_file.getbuffer())
            pdf_paths.append(pdf_path)

        process_zip(zip_path, pdf_paths, output_zip_path)
        
        with open(output_zip_path, "rb") as f:
            st.download_button("Download processed ZIP", f, file_name="processed_documents.zip")
        
        st.success(f"PDFs appended to Word documents in {output_zip_path} successfully.")
    else:
        st.error("Please upload both a ZIP file and at least one PDF file.")

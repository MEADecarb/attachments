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
import tempfile

# Define functions (keep all existing functions as they are)
# ...

# Streamlit app
st.title("Append PDFs or Word Documents to Word Documents")

# Add a text box with a link to MEA Attachment A
st.markdown("### MEA Attachment A available here: [MEA Attachment A](https://energy.maryland.gov/Pages/all-incentives.aspx)")

# Upload zip file
zip_file = st.file_uploader("Upload ZIP file containing Word or PDF documents", type=["zip"])

# Choose file type to append
file_type = st.selectbox("Select file type to append", ["PDF", "Word Document", "Both PDF and Word Document"])

# Upload multiple files
file_types = {"PDF": "pdf", "Word Document": "docx", "Both PDF and Word Document": ["pdf", "docx"]}
file_extension = file_types[file_type]
if isinstance(file_extension, list):
  files = st.file_uploader(f"Upload PDF and Word Document files to append", type=file_extension, accept_multiple_files=True)
else:
  files = st.file_uploader(f"Upload {file_type} files to append", type=[file_extension], accept_multiple_files=True)

# Option to upload a zip file containing Word or PDF documents to append
append_zip_file = st.file_uploader("Upload ZIP file containing documents to append", type=["zip"])

if st.button("Process"):
  if zip_file and (files or append_zip_file):
      with tempfile.TemporaryDirectory() as temp_dir:
          zip_path = os.path.join(temp_dir, zip_file.name)
          output_zip_path = os.path.join(temp_dir, "output.zip")

          with open(zip_path, "wb") as f:
              f.write(zip_file.getbuffer())

          file_paths = []
          if files:
              for file in files:
                  file_path = os.path.join(temp_dir, file.name)
                  with open(file_path, "wb") as f:
                      f.write(file.getbuffer())
                  file_paths.append(file_path)

          if append_zip_file:
              append_zip_path = os.path.join(temp_dir, append_zip_file.name)
              with open(append_zip_path, "wb") as f:
                  f.write(append_zip_file.getbuffer())
              process_appended_zip(append_zip_path, output_zip_path)
          else:
              if file_type == "Both PDF and Word Document":
                  process_zip(zip_path, [fp for fp in file_paths if fp.endswith('.pdf')], output_zip_path, 'pdf')
                  process_zip(zip_path, [fp for fp in file_paths if fp.endswith('.docx')], output_zip_path, 'docx')
              else:
                  process_zip(zip_path, file_paths, output_zip_path, file_extension)
          
          with open(output_zip_path, "rb") as f:
              st.download_button("Download processed ZIP", f, file_name="processed_documents.zip")
          
          st.success(f"{file_type}s appended to documents in the output ZIP file successfully.")
  else:
      st.error("Please upload both a ZIP file and at least one file to append.")

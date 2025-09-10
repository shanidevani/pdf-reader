# This is a Streamlit application for the PDF converter.
# To run this app, first make sure you have installed all the required libraries:
# pip install streamlit PyMuPDF pandas pdfplumber python-docx openpyxl
# Then, save this file and run the following command in your terminal:
# streamlit run streamlit_converter.py

import streamlit as st
import fitz
import pandas as pd
import docx
import pdfplumber
import os
import tempfile

st.set_page_config(page_title="PDF Converter", layout="centered")

# Set up the title and a brief description
st.title("PDF Converter App")
st.markdown("Easily convert your PDF files to Text, Excel, or Word.")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    # Use a temporary file to save the uploaded PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.read())
        pdf_path = tmp_file.name

    # Conversion selection
    conversion_type = st.radio(
        "Select the output format:",
        ('Text (.txt)', 'Excel (.xlsx)', 'Word (.docx)')
    )

    if st.button("Convert"):
        with st.spinner(f"Converting to {conversion_type}..."):
            output_file_path = ""
            try:
                if conversion_type == 'Text (.txt)':
                    output_file_path = os.path.join(tempfile.gettempdir(), "output.txt")
                    # PDF to Text conversion logic
                    pdf_document = fitz.open(pdf_path)
                    text_content = ""
                    for page_num in range(len(pdf_document)):
                        page = pdf_document.load_page(page_num)
                        text_content += page.get_text()
                    with open(output_file_path, "w", encoding="utf-8") as file:
                        file.write(text_content)
                    
                elif conversion_type == 'Excel (.xlsx)':
                    output_file_path = os.path.join(tempfile.gettempdir(), "output.xlsx")
                    # PDF to Excel conversion logic
                    all_tables = []
                    with pdfplumber.open(pdf_path) as pdf:
                        for page in pdf.pages:
                            tables = page.extract_tables()
                            for table in tables:
                                if table and table[0]:
                                    all_tables.append(pd.DataFrame(table[1:], columns=table[0]))
                    
                    if not all_tables:
                        st.warning("⚠️ No tables were found in the PDF. The Excel file will be empty.")
                    else:
                        combined_df = pd.concat(all_tables)
                        combined_df.to_excel(output_file_path, index=False)

                elif conversion_type == 'Word (.docx)':
                    output_file_path = os.path.join(tempfile.gettempdir(), "output.docx")
                    # PDF to Word conversion logic
                    pdf_document = fitz.open(pdf_path)
                    text_content = ""
                    for page_num in range(len(pdf_document)):
                        page = pdf_document.load_page(page_num)
                        text_content += page.get_text()
                    
                    doc = docx.Document()
                    doc.add_paragraph(text_content)
                    doc.save(output_file_path)

                # Provide a download button for the converted file
                with open(output_file_path, "rb") as f:
                    st.success("✅ Conversion complete!")
                    st.download_button(
                        label="Download Converted File",
                        data=f,
                        file_name=os.path.basename(output_file_path),
                        mime="application/octet-stream"
                    )
            except Exception as e:
                st.error(f"An error occurred during conversion: {e}")
            finally:
                # Clean up the temporary PDF file
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                # Note: The output file is not deleted so it can be downloaded.

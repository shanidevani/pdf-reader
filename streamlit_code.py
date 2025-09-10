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
import traceback

st.set_page_config(page_title="PDF Converter", layout="centered")

# --- Function Definitions for Conversion ---

def convert_to_text(pdf_path):
    """Converts a PDF to plain text."""
    text_content = ""
    try:
        with fitz.open(pdf_path) as pdf_document:
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text_content += page.get_text()
        return text_content
    except Exception as e:
        st.error(f"Failed to extract text: {e}")
        return None

def convert_to_excel(pdf_path):
    """
    Extracts tables from a PDF to an Excel file.
    If no tables are found, it converts the text content to a single-column Excel file.
    """
    all_tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0]:
                        # Create a DataFrame from the table, using the first row as headers
                        all_tables.append(pd.DataFrame(table[1:], columns=table[0]))
    except Exception as e:
        st.error(f"Failed to extract tables: {e}")
        return None

    if all_tables:
        st.info("✅ Tables found and extracted from the PDF.")
        return pd.concat(all_tables)
    else:
        # If no tables were found, convert the entire text content to a DataFrame
        st.warning("⚠️ No tables were found in the PDF. Converting the entire text content to an Excel file.")
        text_content = convert_to_text(pdf_path)
        if text_content:
            return pd.DataFrame([text_content], columns=["Text Content from PDF"])
        return pd.DataFrame() # Return an empty DataFrame if no text is found

def convert_to_word(pdf_path):
    """
    Converts a PDF to a Word document, preserving only the text content.
    Note: Due to library limitations, complex formatting, images, and layout
    will not be preserved.
    """
    doc = docx.Document()
    try:
        with fitz.open(pdf_path) as pdf_document:
            text_content = ""
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text_content += page.get_text()
                # Adding a page break for better document structure
                if page_num < len(pdf_document) - 1:
                    text_content += "\n\n--- Page Break ---\n\n"
            
            # Add text as a single paragraph
            doc.add_paragraph(text_content)
            return doc
    except Exception as e:
        st.error(f"Failed to convert to Word: {e}")
        return None

# --- Main Streamlit App Logic ---

st.title("PDF Converter App")
st.markdown("Easily convert your PDF files to Text, Excel, or Word.")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file:
    conversion_type = st.radio(
        "Select the output format:",
        ('Text (.txt)', 'Excel (.xlsx)', 'Word (.docx)')
    )

    if st.button("Convert"):
        # Use a temporary file to handle the uploaded file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.read())
            pdf_path = tmp_file.name

        output_file_path = ""
        try:
            with st.spinner(f"Converting to {conversion_type}..."):
                if conversion_type == 'Text (.txt)':
                    output_file_path = "output.txt"
                    text_data = convert_to_text(pdf_path)
                    if text_data is not None:
                        with open(output_file_path, "w", encoding="utf-8") as f:
                            f.write(text_data)
                        st.success("✅ Conversion to Text complete!")

                elif conversion_type == 'Excel (.xlsx)':
                    output_file_path = "output.xlsx"
                    df = convert_to_excel(pdf_path)
                    if df is not None:
                        df.to_excel(output_file_path, index=False)
                        st.success("✅ Conversion to Excel complete!")

                elif conversion_type == 'Word (.docx)':
                    st.info("ℹ️ Please note: Word conversion currently only extracts text and does not preserve original formatting like images, fonts, or layout.")
                    output_file_path = "output.docx"
                    doc_obj = convert_to_word(pdf_path)
                    if doc_obj is not None:
                        doc_obj.save(output_file_path)
                        st.success("✅ Conversion to Word complete!")

            # If an output file was created, provide a download button
            if os.path.exists(output_file_path):
                with open(output_file_path, "rb") as f:
                    st.download_button(
                        label="Download Converted File",
                        data=f,
                        file_name=os.path.basename(output_file_path),
                        mime="application/octet-stream"
                    )

        except Exception as e:
            st.error(f"An unexpected error occurred during conversion. Please try again.")
            st.code(traceback.format_exc()) # This shows the full traceback for debugging

        finally:
            # Clean up temporary files
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            if output_file_path and os.path.exists(output_file_path):
                try:
                    os.remove(output_file_path)
                except OSError as e:
                    # File might be in use, which is normal for Streamlit's download button
                    print(f"Error removing file: {e}")

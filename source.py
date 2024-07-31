import streamlit as st
from PIL import Image
import pytesseract
import openpyxl
from openpyxl import Workbook
import subprocess

# Specify the batch file path instead of the Tesseract executable path
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Athar\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# Set up the Streamlit interfacest
st.set_page_config(layout='wide')
st.title("Hello Gina welcome to GINAS :wave:")
st.header("Graphical . Image-to . Note . Analysing . Software")
st.write("This software was created on the 31st of August from Attie to Gina Bina")
st.write("Upload images to extract text and save it to an Excel file.")

uploaded_files = st.file_uploader("Choose image files", accept_multiple_files=True, type=["jpg", "jpeg", "png"])

if uploaded_files:
    extracted_texts = []
    for uploaded_file in uploaded_files:
        image = Image.open(uploaded_file)
        text = pytesseract.image_to_string(image)
        extracted_texts.append((uploaded_file.name, text))
        #st.image(image, caption=uploaded_file.name, use_column_width=True)
        st.write(text)
        st.write('------')

    if st.button("Save to Excel"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Text"

        # Write headers
        ws.append(["Filename", "Extracted Text"])

        # Write extracted text
        for file_name, text in extracted_texts:
            ws.append([file_name, text])

        # Save the Excel file
        excel_filename = "extracted_texts.xlsx"
        wb.save(excel_filename)
        st.success(f"Data saved to {excel_filename}")
        st.download_button(
            label="Download Excel file",
            data=open(excel_filename, "rb"),
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

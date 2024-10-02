import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
import re
import pytesseract
from PIL import Image
import pdf2image
import math
import cv2
import numpy as np
import tempfile

# Set the path to Tesseract OCR executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to extract text from PDF using PyPDF2 (for text-based PDFs)
def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"
    return full_text

# Function to preprocess the image for better OCR
def preprocess_image(image):
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresh_image = cv2.threshold(gray_image, 150, 255, cv2.THRESH_BINARY)
    return thresh_image

# Function to convert PDF to images and use OCR to extract text
def extract_text_using_ocr(pdf_path):
    images = pdf2image.convert_from_path(pdf_path, dpi=300)
    full_text = ""
    for img in images:
        open_cv_image = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        processed_image = preprocess_image(open_cv_image)
        text = pytesseract.image_to_string(processed_image)
        full_text += text + "\n"
    return full_text

# General function to extract data from text using regex
def extract_data_from_text(text):
    data = []
    pattern = re.compile(r"(0801[A-Z\d]*[A-Z]?)\s+([A-Za-z\s]+?)\s+(\d+(\.\d+)?|A|None|Absent)", re.IGNORECASE)
    matches = pattern.findall(text)
    
    for match in matches:
        enrollment_no = match[0].strip()
        name = match[1].strip()
        marks_or_status = match[2].strip() if match[2] else "None"
        
        if marks_or_status.replace('.', '', 1).isdigit():
            marks = math.ceil(float(marks_or_status))
            status = "Present"
        elif marks_or_status.lower() in ["a", "absent", "none"]:
            marks = None
            status = "Absent"
        else:
            marks = None
            status = "Unknown"
        
        data.append((enrollment_no, name, marks, status))
    
    return data

# Function to process the data into categories
def process_data(data):
    df = pd.DataFrame(data, columns=['Enrollment No', 'Name', 'Marks', 'Status'])
    df.dropna(subset=['Enrollment No', 'Name'], inplace=True)
    
    df.loc[(df['Marks'].notnull()) & (df['Marks'] >= 7), 'Status'] = 'Pass'
    df.loc[(df['Marks'].notnull()) & (df['Marks'] < 7), 'Status'] = 'Fail'
    df['Status'] = df['Status'].fillna('Absent')
    
    passed = df[df['Status'] == 'Pass']
    failed = df[df['Status'] == 'Fail']
    absent = df[df['Status'] == 'Absent']
    
    return passed, failed, absent

# Function to generate Excel sheets for passed, failed, and absent students
def generate_excel(passed, failed, absent):
    # Create a named temporary file to hold the Excel file
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_file.close()  # Close the file so Pandas can write to it

    with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
        if not passed.empty:
            passed.to_excel(writer, sheet_name="Passed Students", index=False)
        if not failed.empty:
            failed.to_excel(writer, sheet_name="Failed Students", index=False)
        if not absent.empty:
            absent.to_excel(writer, sheet_name="Absent Students", index=False)

    return temp_file.name  # Return the path to the Excel file

# Streamlit App
def main():
    st.title("Student Marks Categorization App")
    
    uploaded_file = st.file_uploader("Upload a PDF or Image file", type=["pdf", "jpg", "jpeg", "png"])
    
    if uploaded_file:
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=uploaded_file.name.split('.')[-1]) as temp_file:
            temp_file.write(uploaded_file.getbuffer())
            file_path = temp_file.name

        st.write("Processing the file...")

        # Try to extract text from the PDF or image
        if uploaded_file.name.endswith(".pdf"):
            text = extract_text_from_pdf(file_path)
            if not text.strip():
                st.write("No text extracted using PyPDF2, attempting OCR...")
                text = extract_text_using_ocr(file_path)
        else:
            image = Image.open(uploaded_file)
            text = pytesseract.image_to_string(image)

        if not text.strip():
            st.error("No data extracted. Please check the file format.")
            return

        data = extract_data_from_text(text)
        passed, failed, absent = process_data(data)

        # Displaying summary of results
        st.write(f"Total Students: {len(data)}")
        st.write(f"Passed: {len(passed)}")
        st.write(f"Failed: {len(failed)}")
        st.write(f"Absent: {len(absent)}")

        # Offer option to download categorized Excel sheet
        if st.button("Generate Excel"):
            try:
                output_path = generate_excel(passed, failed, absent)
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="Download Excel File",
                        data=f,
                        file_name="Students_Categorized.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("Excel file created successfully.")
            except Exception as e:
                st.error(f"Error generating Excel file: {e}")

if __name__ == "__main__":
    main()

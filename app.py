import streamlit as st
import pdfplumber
from PIL import Image
import pytesseract
import os
import re
from openpyxl import load_workbook

# =========================================
# TESSERACT SETTINGS
# =========================================

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

# =========================================
# APP TITLE
# =========================================

st.title("⚡ Electricity Bill Automation")

# =========================================
# FILE UPLOAD
# =========================================

uploaded_file = st.file_uploader(
    "Upload Electricity Bill",
    type=["pdf", "png", "jpg", "jpeg"]
)

# =========================================
# PDF TEXT EXTRACTION
# =========================================

def extract_text_from_pdf(pdf_file):

    text = ""

    with pdfplumber.open(pdf_file) as pdf:

        for page in pdf.pages:

            extracted = page.extract_text()

            if extracted:
                text += extracted

    return text

# =========================================
# IMAGE OCR EXTRACTION
# =========================================

def extract_text_from_image(image_file):

    image = Image.open(image_file)

    text = pytesseract.image_to_string(image)

    return text

# =========================================
# BILL DATA EXTRACTION
# =========================================

def extract_bill_data(text):

    data = {}

    # Consumer Number
    consumer_match = re.search(r'(\d{12})', text)

    if consumer_match:
        data["consumer_number"] = consumer_match.group(1)
    else:
        data["consumer_number"] = "0"

    # Bill Amount
    amount_match = re.search(r'Rs[,.\s]*([0-9,]+\.?[0-9]*)', text)

    if amount_match:
        data["bill_amount"] = amount_match.group(1)
    else:
        data["bill_amount"] = "0"

    # Units Consumed
    units_match = re.search(r'\s(\d{1,4})\s0\s\d{1,4}', text)

    if units_match:
        data["units_consumed"] = units_match.group(1)
    else:
        data["units_consumed"] = "0"

    # Sanctioned Load
    load_match = re.search(r'(\d+\.\d+\s*KW)', text)

    if load_match:
        data["sanctioned_load"] = load_match.group(1)
    else:
        data["sanctioned_load"] = "0"

    return data

# =========================================
# EXCEL AUTO FILL
# =========================================

def fill_excel(data):

    workbook = load_workbook("template.xlsx")

    sheet = workbook.active

    # =====================================
    # BASIC DETAILS
    # =====================================

    # Consumer Number
    sheet["D2"] = data["consumer_number"]

    # Sanctioned Load
    sheet["D4"] = data["sanctioned_load"]

    # =====================================
    # JANUARY 2026 DATA
    # =====================================

    # Units
    units = int(data["units_consumed"])

    sheet["D20"] = units

    # Bill Amount
    amount = float(
        str(data["bill_amount"]).replace(",", "")
    )

    sheet["E20"] = amount

    # Unit Cost
    if units > 0:
        unit_cost = amount / units
    else:
        unit_cost = 0

    sheet["F20"] = round(unit_cost, 2)

    # =====================================
    # SAVE FILE
    # =====================================

    output_file = "filled_output.xlsx"

    workbook.save(output_file)

    return output_file

# =========================================
# MAIN APP
# =========================================

if uploaded_file:

    st.success("File Uploaded Successfully")

    # PDF FILE
    if uploaded_file.type == "application/pdf":

        extracted_text = extract_text_from_pdf(uploaded_file)

    # IMAGE FILE
    else:

        extracted_text = extract_text_from_image(uploaded_file)

    # EXTRACT DATA
    bill_data = extract_bill_data(extracted_text)

    # GENERATE EXCEL
    output_file = fill_excel(bill_data)

    st.success("Excel File Generated Successfully")

    # DOWNLOAD BUTTON
    with open(output_file, "rb") as file:

        st.download_button(
            label="Download Filled Excel",
            data=file,
            file_name="filled_output.xlsx"
        )
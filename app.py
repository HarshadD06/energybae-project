import streamlit as st
import pdfplumber
from PIL import Image
import pytesseract
import os
import re
import platform
from openpyxl import load_workbook

# =========================================
# PAGE CONFIG
# =========================================

st.set_page_config(
    page_title="Electricity Bill Automation",
    page_icon="⚡",
    layout="centered"
)

# =========================================
# CUSTOM CSS
# =========================================

st.markdown("""
<style>

.stApp {
    background-color: #0e1117;
    color: white;
}

.main-title {
    text-align:center;
    font-size:42px;
    font-weight:bold;
    color:white;
    margin-bottom:10px;
}

.sub-title {
    text-align:center;
    color:#b0b0b0;
    margin-bottom:30px;
}

.result-box{
    background:#1c1f26;
    padding:20px;
    border-radius:12px;
    border:1px solid #2e3440;
    margin-top:20px;
}

.stButton>button {
    background:#1f77ff;
    color:white;
    border:none;
    border-radius:10px;
    padding:10px 20px;
    font-weight:bold;
}

.stDownloadButton>button {
    background:#00b894;
    color:white;
    border:none;
    border-radius:10px;
    padding:10px 20px;
    font-weight:bold;
    width:100%;
}

</style>
""", unsafe_allow_html=True)

# =========================================
# WINDOWS OCR SETUP
# =========================================

if platform.system() == "Windows":

    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

# =========================================
# TITLE
# =========================================

st.markdown(
    '<div class="main-title">⚡ Electricity Bill Automation</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">Upload electricity bill PDF or image and generate Excel automatically</div>',
    unsafe_allow_html=True
)

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

    try:

        with pdfplumber.open(pdf_file) as pdf:

            for page in pdf.pages:

                extracted = page.extract_text()

                if extracted:
                    text += extracted + "\n"

    except Exception as e:

        st.error(f"PDF Error: {e}")

    return text

# =========================================
# IMAGE OCR
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

    patterns = {

        "consumer_number": [
            r"Consumer\s*No\.?\s*[:\-]?\s*(\d+)",
            r"Customer\s*ID\s*[:\-]?\s*(\d+)",
            r"(\d{12})"
        ],

        "bill_amount": [
            r"Bill\s*Amount\s*[:\-]?\s*₹?\s*([\d,.]+)",
            r"Current\s*Bill\s*[:\-]?\s*₹?\s*([\d,.]+)",
            r"Total\s*Amount\s*[:\-]?\s*₹?\s*([\d,.]+)",
            r"Rs[,.\s]*([0-9,]+\.?[0-9]*)"
        ],

        "units_consumed": [
            r"Units\s*Consumed\s*[:\-]?\s*(\d+)",
            r"Consumption\s*[:\-]?\s*(\d+)"
        ],

        "sanctioned_load": [
            r"Sanctioned\s*Load\s*[:\-]?\s*([\d.]+)",
            r"Load\s*[:\-]?\s*([\d.]+)",
            r"(\d+\.\d+\s*KW)"
        ]
    }

    for key, regex_list in patterns.items():

        value = "0"

        for pattern in regex_list:

            match = re.search(pattern, text, re.IGNORECASE)

            if match:
                value = match.group(1)
                break

        data[key] = value

    return data

# =========================================
# FILL EXCEL
# =========================================

def fill_excel(data):

    workbook = load_workbook("template.xlsx")

    sheet = workbook.active

    # Consumer Number
    sheet["D2"] = data["consumer_number"]

    # Load
    sheet["D4"] = data["sanctioned_load"]

    # Units
    try:

        units = int(
            str(data["units_consumed"]).replace(",", "")
        )

    except:

        units = 0

    sheet["D20"] = units

    # Bill Amount
    try:

        amount = float(
            str(data["bill_amount"]).replace(",", "")
        )

    except:

        amount = 0

    sheet["E20"] = amount

    # Unit Cost
    if units > 0:

        unit_cost = amount / units

    else:

        unit_cost = 0

    sheet["F20"] = round(unit_cost, 2)

    # Save Output
    output_file = "filled_output.xlsx"

    workbook.save(output_file)

    return output_file

# =========================================
# MAIN PROCESS
# =========================================

if uploaded_file:

    st.success("✅ File Uploaded Successfully")

    extracted_text = ""

    # PDF
    if uploaded_file.type == "application/pdf":

        extracted_text = extract_text_from_pdf(uploaded_file)

    # IMAGE
    else:

        extracted_text = extract_text_from_image(uploaded_file)

    # VIEW TEXT
    with st.expander("View Extracted Text"):

        st.text(extracted_text)

    # EXTRACT DATA
    bill_data = extract_bill_data(extracted_text)

    # RESULT BOX
    st.markdown('<div class="result-box">', unsafe_allow_html=True)

    st.subheader("📊 Extracted Bill Data")

    st.write(f"**Consumer Number:** {bill_data['consumer_number']}")
    st.write(f"**Bill Amount:** ₹ {bill_data['bill_amount']}")
    st.write(f"**Units Consumed:** {bill_data['units_consumed']}")
    st.write(f"**Sanctioned Load:** {bill_data['sanctioned_load']}")

    st.markdown('</div>', unsafe_allow_html=True)

    # GENERATE EXCEL
    output_file = fill_excel(bill_data)

    st.success("✅ Excel File Generated Successfully")

    # DOWNLOAD BUTTON
    with open(output_file, "rb") as file:

        st.download_button(
            label="⬇ Download Filled Excel",
            data=file,
            file_name="filled_output.xlsx"
        )
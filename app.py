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
    page_title="Solar Load Calculator",
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

.stDownloadButton>button{
    background:#00b894;
    color:white;
    border:none;
    border-radius:10px;
    padding:12px;
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
    '<div class="main-title">⚡ Solar Load Calculator</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">Upload electricity bill and generate solar recommendation automatically</div>',
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

    try:

        image = Image.open(image_file)

        text = pytesseract.image_to_string(image)

        return text

    except:

        return ""

# =========================================
# DATA EXTRACTION
# =========================================

def extract_bill_data(text):

    data = {}

    text = text.replace("\n", " ")

    # Consumer Number
    consumer_match = re.search(r'(\d{12})', text)

    if consumer_match:
        data["consumer_number"] = consumer_match.group(1)
    else:
        data["consumer_number"] = ""

    # Bill Amount
    amount_match = re.search(
        r'([\d,]+\.\d{2})',
        text
    )

    if amount_match:
        data["bill_amount"] = amount_match.group(1)
    else:
        data["bill_amount"] = ""

    # Units
    units_match = re.search(
        r'Units\s*Consumed\s*[:\-]?\s*(\d+)',
        text,
        re.IGNORECASE
    )

    if units_match:
        data["units_consumed"] = units_match.group(1)
    else:
        data["units_consumed"] = ""

    # Load
    load_match = re.search(
        r'([\d.]+)\s*KW',
        text,
        re.IGNORECASE
    )

    if load_match:
        data["sanctioned_load"] = load_match.group(1)
    else:
        data["sanctioned_load"] = ""

    return data

# =========================================
# EXCEL GENERATION
# =========================================

def fill_excel(data):

    workbook = load_workbook("template.xlsx")

    sheet = workbook.active

    # =========================================
    # VALUES
    # =========================================

    try:
        units = int(
            str(data["units_consumed"]).replace(",", "")
        )
    except:
        units = 0

    try:
        amount = float(
            str(data["bill_amount"]).replace(",", "")
        )
    except:
        amount = 0

    # =========================================
    # CALCULATIONS
    # =========================================

    if units > 0:
        unit_cost = amount / units
    else:
        unit_cost = 0

    solar_kw = units / 120

    solar_panels = solar_kw / 0.55

    solar_capacity = round(solar_kw)

    # =========================================
    # FILL EXCEL
    # =========================================

    sheet["D2"] = data["consumer_number"]
    sheet["D4"] = data["sanctioned_load"]

    sheet["C20"] = units
    sheet["D20"] = amount
    sheet["E20"] = round(unit_cost, 2)

    sheet["C22"] = units
    sheet["C23"] = round(solar_kw, 2)
    sheet["C24"] = round(solar_panels, 2)
    sheet["C25"] = solar_capacity
    sheet["C26"] = round(solar_panels)

    # =========================================
    # SAVE
    # =========================================

    output_file = "filled_output.xlsx"

    workbook.save(output_file)

    return output_file

# =========================================
# MAIN PROCESS
# =========================================

if uploaded_file:

    st.success("✅ File Uploaded Successfully")

    extracted_text = ""

    # =========================================
    # PDF
    # =========================================

    if uploaded_file.type == "application/pdf":

        extracted_text = extract_text_from_pdf(uploaded_file)

        if len(extracted_text.strip()) < 20:

            st.warning(
                "Scanned PDF detected. Auto extraction may not work correctly."
            )

    # =========================================
    # IMAGE
    # =========================================

    else:

        extracted_text = extract_text_from_image(uploaded_file)

    # =========================================
    # SHOW TEXT
    # =========================================

    with st.expander("View Extracted Text"):

        st.text(extracted_text)

    # =========================================
    # EXTRACT DATA
    # =========================================

    bill_data = extract_bill_data(extracted_text)

    st.markdown('<div class="result-box">', unsafe_allow_html=True)

    st.subheader("📊 Extracted Bill Data")

    # =========================================
    # EDITABLE INPUTS
    # =========================================

    consumer_number = st.text_input(
        "Consumer Number",
        bill_data["consumer_number"]
    )

    bill_amount = st.text_input(
        "Bill Amount",
        bill_data["bill_amount"]
    )

    units_consumed = st.text_input(
        "Units Consumed",
        bill_data["units_consumed"]
    )

    sanctioned_load = st.text_input(
        "Sanctioned Load",
        bill_data["sanctioned_load"]
    )

    # =========================================
    # UPDATE DATA
    # =========================================

    bill_data["consumer_number"] = consumer_number
    bill_data["bill_amount"] = bill_amount
    bill_data["units_consumed"] = units_consumed
    bill_data["sanctioned_load"] = sanctioned_load

    # =========================================
    # CALCULATIONS
    # =========================================

    try:
        units = int(
            str(units_consumed).replace(",", "")
        )
    except:
        units = 0

    try:
        amount = float(
            str(bill_amount).replace(",", "")
        )
    except:
        amount = 0

    solar_kw = round(units / 120, 2)

    monthly_savings = round(units * 8, 2)

    annual_savings = round(monthly_savings * 12, 2)

    # =========================================
    # RESULTS
    # =========================================

    st.subheader("☀ Solar Recommendation")

    st.write(f"**Recommended Solar Capacity:** {solar_kw} kW")
    st.write(f"**Estimated Monthly Savings:** ₹ {monthly_savings}")
    st.write(f"**Estimated Annual Savings:** ₹ {annual_savings}")

    st.markdown('</div>', unsafe_allow_html=True)

    # =========================================
    # GENERATE EXCEL
    # =========================================

    output_file = fill_excel(bill_data)

    st.success("✅ Excel File Generated Successfully")

    # =========================================
    # DOWNLOAD
    # =========================================

    with open(output_file, "rb") as file:

        st.download_button(
            label="⬇ Download Filled Excel",
            data=file,
            file_name="filled_output.xlsx"
        )
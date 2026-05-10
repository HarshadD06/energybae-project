import streamlit as st
import pdfplumber
from PIL import Image
import pytesseract
import os
import re
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
# CUSTOM UI
# =========================================

st.markdown("""
<style>

.main {
    background-color: #f4f7fb;
}

.title {
    text-align:center;
    color:#1f4e79;
    font-size:42px;
    font-weight:bold;
    margin-bottom:10px;
}

.subtitle{
    text-align:center;
    color:gray;
    margin-bottom:30px;
}

.card {
    background:white;
    padding:30px;
    border-radius:18px;
    box-shadow:0 0 20px rgba(0,0,0,0.08);
}

.stButton>button {
    background:#1f4e79;
    color:white;
    border:none;
    padding:12px 20px;
    border-radius:10px;
    width:100%;
    font-size:16px;
    font-weight:bold;
}

.stDownloadButton>button{
    background:#198754;
    color:white;
    border:none;
    padding:12px 20px;
    border-radius:10px;
    width:100%;
    font-size:16px;
    font-weight:bold;
}

.result-box{
    background:#f8f9fa;
    padding:15px;
    border-radius:10px;
    margin-top:15px;
    border-left:5px solid #1f4e79;
}

</style>
""", unsafe_allow_html=True)

# =========================================
# TESSERACT CONFIG
# =========================================

# LOCAL WINDOWS
# Uncomment locally only

# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# =========================================
# TITLE
# =========================================

st.markdown(
    '<div class="title">⚡ Electricity Bill Automation</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">Upload electricity bill PDF or image and generate Excel automatically</div>',
    unsafe_allow_html=True
)

# =========================================
# MAIN CARD
# =========================================

with st.container():

    st.markdown('<div class="card">', unsafe_allow_html=True)

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
                r"Consumption\s*[:\-]?\s*(\d+)",
                r"\s(\d{1,4})\s0\s\d{1,4}"
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

        # Output
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

        # DEBUG TEXT
        with st.expander("View Extracted Text"):

            st.text(extracted_text)

        # Extract Data
        bill_data = extract_bill_data(extracted_text)

        # Results UI
        st.markdown(
            '<div class="result-box">',
            unsafe_allow_html=True
        )

        st.subheader("📊 Extracted Bill Data")

        st.write(f"**Consumer Number:** {bill_data['consumer_number']}")
        st.write(f"**Bill Amount:** ₹ {bill_data['bill_amount']}")
        st.write(f"**Units Consumed:** {bill_data['units_consumed']}")
        st.write(f"**Sanctioned Load:** {bill_data['sanctioned_load']}")

        st.markdown('</div>', unsafe_allow_html=True)

        # Generate Excel
        output_file = fill_excel(bill_data)

        st.success("✅ Excel File Generated Successfully")

        # Download Button
        with open(output_file, "rb") as file:

            st.download_button(
                label="⬇ Download Filled Excel",
                data=file,
                file_name="filled_output.xlsx"
            )

    st.markdown('</div>', unsafe_allow_html=True)
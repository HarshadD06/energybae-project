import streamlit as st
import pdfplumber
import re
import os
from openpyxl import load_workbook

# =========================================
# PAGE SETTINGS
# =========================================

st.set_page_config(
    page_title="Electricity Bill Automation",
    page_icon="⚡",
    layout="centered"
)

# =========================================
# APP TITLE
# =========================================

st.title("⚡ Electricity Bill Automation")
st.write("Upload Electricity Bill PDF and Generate Excel Automatically")

# =========================================
# FILE UPLOAD
# =========================================

uploaded_file = st.file_uploader(
    "Upload Electricity Bill PDF",
    type=["pdf"]
)

# =========================================
# PDF TEXT EXTRACTION
# =========================================

def extract_text_from_pdf(pdf_file):

    text = ""

    try:

        st.info("📄 Reading PDF...")

        with pdfplumber.open(pdf_file) as pdf:

            total_pages = len(pdf.pages)

            st.write(f"Total Pages Found: {total_pages}")

            # Process only first 3 pages
            # Faster for Render free tier

            max_pages = min(total_pages, 3)

            for i in range(max_pages):

                st.write(f"Processing Page {i+1}...")

                page = pdf.pages[i]

                extracted = page.extract_text()

                if extracted:

                    text += extracted + "\n"

        st.success("✅ PDF Reading Completed")

    except Exception as e:

        st.error(f"PDF Reading Error: {e}")

    return text

# =========================================
# BILL DATA EXTRACTION
# =========================================

def extract_bill_data(text):

    data = {
        "consumer_number": "0",
        "bill_amount": "0",
        "units_consumed": "0",
        "sanctioned_load": "0"
    }

    # =====================================
    # CONSUMER NUMBER
    # =====================================

    consumer_patterns = [

        r'Consumer\s*No\.?\s*[:\-]?\s*(\d+)',

        r'Consumer\s*Number\s*[:\-]?\s*(\d+)',

        r'Customer\s*No\.?\s*[:\-]?\s*(\d+)',

        r'(\d{12})'
    ]

    for pattern in consumer_patterns:

        match = re.search(pattern, text, re.IGNORECASE)

        if match:

            data["consumer_number"] = match.group(1)

            break

    # =====================================
    # BILL AMOUNT
    # =====================================

    amount_patterns = [

        r'Current\s*Bill\s*Amount\s*[:\-]?\s*Rs\.?\s*([0-9,]+\.?[0-9]*)',

        r'Bill\s*Amount\s*[:\-]?\s*Rs\.?\s*([0-9,]+\.?[0-9]*)',

        r'Total\s*Amount\s*[:\-]?\s*Rs\.?\s*([0-9,]+\.?[0-9]*)',

        r'Rs\.?\s*([0-9,]+\.?[0-9]*)'
    ]

    for pattern in amount_patterns:

        match = re.search(pattern, text, re.IGNORECASE)

        if match:

            data["bill_amount"] = match.group(1)

            break

    # =====================================
    # UNITS CONSUMED
    # =====================================

    unit_patterns = [

        r'Units\s*Consumed\s*[:\-]?\s*(\d+)',

        r'Consumption\s*[:\-]?\s*(\d+)',

        r'(\d+)\s*Units'
    ]

    for pattern in unit_patterns:

        match = re.search(pattern, text, re.IGNORECASE)

        if match:

            data["units_consumed"] = match.group(1)

            break

    # =====================================
    # SANCTIONED LOAD
    # =====================================

    load_patterns = [

        r'Sanctioned\s*Load\s*[:\-]?\s*([0-9.]+\s*KW)',

        r'Load\s*[:\-]?\s*([0-9.]+\s*KW)',

        r'([0-9.]+\s*KW)'
    ]

    for pattern in load_patterns:

        match = re.search(pattern, text, re.IGNORECASE)

        if match:

            data["sanctioned_load"] = match.group(1)

            break

    return data

# =========================================
# EXCEL AUTO FILL
# =========================================

def fill_excel(data):

    try:

        if not os.path.exists("template.xlsx"):

            st.error("template.xlsx file not found")

            return None

        workbook = load_workbook("template.xlsx")

        sheet = workbook.active

        # =====================================
        # BASIC DETAILS
        # =====================================

        sheet["D2"] = data["consumer_number"]

        sheet["D4"] = data["sanctioned_load"]

        # =====================================
        # BILL DATA
        # =====================================

        try:

            units = int(
                str(data["units_consumed"]).replace(",", "")
            )

        except:

            units = 0

        sheet["D20"] = units

        try:

            amount = float(
                str(data["bill_amount"]).replace(",", "")
            )

        except:

            amount = 0

        sheet["E20"] = amount

        # =====================================
        # UNIT COST
        # =====================================

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

    except Exception as e:

        st.error(f"Excel Error: {e}")

        return None

# =========================================
# MAIN APP
# =========================================

if uploaded_file:

    st.success("✅ File Uploaded Successfully")

    # =====================================
    # EXTRACT PDF TEXT
    # =====================================

    extracted_text = extract_text_from_pdf(uploaded_file)

    # =====================================
    # SHOW RAW TEXT
    # =====================================

    st.subheader("📄 Extracted Raw Text")

    if extracted_text.strip():

        st.text(extracted_text[:5000])

    else:

        st.error(
            "❌ No readable text found in PDF.\n"
            "Use original electricity bill PDF.\n"
            "Scanned/WhatsApp PDFs may not work."
        )

    # =====================================
    # EXTRACT DATA
    # =====================================

    bill_data = extract_bill_data(extracted_text)

    st.subheader("📊 Extracted Data")

    st.json(bill_data)

    # =====================================
    # GENERATE EXCEL
    # =====================================

    output_file = fill_excel(bill_data)

    # =====================================
    # DOWNLOAD BUTTON
    # =====================================

    if output_file:

        st.success("✅ Excel File Generated Successfully")

        with open(output_file, "rb") as file:

            st.download_button(
                label="⬇ Download Filled Excel",
                data=file,
                file_name="filled_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
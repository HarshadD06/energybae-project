import os
import re
import pdfplumber
import pandas as pd

from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename

# =========================
# CONFIG
# =========================

UPLOAD_FOLDER = "uploads"
EXCEL_FOLDER = "excel"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXCEL_FOLDER, exist_ok=True)

app = Flask(__name__)

# =========================
# PDF TEXT EXTRACTION
# =========================

def extract_text_from_pdf(pdf_path):

    text = ""

    try:
        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                extracted = page.extract_text()

                if extracted:
                    text += extracted + "\n"

    except Exception as e:
        print("PDF Error:", e)

    return text

# =========================
# DATA EXTRACTION
# =========================

def extract_bill_data(text):

    print("\n========== PDF TEXT ==========\n")
    print(text)
    print("\n==============================\n")

    consumer_number = re.search(
        r'(Consumer\s*(No|Number)|Customer\s*ID)\s*[:\-]?\s*(\d+)',
        text,
        re.IGNORECASE
    )

    bill_amount = re.search(
        r'(Bill\s*Amount|Current\s*Bill|Total\s*Amount)\s*[:\-]?\s*₹?\s*([\d,]+\.\d+|[\d,]+)',
        text,
        re.IGNORECASE
    )

    units_consumed = re.search(
        r'(Units\s*Consumed|Consumption|Units)\s*[:\-]?\s*(\d+)',
        text,
        re.IGNORECASE
    )

    sanctioned_load = re.search(
        r'(Sanctioned\s*Load|Load)\s*[:\-]?\s*([\d.]+)',
        text,
        re.IGNORECASE
    )

    data = {
        "consumer_number": consumer_number.group(3) if consumer_number else "0",
        "bill_amount": bill_amount.group(2) if bill_amount else "0",
        "units_consumed": units_consumed.group(2) if units_consumed else "0",
        "sanctioned_load": sanctioned_load.group(2) if sanctioned_load else "0"
    }

    return data

# =========================
# EXCEL GENERATION
# =========================

def create_excel(data, excel_path):

    df = pd.DataFrame([data])

    df.to_excel(excel_path, index=False)

# =========================
# HOME ROUTE
# =========================

@app.route("/")
def home():

    return """
    <h1>⚡ Electricity Bill Automation</h1>

    <form action="/upload" method="POST" enctype="multipart/form-data">

        <input type="file" name="file" required>

        <button type="submit">Upload PDF</button>

    </form>
    """

# =========================
# UPLOAD ROUTE
# =========================

@app.route("/upload", methods=["POST"])
def upload_file():

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"})

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No selected file"})

    filename = secure_filename(file.filename)

    pdf_path = os.path.join(UPLOAD_FOLDER, filename)

    file.save(pdf_path)

    # =========================
    # EXTRACT TEXT
    # =========================

    text = extract_text_from_pdf(pdf_path)

    # =========================
    # EXTRACT DATA
    # =========================

    extracted_data = extract_bill_data(text)

    # =========================
    # CREATE EXCEL
    # =========================

    excel_filename = filename.replace(".pdf", ".xlsx")

    excel_path = os.path.join(EXCEL_FOLDER, excel_filename)

    create_excel(extracted_data, excel_path)

    return jsonify({
        "message": "File Uploaded Successfully",
        "extracted_data": extracted_data,
        "excel_download": f"/download/{excel_filename}"
    })

# =========================
# DOWNLOAD EXCEL
# =========================

@app.route("/download/<filename>")
def download_excel(filename):

    file_path = os.path.join(EXCEL_FOLDER, filename)

    return send_file(
        file_path,
        as_attachment=True
    )

# =========================
# MAIN
# =========================

if __name__ == "__main__":
    app.run(debug=True)
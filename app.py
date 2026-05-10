import os
import re
import pdfplumber
import pandas as pd

from flask import Flask, request, send_file
from werkzeug.utils import secure_filename

# ======================================
# CONFIG
# ======================================

UPLOAD_FOLDER = "uploads"
EXCEL_FOLDER = "excel"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXCEL_FOLDER, exist_ok=True)

app = Flask(__name__)

# ======================================
# PDF TEXT EXTRACTION
# ======================================

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

# ======================================
# DATA EXTRACTION
# ======================================

def extract_bill_data(text):

    print("\n========== PDF TEXT ==========\n")
    print(text)
    print("\n==============================\n")

    data = {}

    patterns = {

        "consumer_number": [
            r"Consumer\s*No\.?\s*[:\-]?\s*(\d+)",
            r"Customer\s*ID\s*[:\-]?\s*(\d+)",
        ],

        "bill_amount": [
            r"Bill\s*Amount\s*[:\-]?\s*₹?\s*([\d,.]+)",
            r"Current\s*Bill\s*[:\-]?\s*₹?\s*([\d,.]+)",
            r"Total\s*Amount\s*[:\-]?\s*₹?\s*([\d,.]+)",
        ],

        "units_consumed": [
            r"Units\s*Consumed\s*[:\-]?\s*(\d+)",
            r"Consumption\s*[:\-]?\s*(\d+)",
        ],

        "sanctioned_load": [
            r"Sanctioned\s*Load\s*[:\-]?\s*([\d.]+)",
            r"Load\s*[:\-]?\s*([\d.]+)",
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

# ======================================
# CREATE EXCEL
# ======================================

def create_excel(data, excel_path):

    df = pd.DataFrame([data])

    df.to_excel(excel_path, index=False)

# ======================================
# HOME PAGE
# ======================================

@app.route("/")
def home():

    return """
    <!DOCTYPE html>
    <html>

    <head>

        <title>Electricity Bill Automation</title>

        <style>

            body{
                font-family:Arial;
                background:#f4f7fb;
                display:flex;
                justify-content:center;
                align-items:center;
                height:100vh;
                margin:0;
            }

            .container{
                background:white;
                padding:40px;
                border-radius:15px;
                box-shadow:0 0 20px rgba(0,0,0,0.1);
                width:400px;
                text-align:center;
            }

            h1{
                color:#1f4e79;
                margin-bottom:25px;
            }

            input[type=file]{
                margin:20px 0;
            }

            button{
                background:#1f4e79;
                color:white;
                border:none;
                padding:12px 20px;
                border-radius:8px;
                cursor:pointer;
                width:100%;
                font-size:16px;
            }

            button:hover{
                background:#163a5c;
            }

        </style>

    </head>

    <body>

        <div class="container">

            <h1>⚡ Electricity Bill Automation</h1>

            <form action="/upload" method="POST" enctype="multipart/form-data">

                <input type="file" name="file" required>

                <br>

                <button type="submit">
                    Upload PDF
                </button>

            </form>

        </div>

    </body>

    </html>
    """

# ======================================
# UPLOAD ROUTE
# ======================================

@app.route("/upload", methods=["POST"])
def upload_file():

    if "file" not in request.files:
        return "No file uploaded"

    file = request.files["file"]

    if file.filename == "":
        return "No selected file"

    filename = secure_filename(file.filename)

    pdf_path = os.path.join(UPLOAD_FOLDER, filename)

    file.save(pdf_path)

    # ======================================
    # EXTRACT PDF TEXT
    # ======================================

    text = extract_text_from_pdf(pdf_path)

    # ======================================
    # EXTRACT BILL DATA
    # ======================================

    extracted_data = extract_bill_data(text)

    # ======================================
    # CREATE EXCEL
    # ======================================

    excel_filename = filename.replace(".pdf", ".xlsx")

    excel_path = os.path.join(EXCEL_FOLDER, excel_filename)

    create_excel(extracted_data, excel_path)

    # ======================================
    # RESULT PAGE
    # ======================================

    return f"""
    <!DOCTYPE html>
    <html>

    <head>

        <title>Result</title>

        <style>

            body{{
                font-family:Arial;
                background:#f4f7fb;
                padding:40px;
            }}

            .card{{
                background:white;
                padding:30px;
                border-radius:15px;
                box-shadow:0 0 15px rgba(0,0,0,0.1);
                max-width:500px;
                margin:auto;
            }}

            h2{{
                color:#1f4e79;
            }}

            p{{
                font-size:18px;
            }}

            a{{
                display:inline-block;
                margin-top:20px;
                background:#1f4e79;
                color:white;
                padding:12px 20px;
                border-radius:8px;
                text-decoration:none;
            }}

        </style>

    </head>

    <body>

    <div class="card">

    <h2>⚡ Extracted Bill Data</h2>

    <p><b>Consumer Number:</b> {extracted_data['consumer_number']}</p>

    <p><b>Bill Amount:</b> ₹ {extracted_data['bill_amount']}</p>

    <p><b>Units Consumed:</b> {extracted_data['units_consumed']}</p>

    <p><b>Sanctioned Load:</b> {extracted_data['sanctioned_load']}</p>

    <a href="/download/{excel_filename}">
    Download Excel
    </a>

    </div>

    </body>

    </html>
    """

# ======================================
# DOWNLOAD EXCEL
# ======================================

@app.route("/download/<filename>")
def download_excel(filename):

    file_path = os.path.join(EXCEL_FOLDER, filename)

    return send_file(
        file_path,
        as_attachment=True
    )

# ======================================
# MAIN
# ======================================

if __name__ == "__main__":
    app.run(debug=True)
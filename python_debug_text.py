# python_debug_text.py

import pdfplumber
import json
from process_credit_report import CreditReportProcessor  # ✅ import the parser class

# 👇 Use the correct PDF path
PDF_PATH = "temp/3-Bureau Credit Report & Scores  MyFreeScoreNow (2).pdf"

# Step 1: Load text from the PDF
with pdfplumber.open(PDF_PATH) as pdf:
    full_text = ""
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"

# Step 2: Use the parser
parser = CreditReportProcessor()
parser.parse_credit_report(full_text)

# Step 3: Print only the parsed result, not the class
print("📦 Parsed Credit Report:")
print(json.dumps(parser.credit_report, indent=2))

#!/usr/bin/env python
import logging
import os
import re
import json
import sys
from datetime import datetime
import PyPDF2
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Set pdfplumber's logger to ERROR level to reduce noise
logging.getLogger("pdfplumber").setLevel(logging.ERROR)

class CreditReportProcessor:
    def __init__(self, pdf_path=None, output_dir="output"):
        """Initialize the CreditReportProcessor with default parameters"""
        self.pdf_path = pdf_path
        self.output_dir = output_dir
        self.json_output = os.path.join(output_dir, "credit_report.json")
        self.excel_output = os.path.join(output_dir, "credit_report.xlsx")
        self.initialize_credit_report()
        
    def initialize_credit_report(self):
        """Initialize the credit report data structure"""
        self.credit_report = {
            "personal_info": {"transunion": {}, "experian": {}, "equifax": {}},
            "summary": {"transunion": {}, "experian": {}, "equifax": {}},
            "accounts": [],
            "collections": [],
            "inquiries": [],
            "scores": {"transunion": "N/A", "experian": "N/A", "equifax": "N/A"}
        }

    def process(self, pdf_path=None, output_dir=None):
        if pdf_path:
            self.pdf_path = pdf_path
        if output_dir:
            self.output_dir = output_dir
            self.excel_output = os.path.join(output_dir, "credit_report.xlsx")

        if not self.pdf_path:
            raise ValueError("PDF path is required.")

        os.makedirs(self.output_dir, exist_ok=True)
        self.initialize_credit_report()

        print(f"📄 Processing credit report: {self.pdf_path}")

        lines = self.extract_lines_from_pdf()
        if not lines:
            print("❌ Failed to extract text from PDF")
            return None
        print("\n--- Raw Extracted Lines for Inspection ---")
        for idx, line in enumerate(lines):
        # Pay close attention to lines that should contain phrases like "Account not disputed"
     # Look for extra spaces, non-standard characters, or unexpected line breaks.
            print(f"Line {idx}: '{line}'")
        print("------------------------------------------\n")

        print("🔍 Parsing credit report data...")
        self.parse_credit_report(lines)

        print(f"📊 Creating Excel file: {self.excel_output}")
        self.create_excel()

        print("✅ Processing complete!")
        return self.credit_report  # ✅ <--- this is what your frontend needs



    def extract_lines_from_pdf(self):
        """Extract text lines from PDF using pdfplumber"""
        if not os.path.exists(self.pdf_path):
            print(f"Error: PDF file {self.pdf_path} not found")
            return None
        
        print("Extracting text from PDF...")
        all_lines = []
        
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_lines.extend(text.split('\n'))
            print(f"Extracted {len(all_lines)} lines from PDF")
            return all_lines
        except Exception as e:
            print(f"PDF extraction error: {e}")
            return None

    def parse_credit_report(self, lines):
        """Parse the entire credit report lines into structured data"""
        self.credit_report["scores"] = self.extract_scores_lines(lines)
        print("✅ Final scores saved in report:", self.credit_report["scores"])
        self.extract_personal_info(lines)
        self.extract_summary(lines)
        self.extract_inquiries(lines)
        self.extract_accounts_data(lines)
        self.extract_collections(lines)
        
    def extract_personal_info(self, lines):
        """Extract personal information from lines"""
        bureaus = ["transunion", "experian", "equifax"]
        
        # Extract name for each bureau
        for i, line in enumerate(lines):
            line_lower = line.strip().lower()
            if line_lower in bureaus and i + 2 < len(lines):
                name = lines[i + 2].strip()
                if name and name.isupper():
                    self.credit_report["personal_info"][line_lower]["name"] = name
                    print(f"✅ Found name for {line_lower}: {name}")

        # Extract DOB
        for line in lines:
            dob_match = re.search(r'Date of Birth\s+(\d{4})', line)
            if dob_match:
                for bureau in bureaus:
                    self.credit_report["personal_info"][bureau]["year_of_birth"] = dob_match.group(1)
                print(f"✅ Found birth year: {dob_match.group(1)}")
                break

        # Extract address
        address_lines = []
        in_address_section = False
        for line in lines:
            if "Current Address" in line:
                in_address_section = True
                continue
            if in_address_section and ("Previous Address" in line or "Employer" in line):
                break
            if in_address_section and line.strip():
                address_lines.append(line.strip())
        
        if address_lines:
            for i, bureau in enumerate(bureaus):
                if i < len(address_lines):
                    self.credit_report["personal_info"][bureau]["current_address"] = address_lines[i]
                    print(f"✅ Found address for {bureau}: {address_lines[i]}")

    def extract_summary(self, lines):
        """Extract summary information from lines"""
        summary_patterns = {
            r'Total Accounts\s+(\d+)\s+(\d+)\s+(\d+)': "total_accounts",
            r'Open Accounts:\s+(\d+)\s+(\d+)\s+(\d+)': "open_accounts",
            r'Closed Accounts:\s+(\d+)\s+(\d+)\s+(\d+)': "closed_accounts",
            r'Delinquent:\s+(\d+)\s+(\d+)\s+(\d+)': "delinquent",
            r'Derogatory:\s+(\d+)\s+(\d+)\s+(\d+)': "derogatory",
            r'Balances:\s+\$([0-9,]+)\s+\$([0-9,]+)\s+\$([0-9,]+)': "balances",
            r'(?:Monthly )?Payments?:\s+\$([0-9,]+)\s+\$([0-9,]+)\s+\$([0-9,]+)': "monthly_payments",
            r'Credit Utilization:\s+(\d+)%\s+(\d+)%\s+(\d+)%': "credit_utilization",
            r'Public Records:\s+(\d+)\s+(\d+)\s+(\d+)': "public_records",
            r'Inquiries\s+\(2\s+years?\):\s+(\d+)\s+(\d+)\s+(\d+)': "inquiries_2y"
        }
        
        bureaus = ["transunion", "experian", "equifax"]
        
        for line in lines:
            for pattern, key in summary_patterns.items():
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    for i, bureau in enumerate(bureaus):
                        self.credit_report["summary"][bureau][key] = match.group(i+1)
                    break

    def extract_inquiries(self, lines):
        """Extract inquiries data from lines"""
        print("Extracting inquiries...")
        in_inquiries_section = False
        inquiry_lines = []
        
        for line in lines:
            if "Inquiries" in line and ("Creditor Name" in line or "Date of Inquiry" in line):
                in_inquiries_section = True
                continue
            if in_inquiries_section and ("Accounts" in line or "Public Records" in line or "Summary" in line):
                break
            if in_inquiries_section and line.strip():
                inquiry_lines.append(line.strip())
        
        for line in inquiry_lines:
            parts = self.split_into_three_parts(line)
            if len(parts) == 3 and parts[0] and parts[1] and parts[2]:
                creditor, date, bureau = parts
                if creditor.lower() not in ["creditor name", "inquiries"]:
                    self.credit_report["inquiries"].append({
                        "creditor": creditor,
                        "date": self.clean_value(date),
                        "bureau": bureau.lower()
                    })
        
        print(f"Extracted {len(self.credit_report['inquiries'])} inquiries")

    def split_into_three_parts(self, line):
        """Split a line into three parts for creditor, date, and bureau"""
        parts = re.split(r'\t+|\s{2,}', line.strip())
        while len(parts) < 3:
            parts.append("")
        return parts[:3]

    def extract_accounts_data(self, lines):
        """Extract account data from lines"""
        print("Extracting account data...")
        account_sections = self.identify_account_sections(lines)
        
        for account_lines in account_sections:
            try:
                account_data = self.parse_account(account_lines)
                if account_data:
                    self.credit_report["accounts"].append(account_data)
            except Exception as e:
                print(f"Error parsing account: {e}")
        
        print(f"Extracted {len(self.credit_report['accounts'])} accounts")

    def identify_account_sections(self, lines):
        """Identify sections that contain account information"""
        account_sections = []
        current_section = []
        in_account_section = False
        
        for line in lines:
            # Check if line is a potential creditor name
            if (re.match(r'^[A-Z][A-Z0-9\s&.,\'"-/]+$', line.strip()) and 
                len(line.strip()) < 40 and
                not any(char.isdigit() for char in line.strip()) and
                not line.strip().startswith("Page")):
                
                if current_section:
                    account_sections.append(current_section)
                current_section = [line.strip()]
                in_account_section = True
            elif in_account_section:
                current_section.append(line.strip())
        
        if current_section:
            account_sections.append(current_section)
        
        return account_sections
    def parse_account(self, account_lines):
        if not account_lines:
            return None

        creditor = account_lines[0].strip()
        account_data = {
        "creditor": creditor,
        "transunion": {},
        "experian": {},
        "equifax": {}
        }

        def extract_bureau_values(line, prefix=""):
            """Improved value extraction that handles multi-word values"""
            # Remove the prefix if present
            if prefix and line.startswith(prefix):
                line = line[len(prefix):].strip()
        
        # Split into bureau columns while preserving multi-word values
            parts = []
            current_part = []
            in_value = False
        
            for char in line:
                if char == '\t' or (char == ' ' and len(current_part) > 0 and not in_value):
                    if current_part:
                        parts.append(''.join(current_part).strip())
                        current_part = []
                else:
                    current_part.append(char)
                    if char == ' ':
                        in_value = True
        
            if current_part:
                parts.append(''.join(current_part).strip())
        
        # Ensure we have 3 parts
            while len(parts) < 3:
                parts.append("N/A")
            
            return parts[:3]

        i = 1
        while i < len(account_lines):
            line = account_lines[i].strip()

        if "Account #" in line:
            parts = extract_bureau_values(line, "Account #")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["account_number"] = parts[idx]

        elif "High Balance:" in line:
            parts = extract_bureau_values(line, "High Balance:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["high_balance"] = parts[idx]

        elif "Balance Owed:" in line:
            parts = extract_bureau_values(line, "Balance Owed:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["balance_owed"] = parts[idx]

        elif "Dispute Status:" in line:
            parts = extract_bureau_values(line, "Dispute Status:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["dispute_status"] = parts[idx]

        elif "Payment Status:" in line:
            parts = extract_bureau_values(line, "Payment Status:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["payment_status"] = parts[idx]

        elif "Account Status:" in line:
            parts = extract_bureau_values(line, "Account Status:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["account_status"] = parts[idx]

        elif "Creditor Remarks:" in line:
            parts = extract_bureau_values(line, "Creditor Remarks:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["remarks"] = parts[idx]

        elif "Date Opened:" in line:
            parts = extract_bureau_values(line, "Date Opened:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["date_opened"] = parts[idx]

        elif "Date Closed:" in line:
            parts = extract_bureau_values(line, "Date Closed:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["date_closed"] = parts[idx]

        elif "Last Reported:" in line:
            parts = extract_bureau_values(line, "Last Reported:")
            for idx, bureau in enumerate(["transunion", "experian", "equifax"]):
                account_data[bureau]["last_reported"] = parts[idx]

        i += 1

        return account_data



    def extract_collections(self, lines):
        """Extract collection accounts from lines"""
        print("Extracting collection accounts...")
        collection_sections = []
        current_section = []
        in_collection_section = False
        
        for line in lines:
            if "Collection" in line.strip():
                if current_section:
                    collection_sections.append(current_section)
                current_section = [line.strip()]
                in_collection_section = True
            elif in_collection_section:
                if line.strip() and not line.strip().startswith("Page"):
                    current_section.append(line.strip())
                else:
                    in_collection_section = False
                    if current_section:
                        collection_sections.append(current_section)
                        current_section = []
        
        if current_section:
            collection_sections.append(current_section)
        
        for section in collection_sections:
            try:
                collection_data = self.parse_collection(section)
                if collection_data:
                    self.credit_report["collections"].append(collection_data)
            except Exception as e:
                print(f"Error parsing collection: {e}")
        
        print(f"Extracted {len(self.credit_report['collections'])} collection accounts")

    def parse_collection(self, collection_lines):
        """Parse a collection account section into structured data"""
        if not collection_lines:
            return None
            
        agency = collection_lines[0].replace("Collection", "").strip()
        collection_data = {
            "agency": agency,
            "transunion": {},
            "experian": {},
            "equifax": {}
        }
        
        # Parse collection details from the lines
        for line in collection_lines[1:]:
            if "Account #" in line:
                parts = line.split("Account #")[1].strip().split()
                if len(parts) >= 3:
                    for i, bureau in enumerate(["transunion", "experian", "equifax"]):
                        collection_data[bureau]["account_number"] = parts[i] if i < len(parts) else ""
            
            if "Balance Owed:" in line:
                parts = line.split("Balance Owed:")[1].strip().split()
                if len(parts) >= 3:
                    for i, bureau in enumerate(["transunion", "experian", "equifax"]):
                        collection_data[bureau]["balance_owed"] = parts[i] if i < len(parts) else ""
        
        return collection_data

    def clean_value(self, value):
        """Clean and standardize values"""
        if not value or value == '--' or value == 'N/A':
            return ''
        return str(value).replace('$', '').replace(',', '').strip()

    def save_json(self):
        print(f"✅ Writing to: {self.json_output}")
        with open(self.json_output, 'w', encoding='utf-8') as f:
            json.dump(self.credit_report, f, indent=2)
        print(f"✅ JSON data saved to {self.json_output}")
    
    def extract_scores_lines(lines):
        """Extract VantageScore 3.0 credit scores from the PDF lines"""
        scores = {"transunion": "N/A", "experian": "N/A", "equifax": "N/A"}
    
    # Look for the scores header line
        for i, line in enumerate(lines):
            if "Your 3B Report & Vantage Scores® 3.0" in line:
            # The scores are typically in a table format after this header
            # Look for the bureau names line
                for j in range(i, min(i+10, len(lines))):  # Search next 10 lines
                    if "Transunion" in lines[j] and "Experian" in lines[j] and "Equifax" in lines[j]:
                    # Scores are usually on the next line
                        if j+1 < len(lines):
                            score_line = lines[j+1]
                        # Extract 3-digit numbers
                            score_matches = re.findall(r'\b\d{3}\b', score_line)
                            if len(score_matches) >= 3:
                                scores["transunion"] = score_matches[0]
                                scores["experian"] = score_matches[1]
                                scores["equifax"] = score_matches[2]
                                break
                break
            
        print(f"Extracted scores: {scores}")
        return scores



def main():
    # Default values
    pdf_path = "/Users/Alwinsaji/Downloads/newproject4/Credit report 0- Mylashia Monae Montgomery.pdf"
    output_dir = "output"

    # Use command-line arguments if provided
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    if len(sys.argv) > 2:
        output_dir = sys.argv[2]

    processor = CreditReportProcessor()
    processor.process(pdf_path, output_dir)

if __name__ == "__main__": 
    main()
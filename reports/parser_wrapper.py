from process_credit_report import CreditReportProcessor
import json
def process_credit_report(pdf_path):
    processor = CreditReportProcessor(pdf_path)
    processor.process()
    print("✅ FINAL REPORT:", json.dumps(processor.credit_report, indent=2))  # Add this
    return processor.credit_report

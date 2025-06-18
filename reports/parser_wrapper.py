from process_credit_report import CreditReportProcessor
import json
def process_credit_report(pdf_path):
    print("⚠️ parser_wrapper: Wrapper function called")
    processor = CreditReportProcessor(pdf_path)
    print(f"🧪 Loaded processor from: {processor.__class__}")
    processor.process()
    return processor.credit_report

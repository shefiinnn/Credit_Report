from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
from process_credit_report import CreditReportProcessor

app = FastAPI()

# Allow frontend access (CORS)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Change * to specific origin in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve HTML frontend
@app.get("/", response_class=HTMLResponse)
def serve_html():
    with open("reports/templates/reports/MyFreeScoreNow.html", "r", encoding="utf-8") as file:
        return file.read()

# Handle PDF POST and return parsed data
@app.post("/get_report")
async def get_report(pdf: UploadFile = File(...)):
    os.makedirs("temp", exist_ok=True)
    pdf_path = os.path.join("temp", pdf.filename)
    
    with open(pdf_path, "wb") as f:
        content = await pdf.read()
        f.write(content)
    
    try:
        processor = CreditReportProcessor(pdf_path)
        processor.process()
        return JSONResponse(content={"status": "success", "data": processor.credit_report})
    except Exception as e:
        return JSONResponse(content={"status": "error", "message": str(e)}, status_code=500)

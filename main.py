import os
import subprocess
import shutil
import uuid
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pdf2docx import Converter

app = FastAPI(title="Resumora Engine")

# CORS config
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "/tmp/conversions"
os.makedirs(TEMP_DIR, exist_ok=True)

@app.get("/")
def home():
    return {"status": "online", "message": "Resumora Engine Active"}

# 1. WORD TO PDF (LibreOffice)
@app.post("/convert/word-to-pdf")
async def word_to_pdf(file: UploadFile = File(...)):
    if not (file.filename.endswith(".docx") or file.filename.endswith(".doc")):
        raise HTTPException(status_code=400, detail="Only DOC or DOCX allowed")

    session_id = str(uuid.uuid4())
    input_file = os.path.join(TEMP_DIR, f"{session_id}_{file.filename}")

    with open(input_file, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            input_file,
            "--outdir", TEMP_DIR
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        base_name = os.path.splitext(input_file)[0]
        output_pdf = f"{base_name}.pdf"

        if not os.path.exists(output_pdf):
            raise HTTPException(status_code=500, detail="Conversion failed")

        return FileResponse(
            output_pdf,
            media_type="application/pdf",
            filename=f"{os.path.splitext(file.filename)[0]}.pdf"
        )
    finally:
        if os.path.exists(input_file):
            os.remove(input_file)

# 2. PDF TO WORD (Fault-Tolerant Hybrid Engine)
@app.post("/convert/pdf-to-word")
async def pdf_to_word(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF allowed")

    session_id = str(uuid.uuid4())
    input_file = os.path.join(TEMP_DIR, f"{session_id}_{file.filename}")
    output_docx = os.path.join(TEMP_DIR, f"{session_id}_converted.docx")

    with open(input_file, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    conversion_successful = False

    # Attempt 1: LibreOffice Writer direct DOCX conversion (Valid XML guaranteed)
    try:
        cmd = [
            "libreoffice",
            "--headless",
            "--infilter=writer_pdf_import",
            "--convert-to", "docx:MS Word 2007 XML",
            input_file,
            "--outdir", TEMP_DIR
        ]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=45)
        base_name = os.path.splitext(input_file)[0]
        lo_output = f"{base_name}.docx"

        if os.path.exists(lo_output) and os.path.getsize(lo_output) > 1000:
            output_docx = lo_output
            conversion_successful = True
    except Exception as e:
        print(f"LibreOffice conversion failed: {e}")

    # Attempt 2: pdf2docx fallback with safe parse settings
    if not conversion_successful:
        try:
            cv = Converter(input_file)
            # connected_borders=False table XML corruptions prevent karta hai
            cv.convert(output_docx, start=0, end=None, connected_borders=False)
            cv.close()
            if os.path.exists(output_docx) and os.path.getsize(output_docx) > 500:
                conversion_successful = True
        except Exception as e:
            print(f"pdf2docx fallback error: {e}")

    if not conversion_successful or not os.path.exists(output_docx):
        if os.path.exists(input_file):
            os.remove(input_file)
        raise HTTPException(status_code=500, detail="Conversion failed. Document contains unparseable vectors.")

    if os.path.exists(input_file):
        os.remove(input_file)

    return FileResponse(
        output_docx,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{os.path.splitext(file.filename)[0]}.docx"
    )

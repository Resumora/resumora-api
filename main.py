import os
import subprocess
import shutil
import uuid
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pdf2docx import Converter

app = FastAPI(title="Resumora Engine")

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

# 1. WORD TO PDF (LibreOffice Headless)
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

# 2. PDF TO WORD (Pure Native Text & Flow Layout)
@app.post("/convert/pdf-to-word")
async def pdf_to_word(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF allowed")

    session_id = str(uuid.uuid4())
    input_file = os.path.join(TEMP_DIR, f"{session_id}_{file.filename}")
    output_docx = os.path.join(TEMP_DIR, f"{session_id}_converted.docx")

    with open(input_file, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        # pdf2docx with clean typography settings (bina font distortion ke)
        cv = Converter(input_file)
        cv.convert(
            output_docx,
            start=0,
            end=None,
            multi_processing=False,
            connected_borders=False
        )
        cv.close()

        if not os.path.exists(output_docx) or os.path.getsize(output_docx) < 500:
            raise HTTPException(status_code=500, detail="Conversion output empty")

        return FileResponse(
            output_docx,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=f"{os.path.splitext(file.filename)[0]}.docx"
        )
    except Exception as e:
        print(f"pdf2docx primary failed: {e}")
        # Secondary fallback
        try:
            cmd = [
                "libreoffice",
                "--headless",
                "--convert-to", "docx",
                input_file,
                "--outdir", TEMP_DIR
            ]
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=40)
            base_name = os.path.splitext(input_file)[0]
            lo_output = f"{base_name}.docx"
            if os.path.exists(lo_output) and os.path.getsize(lo_output) > 500:
                return FileResponse(
                    lo_output,
                    media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    filename=f"{os.path.splitext(file.filename)[0]}.docx"
                )
        except Exception as lo_err:
            print(f"Fallback error: {lo_err}")

        raise HTTPException(status_code=500, detail="Could not parse document structure.")
    finally:
        if os.path.exists(input_file):
            os.remove(input_file)

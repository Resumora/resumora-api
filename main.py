import os
import subprocess
import shutil
import uuid
import asyncio
import docx
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pdf2docx import Converter

app = FastAPI(title="Resumora Engine - Production")

# CORS setup
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "/tmp/conversions"
os.makedirs(TEMP_DIR, exist_ok=True)

# ----------------- CRASH PROTECTION CONFIG -----------------
# 1. File Size Protection: Max 15 MB allowed (OOM crash se bachane ke liye)
MAX_FILE_SIZE = 15 * 1024 * 1024  

# 2. Concurrency Lock: Ek time par maximum 2 heavy tasks chalenge, baaki queue me wait karenge
# (Isse ek sath 20 users aane par bhi server ka CPU 100% freeze nahi hoga)
CONVERSION_SEMAPHORE = asyncio.Semaphore(2)

# 3. Timeout: Agar koi conversion 45 second se zyada le toh terminate kar do
PROCESS_TIMEOUT = 45
# -----------------------------------------------------------


def cleanup_files(*file_paths):
    """Temporary files ko background me safely delete karne ke liye"""
    for path in file_paths:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass


def clean_and_brand_docx(docx_path: str):
    """Python-docx metadata ko override karke Resumora branding set karta hai"""
    try:
        doc = docx.Document(docx_path)
        core_props = doc.core_properties
        core_props.author = "Resumora Universe"
        core_props.last_modified_by = "Resumora Universe"
        core_props.comments = "Processed privately via Resumora Engine"
        doc.save(docx_path)
    except Exception as err:
        print(f"Metadata sanitize error: {err}")


@app.get("/")
def home():
    return {"status": "online", "message": "Resumora Engine Production Active"}


# ==========================================
# 1. WORD TO PDF (Crash Safe)
# ==========================================
@app.post("/convert/word-to-pdf")
async def word_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    # Extension check
    if not (file.filename.endswith(".docx") or file.filename.endswith(".doc")):
        raise HTTPException(status_code=400, detail="Only DOC or DOCX files allowed.")

    # Size check before saving to disk
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File too large! Maximum limit is 15 MB.")
    file.file.seek(0)

    session_id = str(uuid.uuid4())
    input_file = os.path.join(TEMP_DIR, f"{session_id}_{file.filename}")
    base_name = os.path.splitext(input_file)[0]
    output_pdf = f"{base_name}.pdf"

    # Save uploaded file
    with open(input_file, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Concurrency Queue ke andar execute karein
    async with CONVERSION_SEMAPHORE:
        try:
            cmd = [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                input_file,
                "--outdir", TEMP_DIR
            ]
            # Timeout ke sath run karein taaki LibreOffice hang na ho
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=PROCESS_TIMEOUT)

            if not os.path.exists(output_pdf) or os.path.getsize(output_pdf) < 500:
                raise HTTPException(status_code=500, detail="Conversion generated empty file.")

            # Background me clean karein download deliver hone ke baad
            background_tasks.add_task(cleanup_files, input_file, output_pdf)

            return FileResponse(
                output_pdf,
                media_type="application/pdf",
                filename=f"{os.path.splitext(file.filename)[0]}.pdf"
            )

        except subprocess.TimeoutExpired:
            cleanup_files(input_file, output_pdf)
            raise HTTPException(status_code=504, detail="Conversion timed out. File might be too complex.")
        except Exception as e:
            cleanup_files(input_file, output_pdf)
            raise HTTPException(status_code=500, detail=f"Word to PDF conversion failed: {str(e)}")


# ==========================================
# 2. PDF TO WORD (Layout Preserved & Crash Safe)
# ==========================================
@app.post("/convert/pdf-to-word")
async def pdf_to_word(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    # Extension check
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files allowed.")

    # Size check before saving to disk
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File too large! Maximum limit is 15 MB.")
    file.file.seek(0)

    session_id = str(uuid.uuid4())
    input_file = os.path.join(TEMP_DIR, f"{session_id}_{file.filename}")
    output_docx = os.path.join(TEMP_DIR, f"{session_id}_converted.docx")

    # Save uploaded file
    with open(input_file, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Concurrency Queue ke andar execute karein
    async with CONVERSION_SEMAPHORE:
        conversion_success = False

        # Primary Engine: pdf2docx
        try:
            loop = asyncio.get_event_loop()
            
            def run_pdf2docx():
                cv = Converter(input_file)
                cv.convert(
                    output_docx,
                    start=0,
                    end=None,
                    multi_processing=False,
                    keep_layout=True
                )
                cv.close()

            # Async loop me run taaki server event loop block na ho
            await asyncio.wait_for(loop.run_in_executor(None, run_pdf2docx), timeout=PROCESS_TIMEOUT)

            if os.path.exists(output_docx) and os.path.getsize(output_docx) > 500:
                conversion_success = True

        except Exception as e:
            print(f"pdf2docx conversion error: {e}")

        # Secondary Fallback: LibreOffice (agar pdf2docx fail ho jaye)
        if not conversion_success:
            try:
                cmd = [
                    "libreoffice",
                    "--headless",
                    "--convert-to", "docx",
                    input_file,
                    "--outdir", TEMP_DIR
                ]
                subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=PROCESS_TIMEOUT)
                base_name = os.path.splitext(input_file)[0]
                lo_output = f"{base_name}.docx"

                if os.path.exists(lo_output) and os.path.getsize(lo_output) > 500:
                    output_docx = lo_output
                    conversion_success = True
            except Exception as lo_err:
                print(f"LibreOffice fallback error: {lo_err}")

        # Final Verification & Response
        if conversion_success and os.path.exists(output_docx):
            # 'python-docx' branding remove karke Resumora Universe set karein
            clean_and_brand_docx(output_docx)

            background_tasks.add_task(cleanup_files, input_file, output_docx)

            return FileResponse(
                output_docx,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename=f"{os.path.splitext(file.filename)[0]}.docx"
            )
        else:
            cleanup_files(input_file, output_docx)
            raise HTTPException(status_code=500, detail="Could not convert document. Please check PDF format.")

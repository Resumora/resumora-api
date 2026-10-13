import os
import re
import shutil
import uuid
import asyncio
import subprocess
import docx
from docx.oxml import parse_xml
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pdf2docx import Converter

app = FastAPI(title="Resumora Engine - Bulletproof Production")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "/tmp/conversions"
os.makedirs(TEMP_DIR, exist_ok=True)

# ----------------- PRODUCTION CONSTRAINTS -----------------
MAX_FILE_SIZE = 15 * 1024 * 1024  # Strict 15 MB
CONVERSION_SEMAPHORE = asyncio.Semaphore(2)  # Max 2 heavy conversions concurrently
PROCESS_TIMEOUT = 45  # 45 Seconds strict timeout
CHUNK_SIZE = 1024 * 1024  # 1 MB stream chunks
# -----------------------------------------------------------


def sanitize_filename(filename: str) -> str:
    """Special characters aur spaces ko clean karke safe banata hai"""
    name, ext = os.path.splitext(filename)
    clean_name = re.sub(r'[^a-zA-Z0-9_\-]', '_', name)
    return f"{clean_name[:40]}{ext}"


def cleanup_paths(*paths):
    """File aur temporary user-profile folders ko safely clean karta hai"""
    for p in paths:
        if p and os.path.exists(p):
            try:
                if os.path.isdir(p):
                    shutil.rmtree(p, ignore_errors=True)
                else:
                    os.remove(p)
            except Exception:
                pass


def patch_docx_for_windows_explorer(docx_path: str):
    """
    1. Python-docx metadata ko Resumora se replace karta hai.
    2. 'app.xml' me Application='Microsoft Office Word' inject karta hai,
       jisse Windows File Explorer bina MS Word khole full native Word icon render kare.
    """
    try:
        doc = docx.Document(docx_path)
        
        # Core metadata clean
        core_props = doc.core_properties
        core_props.author = "Resumora Universe"
        core_props.last_modified_by = "Resumora Universe"
        core_props.comments = "Generated privately via Resumora Universe Engine"
        core_props.category = "Resume Document"

        # Native Microsoft Office Word application signature injection
        try:
            app_part = doc.part.package.app_properties
            app_part.application = "Microsoft Office Word"
            app_part.app_version = "16.0000"
        except Exception:
            pass

        doc.save(docx_path)
    except Exception as err:
        print(f"DOCX Patching error: {err}")


async def save_file_safely(file: UploadFile, destination_path: str):
    """Memory-safe chunk streaming: RAM full hone se bachata hai"""
    total_size = 0
    with open(destination_path, "wb") as buffer:
        while chunk := await file.read(CHUNK_SIZE):
            total_size += len(chunk)
            if total_size > MAX_FILE_SIZE:
                buffer.close()
                if os.path.exists(destination_path):
                    os.remove(destination_path)
                raise HTTPException(status_code=413, detail="File too large! Max limit is 15 MB.")
            buffer.write(chunk)


@app.get("/")
def health_check():
    return {"status": "online", "engine": "Resumora Universe", "protected": True}


# ========================================================
# 1. WORD TO PDF (Isolated LibreOffice Profiles)
# ========================================================
@app.post("/convert/word-to-pdf")
async def word_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    orig_name = file.filename or "document.docx"
    if not (orig_name.lower().endswith(".docx") or orig_name.lower().endswith(".doc")):
        raise HTTPException(status_code=400, detail="Only .doc or .docx files are accepted.")

    session_id = str(uuid.uuid4())
    safe_name = sanitize_filename(orig_name)
    input_file = os.path.join(TEMP_DIR, f"{session_id}_{safe_name}")
    libreoffice_profile = os.path.join(TEMP_DIR, f"lo_profile_{session_id}")
    
    await save_file_safely(file, input_file)

    async with CONVERSION_SEMAPHORE:
        try:
            # -env flag ensure karta hai ki multiple concurrent requests me LibreOffice lock crash na ho
            cmd = [
                "libreoffice",
                f"-env:UserInstallation=file://{libreoffice_profile}",
                "--headless",
                "--convert-to", "pdf",
                input_file,
                "--outdir", TEMP_DIR
            ]
            
            subprocess.run(
                cmd,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=PROCESS_TIMEOUT
            )

            base_name = os.path.splitext(input_file)[0]
            output_pdf = f"{base_name}.pdf"

            if not os.path.exists(output_pdf) or os.path.getsize(output_pdf) < 500:
                raise HTTPException(status_code=500, detail="Conversion generated empty file.")

            client_download_name = f"{os.path.splitext(safe_name)[0]}.pdf"
            background_tasks.add_task(cleanup_paths, input_file, output_pdf, libreoffice_profile)

            return FileResponse(
                output_pdf,
                media_type="application/pdf",
                filename=client_download_name,
                headers={"Content-Disposition": f'attachment; filename="{client_download_name}"'}
            )

        except subprocess.TimeoutExpired:
            cleanup_paths(input_file, libreoffice_profile)
            raise HTTPException(status_code=504, detail="File too complex: conversion timed out.")
        except Exception as e:
            cleanup_paths(input_file, libreoffice_profile)
            raise HTTPException(status_code=500, detail=f"Word to PDF conversion error: {str(e)}")


# ========================================================
# 2. PDF TO WORD (High-Precision Formatting + Native Shell Icon)
# ========================================================
@app.post("/convert/pdf-to-word")
async def pdf_to_word(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    orig_name = file.filename or "document.pdf"
    if not orig_name.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only .pdf files are accepted.")

    session_id = str(uuid.uuid4())
    safe_name = sanitize_filename(orig_name)
    input_file = os.path.join(TEMP_DIR, f"{session_id}_{safe_name}")
    output_docx = os.path.join(TEMP_DIR, f"{session_id}_converted.docx")
    libreoffice_profile = os.path.join(TEMP_DIR, f"lo_profile_{session_id}")

    await save_file_safely(file, input_file)

    async with CONVERSION_SEMAPHORE:
        conversion_success = False

        # Primary Engine: pdf2docx with precise layout flow
        try:
            loop = asyncio.get_event_loop()

            def execute_conversion():
                cv = Converter(input_file)
                cv.convert(
                    output_docx,
                    start=0,
                    end=None,
                    multi_processing=False,
                    keep_layout=True
                )
                cv.close()

            await asyncio.wait_for(loop.run_in_executor(None, execute_conversion), timeout=PROCESS_TIMEOUT)

            if os.path.exists(output_docx) and os.path.getsize(output_docx) > 500:
                conversion_success = True

        except Exception as primary_err:
            print(f"pdf2docx engine error: {primary_err}")

        # Secondary Engine Fallback: LibreOffice with isolated profile
        if not conversion_success:
            try:
                cmd = [
                    "libreoffice",
                    f"-env:UserInstallation=file://{libreoffice_profile}",
                    "--headless",
                    "--convert-to", "docx",
                    input_file,
                    "--outdir", TEMP_DIR
                ]
                subprocess.run(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=PROCESS_TIMEOUT
                )
                base_name = os.path.splitext(input_file)[0]
                lo_output = f"{base_name}.docx"

                if os.path.exists(lo_output) and os.path.getsize(lo_output) > 500:
                    output_docx = lo_output
                    conversion_success = True
            except Exception as lo_err:
                print(f"LibreOffice fallback error: {lo_err}")

        if conversion_success and os.path.exists(output_docx):
            # Patch Word document metadata and application signature for full native icon
            patch_docx_for_windows_explorer(output_docx)

            client_download_name = f"{os.path.splitext(safe_name)[0]}.docx"
            background_tasks.add_task(cleanup_paths, input_file, output_docx, libreoffice_profile)

            return FileResponse(
                output_docx,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename=client_download_name,
                headers={"Content-Disposition": f'attachment; filename="{client_download_name}"'}
            )
        else:
            cleanup_paths(input_file, output_docx, libreoffice_profile)
            raise HTTPException(status_code=500, detail="Conversion failed. The PDF might be corrupted or image-only.")

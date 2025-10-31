import os
import tempfile
import traceback
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from app.processor import generate_docx

# Create the app (debug=True helps show clear errors while you set up)
app = FastAPI(title="Health Quote Generator", debug=True)

# Serve the frontend (index.html) from ../frontend
FRONTEND_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "frontend"))
app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")


@app.get("/", response_class=HTMLResponse)
async def home():
    """
    Serve the frontend's index.html at the root.
    """
    index_path = os.path.join(FRONTEND_DIR, "index.html")
    if not os.path.exists(index_path):
        return HTMLResponse(
            "<h3>index.html not found. Make sure frontend/index.html exists.</h3>",
            status_code=500,
        )
    with open(index_path, "r", encoding="utf-8") as f:
        return f.read()


@app.post("/generate")
async def generate(file: UploadFile = File(...)):
    """
    Accept an Excel upload (.xlsx/.xlsm/.xls), generate the DOCX, and return it.
    - Reads the Excel from memory (bytes) to avoid Windows file locks.
    - Uses a NamedTemporaryFile for the DOCX and deletes it after response is sent.
    """
    try:
        filename = (file.filename or "").strip().lower()
        if not filename.endswith((".xlsx", ".xlsm", ".xls")):
            raise HTTPException(
                status_code=400,
                detail="Please upload a valid Excel file (.xlsx/.xlsm/.xls)",
            )

        content = await file.read()
        if not content or len(content) < 100:
            raise HTTPException(
                status_code=400,
                detail="The uploaded file seems empty or not a real Excel. Please re-save as .xlsx and try again.",
            )

        # Create a temporary DOCX path that persists until we delete it
        tmp_doc = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        tmp_doc_path = tmp_doc.name
        tmp_doc.close()

        # Generate the DOCX directly to the temp file; read Excel from memory (bytes)
        out_file = generate_docx(
            excel_input=content,
            output_path=tmp_doc_path,
            logo_folder=os.path.join(os.getcwd(), "logos"),
            filename_hint=filename,
        )

        # Ensure file exists
        if not os.path.exists(out_file):
            raise RuntimeError("Failed to create DOCX file.")

        # Delete the file AFTER it's fully sent
        cleanup = BackgroundTask(lambda p: os.path.exists(p) and os.remove(p), out_file)

        return FileResponse(
            out_file,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="Health_Quote.docx",
            background=cleanup,
        )

    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

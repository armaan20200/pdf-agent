"""
Smart PDF AI Agent - Python FastAPI Backend
All PDF operations use PyMuPDF (fitz) and pdfplumber
Mounted at /api/* prefix
"""

import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
import pdfplumber
from PIL import Image
from fastapi import FastAPI, File, UploadFile, HTTPException, APIRouter
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from agent import PDFAgent
from cleanup import cleanup_old_files

# Create main app
app = FastAPI(title="Smart PDF AI Agent", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = Path("temporary_files")
TEMP_DIR.mkdir(exist_ok=True)

# Clean up old temporary files on startup
cleanup_old_files()

# In-memory state
active_files: dict[str, dict] = {}
result_files: dict[str, dict] = {}
chat_history: list[dict] = []

# Agent instance
agent = PDFAgent(TEMP_DIR, active_files, result_files)

# Router mounted under /api
router = APIRouter(prefix="/api")


def get_pdf_page_count(file_path: str) -> Optional[int]:
    try:
        doc = fitz.open(file_path)
        count = doc.page_count
        doc.close()
        return count
    except Exception:
        return None


@router.get("/healthz")
async def health_check():
    return {"status": "ok"}


@router.get("/files")
async def list_files():
    return list(active_files.values())


@router.post("/files/upload")
async def upload_files(files: list[UploadFile] = File(...)):
    uploaded = []
    for upload in files:
        file_id = str(uuid.uuid4())
        filename = upload.filename or f"file_{file_id}.pdf"
        # Sanitize filename
        safe_name = "".join(c for c in filename if c.isalnum() or c in "._- ").strip()
        if not safe_name:
            safe_name = f"file_{file_id}.pdf"

        file_path = TEMP_DIR / f"{file_id}_{safe_name}"
        content = await upload.read()

        with open(file_path, "wb") as f:
            f.write(content)

        page_count = None
        if filename.lower().endswith(".pdf"):
            page_count = get_pdf_page_count(str(file_path))

        file_info = {
            "id": file_id,
            "name": safe_name,
            "size": len(content),
            "pageCount": page_count,
            "uploadedAt": datetime.utcnow().isoformat() + "Z",
            "path": str(file_path),
        }
        active_files[file_id] = file_info
        uploaded.append(file_info)

    return uploaded


@router.delete("/files/{file_id}")
async def delete_file(file_id: str):
    if file_id not in active_files:
        raise HTTPException(status_code=404, detail="File not found")

    file_info = active_files.pop(file_id)
    try:
        os.remove(file_info["path"])
    except Exception:
        pass

    return {"success": True, "message": "File removed"}


class ChatRequest(BaseModel):
    message: str
    fileIds: Optional[list[str]] = None


@router.post("/agent/chat")
async def agent_chat(request: ChatRequest):
    user_msg_id = str(uuid.uuid4())
    user_message = {
        "id": user_msg_id,
        "role": "user",
        "content": request.message,
        "timestamp": datetime.utcnow().isoformat() + "Z",
    }
    chat_history.append(user_message)

    # Run agent
    try:
        response = await agent.process_message(request.message, request.fileIds or [])
    except Exception as e:
        response = {
            "content": f"An error occurred: {str(e)}",
            "resultFiles": [],
            "toolsUsed": [],
        }

    assistant_msg_id = str(uuid.uuid4())
    assistant_message = {
        "id": assistant_msg_id,
        "role": "assistant",
        "content": response["content"],
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "resultFiles": response.get("resultFiles", []),
        "toolsUsed": response.get("toolsUsed", []),
    }
    chat_history.append(assistant_message)

    return {
        "message": assistant_message,
        "updatedFiles": list(active_files.values()),
    }


@router.get("/agent/history")
async def get_chat_history():
    return chat_history


@router.delete("/agent/history")
async def clear_chat_history():
    chat_history.clear()
    return {"success": True, "message": "History cleared"}


@router.get("/agent/download/{file_id}")
async def download_result_file(file_id: str):
    if file_id not in result_files:
        raise HTTPException(status_code=404, detail="Result file not found")

    file_info = result_files[file_id]
    file_path = file_info["path"]

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File no longer exists")

    return FileResponse(
        path=file_path,
        filename=file_info["name"],
        media_type="application/octet-stream",
    )


# Include router
app.include_router(router)

# ── Serve React frontend (built static files) ────────────────────────
# The 'static' folder is created during packaging; contains the Vite build output.
_STATIC_DIR = Path(__file__).parent / "static"
_INDEX_HTML  = _STATIC_DIR / "index.html"

if _STATIC_DIR.exists():
    # Serve all static assets (JS, CSS, images, fonts…)
    app.mount("/assets", StaticFiles(directory=str(_STATIC_DIR / "assets")), name="assets")

    # Serve any other static root files (robots.txt, sitemap.xml, favicon…)
    @app.get("/favicon.svg", include_in_schema=False)
    async def favicon():
        return FileResponse(str(_STATIC_DIR / "favicon.svg"))

    @app.get("/robots.txt", include_in_schema=False)
    async def robots():
        return FileResponse(str(_STATIC_DIR / "robots.txt"), media_type="text/plain")

    @app.get("/sitemap.xml", include_in_schema=False)
    async def sitemap():
        return FileResponse(str(_STATIC_DIR / "sitemap.xml"), media_type="application/xml")

    # SPA catch-all — serves index.html for every route not matched above
    @app.get("/{full_path:path}", include_in_schema=False)
    async def serve_spa(full_path: str):
        return HTMLResponse(content=_INDEX_HTML.read_text(), status_code=200)


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)

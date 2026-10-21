from __future__ import annotations

import asyncio
import logging
import os
import re
import shutil
import tempfile
import uuid
from contextlib import asynccontextmanager
from pathlib import Path
from typing import cast

import fitz
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from .converter import ConversionError, inspect_pdf
from .page_policy import PagePolicy
from .jobs import (
    JobQueueFull,
    JOB_ROOT,
    cleanup_expired,
    cleanup_job,
    create_job,
    get_download,
    get_job,
    recover_jobs,
    release_capacity,
    reserve_capacity,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
MAX_BYTES = 30 * 1024 * 1024
MAX_PAGES = 100
STATIC_DIR = Path(__file__).resolve().parent / "static"


async def _reaper() -> None:
    while True:
        await asyncio.sleep(60)
        cleanup_expired()


@asynccontextmanager
async def lifespan(_app: FastAPI):
    recover_jobs()
    task = asyncio.create_task(_reaper())
    yield
    task.cancel()
    try:
        await task
    except asyncio.CancelledError:
        pass
    cleanup_expired()


app = FastAPI(title="PDF to Word Pro API", docs_url=None, redoc_url=None, lifespan=lifespan)


# Vercel frontend URL(s), comma-separated; set on Render after Vercel deploy.
allowed_origins = [origin.strip().rstrip('/') for origin in os.getenv('FRONTEND_ORIGINS', '').split(',') if origin.strip()]
if allowed_origins:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=allowed_origins,
        allow_methods=['GET', 'POST'],
        allow_headers=['Content-Type'],
        max_age=600,
    )


@app.get("/api/health")
def health():
    return {"ok": True, "service": "pdf-to-word-pro"}


@app.post("/api/conversions", status_code=202)
async def start_conversion(
    file: UploadFile = File(...),
    mode: str = Form("editable"),
    page_policy: str = Form("source"),
):
    if mode not in {"layout", "editable"}:
        raise HTTPException(400, "Chế độ chuyển đổi không hợp lệ.")
    if page_policy not in {"source", "administrative"}:
        raise HTTPException(400, "Chế độ khổ giấy và thể thức không hợp lệ.")
    filename = (file.filename or "document.pdf").replace("\\", "/").split("/")[-1]
    filename = re.sub(r"[\x00-\x1f\x7f]", "", filename).strip()
    filename = filename or "document.pdf"
    if not filename.lower().endswith(".pdf"):
        await file.close()
        raise HTTPException(415, "Vui lòng chọn tệp có phần mở rộng PDF.")
    if len(filename) > 160:
        filename = filename[:156] + filename[-4:]

    reservation = reserve_capacity()
    if reservation is None:
        await file.close()
        raise HTTPException(503, "Máy chủ đang xử lý đủ số tác vụ. Vui lòng chờ một lúc rồi thử lại.")

    job_id = uuid.uuid4().hex
    workdir: Path | None = None
    size = 0
    try:
        JOB_ROOT.mkdir(parents=True, exist_ok=True)
        workdir = Path(tempfile.mkdtemp(prefix="job-", dir=JOB_ROOT))
        workdir.chmod(0o700)
        source = workdir / "source.pdf"
        output = workdir / "result.docx"
        with source.open("wb") as destination:
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_BYTES:
                    raise HTTPException(413, "PDF vượt quá giới hạn 30 MiB.")
                destination.write(chunk)
        with source.open("rb") as uploaded_pdf:
            signature = uploaded_pdf.read(5)
        if size < 8 or signature != b"%PDF-":
            raise HTTPException(415, "Tệp tải lên không có cấu trúc PDF hợp lệ.")
        try:
            info = inspect_pdf(source, max_pages=MAX_PAGES)
        except ConversionError as exc:
            raise HTTPException(422, str(exc)) from exc
        try:
            job = create_job(
                job_id,
                workdir,
                source,
                output,
                filename,
                mode,
                cast(PagePolicy, page_policy),
                info["pages"],
                reservation=reservation,
            )
        except JobQueueFull as exc:
            raise HTTPException(503, str(exc)) from exc
        return job
    except Exception:
        if workdir is not None:
            shutil.rmtree(workdir, ignore_errors=True)
        raise
    finally:
        release_capacity(reservation)
        await file.close()


@app.get("/api/conversions/{job_id}")
def conversion_status(job_id: str):
    job = get_job(job_id)
    if not job:
        raise HTTPException(404, "Không tìm thấy phiên chuyển đổi. Hãy tải tài liệu lên lại.")
    return job


@app.get("/api/conversions/{job_id}/download")
def download_conversion(job_id: str):
    result = get_download(job_id)
    if not result:
        job = get_job(job_id)
        if not job:
            raise HTTPException(404, "Tệp Word không còn khả dụng. Hãy chuyển đổi lại tài liệu.")
        raise HTTPException(409, "Tệp Word chưa sẵn sàng để tải xuống.")
    path, filename, _size = result
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
        background=BackgroundTask(cleanup_job, job_id),
    )


if STATIC_DIR.is_dir() and (STATIC_DIR / "index.html").exists():
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="web")
else:
    @app.get("/")
    def not_built():
        return JSONResponse({"message": "Frontend chưa được build. Chạy pnpm build trước khi mở trang web."}, status_code=503)

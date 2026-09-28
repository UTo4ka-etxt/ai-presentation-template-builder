from __future__ import annotations

import json
import os
import threading
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path, PurePosixPath
from typing import Annotated, Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from dotenv import load_dotenv

from .pipeline import GenerationPipeline

load_dotenv()

DATA_DIR = Path(os.getenv("PRESENTATION_SERVICE_DATA_DIR", "service_data")).resolve()
MAX_TEMPLATE_BYTES = int(os.getenv("PRESENTATION_MAX_TEMPLATE_BYTES", str(100 * 1024 * 1024)))
MAX_CONTENT_BYTES = int(os.getenv("PRESENTATION_MAX_CONTENT_BYTES", str(500 * 1024 * 1024)))
EXECUTOR = ThreadPoolExecutor(max_workers=int(os.getenv("PRESENTATION_WORKERS", "2")))
LOCK = threading.Lock()
JOBS: dict[str, dict[str, Any]] = {}

app = FastAPI(title="Presentation Generator", version="1.0.0")


def _endpoint_configured(base_url: str | None, api_key: str | None) -> bool:
    if not base_url:
        return False
    return bool(api_key) if "openrouter.ai" in base_url.lower() else True


def _save_status(job_id: str) -> None:
    job_dir = DATA_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    (job_dir / "job.json").write_text(json.dumps(JOBS[job_id], ensure_ascii=False, indent=2), encoding="utf-8")


def _update(job_id: str, **values: Any) -> None:
    with LOCK:
        JOBS[job_id].update(values)
        _save_status(job_id)


def _safe_extract(archive_path: Path, destination: Path) -> None:
    total = 0
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive_path) as archive:
        for member in archive.infolist():
            relative = PurePosixPath(member.filename)
            if relative.is_absolute() or ".." in relative.parts:
                raise ValueError(f"Unsafe ZIP path: {member.filename}")
            total += member.file_size
            if total > MAX_CONTENT_BYTES:
                raise ValueError("Unpacked content package exceeds the configured size limit")
            target = (destination / Path(*relative.parts)).resolve()
            if destination.resolve() not in target.parents and target != destination.resolve():
                raise ValueError(f"Unsafe ZIP path: {member.filename}")
            if member.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, target.open("wb") as output:
                    while chunk := source.read(1024 * 1024):
                        output.write(chunk)


def _run_job(job_id: str, template: Path, content_dir: Path, brief: str, slide_count: int) -> None:
    try:
        _update(job_id, status="running", stage="decompose", percent=1, message="Запуск обработки")
        pipeline = GenerationPipeline(
            model=os.getenv("PRESENTATION_LLM_MODEL"),
            vision_model=os.getenv("PRESENTATION_VLM_MODEL") or None,
        )

        def progress(stage: str, percent: int, message: str) -> None:
            _update(job_id, stage=stage, percent=percent, message=message)

        report = pipeline.run(
            template_path=template,
            content_path=content_dir,
            brief=brief,
            output_dir=DATA_DIR / job_id / "output",
            slide_count=slide_count,
            variants=3,
            progress=progress,
        )
        downloads = {
            str(item["variant"]): f"/api/jobs/{job_id}/download/{item['variant']}"
            for item in report["variants"]
        }
        _update(job_id, status="complete", stage="complete", percent=100, message="Готово", downloads=downloads)
    except Exception as exc:
        _update(job_id, status="failed", stage="failed", message=str(exc))


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return _INDEX_HTML


@app.get("/api/health")
def health() -> dict[str, Any]:
    llm_url = os.getenv("PRESENTATION_LLM_BASE_URL")
    llm_key = os.getenv("PRESENTATION_LLM_API_KEY") or os.getenv("OPENROUTER_API_KEY")
    vlm_url = os.getenv("PRESENTATION_VLM_BASE_URL") or llm_url
    vlm_key = os.getenv("PRESENTATION_VLM_API_KEY") or llm_key
    return {
        "status": "ok",
        "llm_configured": _endpoint_configured(llm_url, llm_key),
        "vlm_configured": bool(os.getenv("PRESENTATION_VLM_MODEL")) and _endpoint_configured(vlm_url, vlm_key),
        "model": os.getenv("PRESENTATION_LLM_MODEL", "Qwen/Qwen3-30B-A3B-Instruct-2507"),
        "vision_model": os.getenv("PRESENTATION_VLM_MODEL"),
    }


@app.post("/api/jobs", status_code=202)
async def create_job(
    template: Annotated[UploadFile, File(description="PowerPoint template (.pptx)")],
    content_package: Annotated[UploadFile, File(description="ZIP with text, images and tables")],
    brief: Annotated[str, Form(min_length=1, max_length=20000)],
    slide_count: Annotated[int, Form(ge=3, le=20)] = 10,
) -> dict[str, Any]:
    llm_url = os.getenv("PRESENTATION_LLM_BASE_URL")
    llm_key = os.getenv("PRESENTATION_LLM_API_KEY") or os.getenv("OPENROUTER_API_KEY")
    if not _endpoint_configured(llm_url, llm_key):
        raise HTTPException(503, "The network LLM endpoint or its API key is not configured")
    if not template.filename or not template.filename.lower().endswith(".pptx"):
        raise HTTPException(400, "Template must be a .pptx file")
    if not content_package.filename or not content_package.filename.lower().endswith(".zip"):
        raise HTTPException(400, "Content package must be a .zip file")
    job_id = uuid.uuid4().hex
    job_dir = DATA_DIR / job_id
    input_dir = job_dir / "input"
    input_dir.mkdir(parents=True, exist_ok=False)
    template_path = input_dir / "template.pptx"
    content_zip = input_dir / "content.zip"
    template_data = await template.read(MAX_TEMPLATE_BYTES + 1)
    content_data = await content_package.read(MAX_CONTENT_BYTES + 1)
    if len(template_data) > MAX_TEMPLATE_BYTES or len(content_data) > MAX_CONTENT_BYTES:
        raise HTTPException(413, "Uploaded files exceed the configured size limit")
    template_path.write_bytes(template_data)
    content_zip.write_bytes(content_data)
    content_dir = input_dir / "content"
    try:
        _safe_extract(content_zip, content_dir)
    except (ValueError, zipfile.BadZipFile) as exc:
        raise HTTPException(400, str(exc)) from exc
    with LOCK:
        JOBS[job_id] = {
            "job_id": job_id,
            "status": "queued",
            "stage": "queued",
            "percent": 0,
            "message": "Задание принято",
            "slide_count": slide_count,
            "variants": 3,
        }
        _save_status(job_id)
    EXECUTOR.submit(_run_job, job_id, template_path, content_dir, brief, slide_count)
    return JOBS[job_id]


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str) -> dict[str, Any]:
    if job_id not in JOBS:
        status_path = DATA_DIR / job_id / "job.json"
        if not status_path.exists():
            raise HTTPException(404, "Job not found")
        JOBS[job_id] = json.loads(status_path.read_text(encoding="utf-8"))
    return JOBS[job_id]


@app.get("/api/jobs/{job_id}/download/{variant}")
def download(job_id: str, variant: int) -> FileResponse:
    if variant not in (1, 2, 3):
        raise HTTPException(404, "Variant not found")
    status = get_job(job_id)
    if status.get("status") != "complete":
        raise HTTPException(409, "Job is not complete")
    path = DATA_DIR / job_id / "output" / f"variant-{variant}.pptx"
    if not path.exists():
        raise HTTPException(404, "Result file not found")
    return FileResponse(path, filename=f"presentation-variant-{variant}.pptx", media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")


def main() -> None:
    import uvicorn

    uvicorn.run(
        "presentation_builder.web:app",
        host=os.getenv("PRESENTATION_HOST", "0.0.0.0"),
        port=int(os.getenv("PRESENTATION_PORT", "8000")),
        reload=False,
    )


_INDEX_HTML = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Генератор презентаций</title><style>
body{font-family:Arial,sans-serif;background:#090b10;color:#f4f7fb;margin:0}main{max-width:760px;margin:48px auto;padding:32px;background:#151922;border:1px solid #283041;border-radius:18px}h1{margin-top:0}label{display:block;margin:18px 0 7px;color:#b9c4d5}input,textarea,button{box-sizing:border-box;width:100%;font:inherit}input,textarea{padding:12px;background:#0d1118;color:white;border:1px solid #354057;border-radius:8px}textarea{min-height:130px}button{margin-top:22px;padding:14px;border:0;border-radius:9px;background:#087cff;color:white;font-weight:700;cursor:pointer}.status{margin-top:24px;white-space:pre-wrap}.downloads a{display:block;color:#51b5ff;margin:8px 0}</style></head>
<body><main><h1>Генератор презентаций</h1><p>Загрузите корпоративный шаблон PPTX и ZIP с текстами, изображениями и таблицами. Сервис подготовит три варианта.</p>
<form id="form"><label>Шаблон PPTX</label><input name="template" type="file" accept=".pptx" required><label>Контент-пакет ZIP</label><input name="content_package" type="file" accept=".zip" required><label>Задача и аудитория</label><textarea name="brief" required></textarea><label>Количество слайдов</label><input name="slide_count" type="number" min="3" max="20" value="10"><button>Создать три варианта</button></form><div class="status" id="status"></div><div class="downloads" id="downloads"></div></main>
<script>const f=document.querySelector('#form'),s=document.querySelector('#status'),d=document.querySelector('#downloads');f.onsubmit=async e=>{e.preventDefault();d.innerHTML='';s.textContent='Загрузка...';const r=await fetch('/api/jobs',{method:'POST',body:new FormData(f)});const j=await r.json();if(!r.ok){s.textContent=j.detail||'Ошибка';return}poll(j.job_id)};async function poll(id){const r=await fetch('/api/jobs/'+id),j=await r.json();s.textContent=`${j.message}\n${j.percent||0}%`;if(j.status==='complete'){d.innerHTML=Object.entries(j.downloads).map(([n,u])=>`<a href="${u}">Скачать вариант ${n}</a>`).join('');return}if(j.status==='failed')return;setTimeout(()=>poll(id),2000)}</script></body></html>"""

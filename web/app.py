from __future__ import annotations

import asyncio
import json
import os
import queue
import shutil
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import BackgroundTasks, FastAPI, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from check_blogs.check_blogs import process_excel
from processor.process_backlinks import process_backlink_files


ROOT_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = Path(__file__).resolve().parent / "static"
TERMINAL_STATUSES = {"completed", "failed"}


@dataclass(frozen=True)
class AppConfig:
    storage_root: Path = field(
        default_factory=lambda: Path(os.getenv("BACKLINK_STORAGE_DIR", ROOT_DIR / "storage"))
    )
    max_upload_bytes: int = field(
        default_factory=lambda: int(os.getenv("BACKLINK_MAX_UPLOAD_MB", "100")) * 1024 * 1024
    )
    worker_count: int = field(default_factory=lambda: max(1, int(os.getenv("BACKLINK_WORKERS", "1"))))
    check_blogs_concurrency: int = field(default_factory=lambda: max(1, int(os.getenv("CHECK_BLOGS_CONCURRENCY", "16"))))
    check_blogs_domain_interval: float = field(default_factory=lambda: float(os.getenv("CHECK_BLOGS_DOMAIN_INTERVAL_SECONDS", "1.5")))
    checkpoint_batch_size: int = field(default_factory=lambda: max(1, int(os.getenv("CHECK_BLOGS_CHECKPOINT_BATCH_SIZE", "25"))))

    @property
    def jobs_dir(self) -> Path:
        return self.storage_root / "jobs"

    @property
    def cache_path(self) -> Path:
        return self.storage_root / "cache" / "domain_results.sqlite3"


@dataclass
class JobRecord:
    job_id: str
    status: str
    stage: str
    progress: float
    created_at: str
    updated_at: str
    files: list[dict[str, Any]]
    logs: list[dict[str, str]] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)
    error: str | None = None
    download_path: str | None = None

    def public_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["download_ready"] = self.status == "completed" and bool(self.download_path)
        payload.pop("download_path", None)
        return payload


class JobManager:
    def __init__(self, config: AppConfig):
        self.config = config
        self.jobs: dict[str, JobRecord] = {}
        self.event_queues: dict[str, queue.Queue[dict[str, Any]]] = {}
        self.lock = threading.RLock()
        self.executor = ThreadPoolExecutor(max_workers=config.worker_count)
        self.recoverable_job_ids: list[str] = []
        self.restart_updated_job_ids: list[str] = []
        self.config.jobs_dir.mkdir(parents=True, exist_ok=True)
        self._load_existing_jobs()

    def _job_dir(self, job_id: str) -> Path:
        return self.config.jobs_dir / job_id

    def _metadata_path(self, job_id: str) -> Path:
        return self._job_dir(job_id) / "metadata.json"

    def _load_existing_jobs(self) -> None:
        for metadata_path in sorted(self.config.jobs_dir.glob("*/metadata.json")):
            try:
                payload = json.loads(metadata_path.read_text(encoding="utf-8"))
                record = JobRecord(**payload)
                if record.status not in TERMINAL_STATUSES:
                    job_dir = metadata_path.parent
                    has_uploads = any((job_dir / "uploads").glob("*.xlsx"))
                    has_processor_output = (job_dir / "processor" / "processed_backlinks.xlsx").exists()
                    if has_uploads or has_processor_output:
                        record.status = "queued"
                        record.error = None
                        record.logs.append({
                            "time": datetime.now().strftime("%H:%M:%S"),
                            "level": "INFO",
                            "message": "服务重启，任务已加入断点续跑队列。",
                        })
                        self.recoverable_job_ids.append(record.job_id)
                    else:
                        record.status = "failed"
                        record.error = "服务重启后缺少上传文件和中间结果，无法续跑。"
                    record.updated_at = datetime.now().isoformat(timespec="seconds")
                    self.restart_updated_job_ids.append(record.job_id)
                self.jobs[record.job_id] = record
            except Exception:
                continue

    def _persist(self, record: JobRecord) -> None:
        path = self._metadata_path(record.job_id)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(asdict(record), ensure_ascii=False, indent=2), encoding="utf-8")

    def _publish(self, job_id: str, event_type: str, payload: dict[str, Any]) -> None:
        event = {"type": event_type, "payload": payload}
        with self.lock:
            event_queue = self.event_queues.setdefault(job_id, queue.Queue())
        event_queue.put(event)

    def _snapshot(self, record: JobRecord) -> dict[str, Any]:
        return record.public_dict()

    def get_job(self, job_id: str) -> JobRecord:
        with self.lock:
            record = self.jobs.get(job_id)
            if not record:
                raise KeyError(job_id)
            return record

    def list_jobs(self, limit: int = 20) -> list[dict[str, Any]]:
        with self.lock:
            records = sorted(self.jobs.values(), key=lambda item: item.created_at, reverse=True)
            return [record.public_dict() for record in records[:limit]]

    def update_job(
        self,
        job_id: str,
        *,
        status: str | None = None,
        stage: str | None = None,
        progress: float | None = None,
        stats: dict[str, Any] | None = None,
        error: str | None = None,
        download_path: str | None = None,
    ) -> None:
        with self.lock:
            record = self.jobs[job_id]
            if status is not None:
                record.status = status
            if stage is not None:
                record.stage = stage
            if progress is not None:
                record.progress = round(max(0.0, min(100.0, progress)), 1)
            if stats:
                record.stats.update(stats)
            if error is not None:
                record.error = error
            if download_path is not None:
                record.download_path = download_path
            record.updated_at = datetime.now().isoformat(timespec="seconds")
            self._persist(record)
            snapshot = self._snapshot(record)
        self._publish(job_id, "snapshot", snapshot)

    def add_log(self, job_id: str, message: str, level: str = "INFO") -> None:
        entry = {
            "time": datetime.now().strftime("%H:%M:%S"),
            "level": level,
            "message": str(message),
        }
        with self.lock:
            record = self.jobs[job_id]
            record.logs.append(entry)
            record.logs = record.logs[-300:]
            record.updated_at = datetime.now().isoformat(timespec="seconds")
            self._persist(record)
        self._publish(job_id, "log", entry)

    def create_job(self, files: list[dict[str, Any]]) -> JobRecord:
        now = datetime.now().isoformat(timespec="seconds")
        job_id = datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:8]
        record = JobRecord(
            job_id=job_id,
            status="queued",
            stage="upload",
            progress=5,
            created_at=now,
            updated_at=now,
            files=files,
        )
        with self.lock:
            self.jobs[job_id] = record
            self.event_queues[job_id] = queue.Queue()
            self._persist(record)
        self._publish(job_id, "snapshot", record.public_dict())
        return record

    def enqueue(self, job_id: str) -> None:
        self.executor.submit(self._run_job, job_id)

    def _run_job(self, job_id: str) -> None:
        job_dir = self._job_dir(job_id)
        upload_dir = job_dir / "uploads"
        processor_dir = job_dir / "processor"
        result_dir = job_dir / "result"
        processor_output = processor_dir / "processed_backlinks.xlsx"
        final_output = result_dir / f"check_blogs_result_{job_id}.xlsx"

        try:
            record = self.get_job(job_id)
            if processor_output.exists():
                self.update_job(job_id, status="running", stage="check_blogs", progress=max(record.progress, 48))
                self.add_log(job_id, "检测到 processor 中间结果，跳过重复清洗")
            else:
                if not upload_dir.exists() or not any(upload_dir.glob("*.xlsx")):
                    raise RuntimeError("缺少上传文件，无法继续任务")
                self.update_job(job_id, status="running", stage="processor", progress=10)
                self.add_log(job_id, "开始 processor 清洗")

                def processor_progress(percent: float, message: str) -> None:
                    self.update_job(job_id, stage="processor", progress=10 + percent * 0.35)
                    if message:
                        self.add_log(job_id, message)

                processor_result = process_backlink_files(
                    source_dir=upload_dir,
                    output_file=processor_output,
                    use_processed_log=False,
                    logger=lambda message: self.add_log(job_id, message),
                    progress_callback=processor_progress,
                )
                self.update_job(job_id, stats={"processor": processor_result.to_dict()})
            if not processor_output.exists():
                raise RuntimeError("processor 未生成清洗后的 Excel 文件")

            self.update_job(job_id, stage="check_blogs", progress=48)
            self.add_log(job_id, "开始博客评论检测")

            def check_blogs_progress(percent: float, message: str) -> None:
                self.update_job(job_id, stage="check_blogs", progress=48 + percent * 0.42)
                if message:
                    self.add_log(job_id, message)

            def check_blogs_stats(stats: dict[str, Any]) -> None:
                self.update_job(job_id, stats={"check_blogs": stats})

            check_blogs_result = process_excel(
                str(processor_output),
                output_path=str(final_output),
                resume=True,
                logger=lambda message: self.add_log(job_id, message),
                progress_callback=check_blogs_progress,
                max_workers=self.config.check_blogs_concurrency,
                checkpoint_batch_size=self.config.checkpoint_batch_size,
                cache_path=str(self.config.cache_path),
                domain_interval=self.config.check_blogs_domain_interval,
                stats_callback=check_blogs_stats,
            )

            if not final_output.exists():
                raise RuntimeError("check_blogs 未生成最终导出文件")
            self.update_job(job_id, stage="export", progress=96, stats={"check_blogs": check_blogs_result.to_dict()})
            self.add_log(job_id, "导出结果已生成")
            self.update_job(
                job_id,
                status="completed",
                stage="done",
                progress=100,
                download_path=str(final_output),
            )
        except Exception as exc:
            self.add_log(job_id, f"任务失败：{exc}", level="ERROR")
            self.update_job(job_id, status="failed", stage="failed", error=str(exc))

    async def event_stream(self, job_id: str):
        try:
            record = self.get_job(job_id)
        except KeyError:
            raise HTTPException(status_code=404, detail="Job not found")

        yield self.format_sse({"type": "snapshot", "payload": record.public_dict()})
        for entry in record.logs:
            yield self.format_sse({"type": "log", "payload": entry})

        with self.lock:
            event_queue = self.event_queues.setdefault(job_id, queue.Queue())

        while True:
            drained = False
            while True:
                try:
                    event = event_queue.get_nowait()
                except queue.Empty:
                    break
                drained = True
                yield self.format_sse(event)

            current = self.get_job(job_id)
            if current.status in TERMINAL_STATUSES and not drained:
                yield self.format_sse({"type": "snapshot", "payload": current.public_dict()})
                break
            await asyncio.sleep(0.5)

    @staticmethod
    def format_sse(event: dict[str, Any]) -> str:
        return f"data: {json.dumps(event, ensure_ascii=False)}\n\n"


def safe_filename(filename: str) -> str:
    name = Path(filename or "upload.xlsx").name.strip()
    return name or "upload.xlsx"


async def save_upload(upload: UploadFile, destination: Path, max_bytes: int) -> int:
    destination.parent.mkdir(parents=True, exist_ok=True)
    total = 0
    with destination.open("wb") as handle:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > max_bytes:
                handle.close()
                destination.unlink(missing_ok=True)
                raise HTTPException(status_code=413, detail=f"{upload.filename} 超过单文件大小限制")
            handle.write(chunk)
    return total


def create_app(config: AppConfig | None = None) -> FastAPI:
    active_config = config or AppConfig()
    app = FastAPI(title="外链清洗与筛选", version="1.0.0")
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    manager = JobManager(active_config)
    app.state.job_manager = manager

    @app.on_event("startup")
    def resume_interrupted_jobs() -> None:
        for updated_job_id in manager.restart_updated_job_ids:
            manager._persist(manager.get_job(updated_job_id))
        for recoverable_job_id in manager.recoverable_job_ids:
            manager.enqueue(recoverable_job_id)

    @app.get("/api/health")
    def health() -> dict[str, Any]:
        return {
            "ok": True,
            "storage_root": str(active_config.storage_root),
            "pipeline": "processor->check_blogs",
        }

    @app.get("/api/jobs")
    def list_jobs() -> list[dict[str, Any]]:
        return manager.list_jobs()

    @app.post("/api/jobs")
    async def create_job(files: list[UploadFile], background_tasks: BackgroundTasks) -> JSONResponse:
        if not files:
            raise HTTPException(status_code=400, detail="请至少上传一个 XLSX 文件")
        invalid = [file.filename for file in files if not safe_filename(file.filename or "").lower().endswith(".xlsx")]
        if invalid:
            raise HTTPException(status_code=400, detail=f"仅支持 XLSX 文件：{', '.join(invalid)}")

        temp_job_id = "pending-" + uuid.uuid4().hex[:8]
        temp_upload_dir = active_config.jobs_dir / temp_job_id / "uploads"
        saved_files: list[dict[str, Any]] = []
        used_names: set[str] = set()
        try:
            for upload in files:
                filename = safe_filename(upload.filename or "")
                if filename in used_names:
                    filename = f"{Path(filename).stem}_{len(used_names) + 1}{Path(filename).suffix}"
                used_names.add(filename)
                destination = temp_upload_dir / filename
                size = await save_upload(upload, destination, active_config.max_upload_bytes)
                saved_files.append({"name": filename, "size": size})

            record = manager.create_job(saved_files)
            final_job_dir = active_config.jobs_dir / record.job_id
            if final_job_dir.exists():
                shutil.rmtree(final_job_dir)
            shutil.move(str(temp_upload_dir.parent), str(final_job_dir))
            manager._persist(record)
            background_tasks.add_task(manager.enqueue, record.job_id)
            return JSONResponse(record.public_dict(), status_code=201)
        except Exception:
            shutil.rmtree(temp_upload_dir.parent, ignore_errors=True)
            raise

    @app.get("/api/jobs/{job_id}")
    def get_job(job_id: str) -> dict[str, Any]:
        try:
            return manager.get_job(job_id).public_dict()
        except KeyError:
            raise HTTPException(status_code=404, detail="Job not found")

    @app.get("/api/jobs/{job_id}/events")
    async def job_events(job_id: str):
        try:
            manager.get_job(job_id)
        except KeyError:
            raise HTTPException(status_code=404, detail="Job not found")
        return StreamingResponse(manager.event_stream(job_id), media_type="text/event-stream")

    @app.get("/api/jobs/{job_id}/download")
    def download_job(job_id: str) -> FileResponse:
        try:
            record = manager.get_job(job_id)
        except KeyError:
            raise HTTPException(status_code=404, detail="Job not found")
        if record.status != "completed" or not record.download_path:
            raise HTTPException(status_code=409, detail="任务尚未完成")
        path = Path(record.download_path)
        if not path.exists():
            raise HTTPException(status_code=404, detail="导出文件不存在")
        return FileResponse(
            path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=path.name,
        )

    if STATIC_DIR.exists():
        assets_dir = STATIC_DIR / "assets"
        if assets_dir.exists():
            app.mount("/assets", StaticFiles(directory=assets_dir), name="assets")

        @app.get("/{full_path:path}", response_class=HTMLResponse)
        async def serve_frontend(full_path: str, request: Request) -> HTMLResponse:
            if full_path.startswith("api/"):
                raise HTTPException(status_code=404, detail="Not found")
            index_path = STATIC_DIR / "index.html"
            return HTMLResponse(index_path.read_text(encoding="utf-8"))

    return app


app = create_app()

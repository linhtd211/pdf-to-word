from __future__ import annotations

import json
import logging
import multiprocessing
import os
import resource
import signal
import shutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from queue import Empty
from typing import Any, cast

from .converter import ConversionError, convert_pdf
from .page_policy import PagePolicy

logger = logging.getLogger(__name__)
TTL_SECONDS = 20 * 60
MAX_WORKERS = 2
MAX_QUEUED_JOBS = 4
MAX_ACCEPTED_JOBS = MAX_WORKERS + MAX_QUEUED_JOBS
CONVERSION_TIMEOUT_SECONDS = 5 * 60
MAX_WORKER_MEMORY_MIB = int(os.environ.get("PDFWORD_WORKER_MEMORY_MIB", "1536"))
MAX_OUTPUT_BYTES = 100 * 1024 * 1024
JOB_ROOT = Path(os.environ.get("PDFWORD_JOB_ROOT", "/tmp/pdfword-jobs"))
JOB_ROOT.mkdir(parents=True, exist_ok=True)
executor = ThreadPoolExecutor(max_workers=MAX_WORKERS, thread_name_prefix="pdf-convert")
_slots = threading.BoundedSemaphore(MAX_ACCEPTED_JOBS)
_jobs: dict[str, dict[str, Any]] = {}
_lock = threading.RLock()


class JobQueueFull(Exception):
    """Raised when both the active worker pool and bounded queue are occupied."""


class SlotReservation:
    def __init__(self) -> None:
        self.active = True


def _cpu_limit_reached(_signum: int, _frame: Any) -> None:
    raise TimeoutError("CPU time limit reached")


def _apply_worker_limits() -> None:
    """Apply per-child process caps; the supervisor also enforces a wall-clock timeout."""
    memory_limit = max(256, MAX_WORKER_MEMORY_MIB) * 1024 * 1024
    try:
        _, memory_hard = resource.getrlimit(resource.RLIMIT_AS)
        if memory_hard != resource.RLIM_INFINITY:
            memory_limit = min(memory_limit, int(memory_hard))
        resource.setrlimit(resource.RLIMIT_AS, (memory_limit, memory_limit))
        cpu_soft, cpu_hard = resource.getrlimit(resource.RLIMIT_CPU)
        requested_soft = max(30, CONVERSION_TIMEOUT_SECONDS - 10)
        requested_hard = CONVERSION_TIMEOUT_SECONDS
        if cpu_hard != resource.RLIM_INFINITY:
            requested_soft = min(requested_soft, int(cpu_hard))
            requested_hard = min(requested_hard, int(cpu_hard))
        signal.signal(signal.SIGXCPU, _cpu_limit_reached)
        resource.setrlimit(resource.RLIMIT_CPU, (requested_soft, requested_hard))
    except (OSError, ValueError, resource.error):
        logger.warning("Could not apply worker resource limits", exc_info=True)


def reserve_capacity() -> SlotReservation | None:
    if not _slots.acquire(blocking=False):
        return None
    return SlotReservation()


def release_capacity(reservation: SlotReservation | None) -> None:
    if reservation is not None and reservation.active:
        reservation.active = False
        _slots.release()


def _manifest_path(job: dict[str, Any]) -> Path:
    return Path(job["workdir"]) / "job.json"


def _persist(job: dict[str, Any]) -> None:
    try:
        path = _manifest_path(job)
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(".json.tmp")
        temporary.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
        temporary.replace(path)
    except OSError:
        logger.warning("Could not persist conversion job %s", job.get("id"), exc_info=True)


def _update(job_id: str, fields: dict[str, Any]) -> None:
    snapshot = None
    with _lock:
        job = _jobs.get(job_id)
        if job:
            job.update(fields)
            snapshot = dict(job)
    if snapshot:
        _persist(snapshot)


def _conversion_worker(
    message_queue,
    source: str,
    output: str,
    mode: str,
    page_policy: str,
) -> None:
    try:
        _apply_worker_limits()
        def report(percent: int, stage: str) -> None:
            message_queue.put({"type": "progress", "progress": percent, "stage": stage})

        result = convert_pdf(Path(source), Path(output), mode, report, page_policy=cast(PagePolicy, page_policy))
        message_queue.put({"type": "complete", "result": result})
    except TimeoutError:
        message_queue.put({"type": "error", "error": "Tác vụ đã chạm giới hạn thời gian CPU. Hãy thử PDF ngắn hơn hoặc giảm độ phân giải scan.", "expected": True})
    except MemoryError:
        message_queue.put({"type": "error", "error": "PDF vượt giới hạn bộ nhớ của worker. Hãy thử tài liệu ít trang hoặc hình ảnh nhẹ hơn.", "expected": True})
    except ConversionError as exc:
        message_queue.put({"type": "error", "error": str(exc), "expected": True})
    except BaseException:
        logger.exception("Unhandled conversion worker failure")
        message_queue.put({"type": "error", "error": "Đã xảy ra lỗi khi chuyển đổi. Vui lòng thử lại với một PDF khác.", "expected": False})


def _run(job_id: str) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            _slots.release()
            return
        job.update({"status": "processing", "stage": "Đọc cấu trúc PDF", "progress": max(5, int(job.get("progress", 0))), "updated_at": time.time()})
        snapshot = dict(job)
    _persist(snapshot)

    process = None
    message_queue = None
    terminal = False
    try:
        context = multiprocessing.get_context("spawn")
        message_queue = context.Queue()
        process = context.Process(
            target=_conversion_worker,
            args=(message_queue, snapshot["source"], snapshot["output"], snapshot["mode"], snapshot["page_policy"]),
            name=f"pdf-convert-{job_id[:8]}",
        )
        process.start()
        deadline = time.monotonic() + CONVERSION_TIMEOUT_SECONDS
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                process.terminate()
                process.join(timeout=3)
                if process.is_alive():
                    process.kill()
                    process.join(timeout=2)
                _update(job_id, {
                    "status": "failed",
                    "stage": "Quá thời gian xử lý",
                    "error": "Tác vụ vượt quá giới hạn 5 phút và đã được dừng để bảo vệ máy chủ. Hãy thử PDF ít trang hơn hoặc giảm độ phân giải scan.",
                    "updated_at": time.time(),
                })
                terminal = True
                break
            try:
                message = message_queue.get(timeout=min(0.5, remaining))
            except Empty:
                if process.is_alive():
                    continue
                _update(job_id, {
                    "status": "failed",
                    "stage": "Worker bị gián đoạn",
                    "error": "Tiến trình chuyển đổi bị dừng trước khi hoàn tất. Hãy tải PDF lên để thử lại.",
                    "updated_at": time.time(),
                })
                terminal = True
                break

            if message.get("type") == "progress":
                _update(job_id, {
                    "progress": message.get("progress", 0),
                    "stage": message.get("stage", "Đang xử lý"),
                    "updated_at": time.time(),
                })
            elif message.get("type") == "error":
                _update(job_id, {
                    "status": "failed",
                    "stage": "Chuyển đổi thất bại",
                    "error": message.get("error") or "Đã xảy ra lỗi khi chuyển đổi.",
                    "updated_at": time.time(),
                })
                terminal = True
                break
            elif message.get("type") == "complete":
                process.join(timeout=3)
                output_path = Path(snapshot["output"])
                if not output_path.is_file() or output_path.stat().st_size == 0:
                    raise RuntimeError("Conversion worker returned without a DOCX output")
                if output_path.stat().st_size > MAX_OUTPUT_BYTES:
                    output_path.unlink(missing_ok=True)
                    _update(job_id, {
                        "status": "failed",
                        "stage": "Tệp đầu ra quá lớn",
                        "error": "Tệp Word vượt giới hạn lưu tạm 100 MiB. Hãy thử tài liệu ít trang/hình hơn.",
                        "updated_at": time.time(),
                    })
                    terminal = True
                    break
                result = message.get("result", {})
                _update(job_id, {
                    "status": "completed",
                    "progress": 100,
                    "stage": "Hoàn tất",
                    "updated_at": time.time(),
                    "size": output_path.stat().st_size,
                    "warnings": result.get("warnings", []),
                    "kind": result.get("kind"),
                    "quality": result.get("quality"),
                })
                terminal = True
                break
    except Exception:
        logger.exception("Conversion supervisor failed for job %s", job_id)
        _update(job_id, {
            "status": "failed",
            "stage": "Worker bị gián đoạn",
            "error": "Máy chủ không thể hoàn tất tác vụ. Hãy tải PDF lên để thử lại.",
            "updated_at": time.time(),
        })
        terminal = True
    finally:
        if process is not None and process.is_alive():
            process.terminate()
            process.join(timeout=2)
            if process.is_alive():
                process.kill()
                process.join(timeout=2)
        if message_queue is not None:
            message_queue.close()
            message_queue.join_thread()
        if not terminal:
            _update(job_id, {
                "status": "failed",
                "stage": "Worker bị gián đoạn",
                "error": "Tác vụ bị gián đoạn. Hãy tải PDF lên để thử lại.",
                "updated_at": time.time(),
            })
        _slots.release()


def _submit(job_id: str) -> bool:
    if not _slots.acquire(blocking=False):
        return False
    try:
        executor.submit(_run, job_id)
        return True
    except Exception:
        _slots.release()
        logger.exception("Could not submit conversion job %s", job_id)
        return False


def create_job(
    job_id: str,
    workdir: Path,
    source: Path,
    output: Path,
    original_name: str,
    mode: str,
    page_policy: str,
    pages: int,
    reservation: SlotReservation | None = None,
) -> dict[str, Any]:
    if reservation is None:
        reservation = reserve_capacity()
    if reservation is None or not reservation.active:
        raise JobQueueFull("Máy chủ đang xử lý đủ số tác vụ. Vui lòng chờ một lúc rồi thử lại.")
    now = time.time()
    job: dict[str, Any] = {
        "id": job_id,
        "status": "queued",
        "progress": 0,
        "stage": "Đang chờ xử lý",
        "created_at": now,
        "updated_at": now,
        "workdir": str(workdir),
        "source": str(source),
        "output": str(output),
        "download_name": f"{Path(original_name).stem}.docx",
        "mode": mode,
        "page_policy": page_policy,
        "pages": pages,
        "size": 0,
        "warnings": [],
        "quality": None,
        "error": None,
    }
    with _lock:
        _jobs[job_id] = job
    _persist(job)
    try:
        executor.submit(_run, job_id)
        reservation.active = False
    except Exception:
        with _lock:
            _jobs.pop(job_id, None)
        release_capacity(reservation)
        raise
    return _public(job)


def recover_jobs() -> int:
    """Restore manifests after an application-process restart and restart interrupted work."""
    JOB_ROOT.mkdir(parents=True, exist_ok=True)
    recovered = 0
    now = time.time()
    for manifest in JOB_ROOT.glob("*/job.json"):
        try:
            job = json.loads(manifest.read_text(encoding="utf-8"))
            if not isinstance(job, dict) or not job.get("id"):
                continue
            if now - float(job.get("updated_at", now)) > TTL_SECONDS:
                shutil.rmtree(manifest.parent, ignore_errors=True)
                continue
            status = job.get("status")
            source_exists = Path(str(job.get("source", ""))).is_file()
            output_exists = Path(str(job.get("output", ""))).is_file()
            if status in {"queued", "processing"} and source_exists:
                job.update({
                    "status": "queued",
                    "progress": 0,
                    "stage": "Đang khôi phục tác vụ sau khi máy chủ khởi động lại",
                    "error": None,
                    "updated_at": now,
                })
            elif status == "completed" and not output_exists:
                job.update({
                    "status": "failed",
                    "stage": "Tệp tạm không còn khả dụng",
                    "error": "Tệp Word tạm đã mất sau khi máy chủ khởi động lại. Hãy tải PDF lên để chuyển đổi lại.",
                    "updated_at": now,
                })
            elif status in {"queued", "processing"}:
                job.update({
                    "status": "failed",
                    "stage": "Tác vụ bị gián đoạn",
                    "error": "Tác vụ bị gián đoạn khi máy chủ khởi động lại và PDF tạm không còn. Hãy tải PDF lên để thử lại.",
                    "updated_at": now,
                })
            with _lock:
                _jobs[job["id"]] = job
            _persist(job)
            if job.get("status") == "queued":
                if not _submit(job["id"]):
                    _update(job["id"], {
                        "status": "failed",
                        "stage": "Không thể khôi phục hàng đợi",
                        "error": "Máy chủ đã khởi động lại nhưng không thể khôi phục hàng đợi. Hãy tải PDF lên để thử lại.",
                        "updated_at": time.time(),
                    })
            recovered += 1
        except Exception:
            logger.warning("Skipping unreadable job manifest %s", manifest, exc_info=True)
    cleanup_orphan_directories(now)
    return recovered


def get_job(job_id: str) -> dict[str, Any] | None:
    with _lock:
        job = _jobs.get(job_id)
        return _public(job) if job else None


def get_download(job_id: str) -> tuple[Path, str, int] | None:
    with _lock:
        job = _jobs.get(job_id)
        if not job or job["status"] != "completed":
            return None
        path = Path(job["output"])
        if not path.is_file():
            return None
        return path, job["download_name"], int(job["size"])


def cleanup_job(job_id: str) -> None:
    with _lock:
        job = _jobs.pop(job_id, None)
    if job:
        shutil.rmtree(job.get("workdir", ""), ignore_errors=True)


def cleanup_expired() -> int:
    now = time.time()
    with _lock:
        expired = [
            key for key, job in _jobs.items()
            if job["status"] in {"completed", "failed"} and now - job["updated_at"] > TTL_SECONDS
        ]
    for key in expired:
        cleanup_job(key)
    return len(expired) + cleanup_orphan_directories(now)


def cleanup_orphan_directories(now: float | None = None) -> int:
    """Remove abandoned workdirs without a live in-memory job after the temp TTL."""
    current_time = time.time() if now is None else now
    try:
        JOB_ROOT.mkdir(parents=True, exist_ok=True)
        entries = list(JOB_ROOT.iterdir())
    except OSError:
        logger.warning("Could not scan temporary job root %s", JOB_ROOT, exc_info=True)
        return 0
    with _lock:
        active_workdirs = {Path(str(job.get("workdir", ""))) for job in _jobs.values()}
    removed = 0
    for path in entries:
        if not path.is_dir() or path in active_workdirs:
            continue
        try:
            manifest = path / "job.json"
            last_change = max(path.stat().st_mtime, manifest.stat().st_mtime) if manifest.exists() else path.stat().st_mtime
            if current_time - last_change > TTL_SECONDS:
                shutil.rmtree(path, ignore_errors=True)
                removed += 1
        except FileNotFoundError:
            continue
        except OSError:
            logger.warning("Could not inspect temporary job directory %s", path, exc_info=True)
    return removed


def _public(job: dict[str, Any]) -> dict[str, Any]:
    keys = ("id", "status", "progress", "stage", "download_name", "pages", "size", "warnings", "error", "kind", "page_policy", "quality")
    return {key: job.get(key) for key in keys}

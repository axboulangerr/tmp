from __future__ import annotations

import os
import tempfile
from collections import defaultdict, deque
from pathlib import Path
from threading import Lock
from time import monotonic

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from fastapi.staticfiles import StaticFiles

from drafter_to_excel import MAX_DRAFTS, create_workbook


class ExportRequest(BaseModel):
    urls: list[str] = Field(min_length=1, max_length=MAX_DRAFTS)


app = FastAPI(title="Drafter Excel API", version="1.0.0")
RATE_WINDOW_SECONDS = 60.0
MAX_EXPORTS_PER_WINDOW = 3
request_times: dict[str, deque[float]] = defaultdict(deque)
rate_limit_lock = Lock()
export_lock = Lock()


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


def remove_file(path: Path) -> None:
    path.unlink(missing_ok=True)


def enforce_rate_limit(client_ip: str) -> None:
    now = monotonic()
    with rate_limit_lock:
        timestamps = request_times[client_ip]
        while timestamps and now - timestamps[0] >= RATE_WINDOW_SECONDS:
            timestamps.popleft()
        if len(timestamps) >= MAX_EXPORTS_PER_WINDOW:
            retry_after = max(1, int(RATE_WINDOW_SECONDS - (now - timestamps[0])))
            raise HTTPException(
                status_code=429,
                detail="Limite atteinte : maximum 3 exports par minute.",
                headers={"Retry-After": str(retry_after)},
            )
        timestamps.append(now)


@app.post("/api/export")
def export_drafts(export_request: ExportRequest, request: Request) -> FileResponse:
    client_ip = request.client.host if request.client else "unknown"
    enforce_rate_limit(client_ip)
    if not export_lock.acquire(blocking=False):
        raise HTTPException(status_code=429, detail="Un export est déjà en cours. Réessayez plus tard.")

    output_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix="drafter-", suffix=".xlsx", delete=False) as temp_file:
            output_path = Path(temp_file.name)
        create_workbook(export_request.urls, output_path)
    except ValueError as error:
        if output_path is not None:
            remove_file(output_path)
        raise HTTPException(status_code=422, detail=str(error)) from error
    except RuntimeError as error:
        if output_path is not None:
            remove_file(output_path)
        raise HTTPException(status_code=502, detail=str(error)) from error
    except Exception as error:
        if output_path is not None:
            remove_file(output_path)
        raise HTTPException(status_code=500, detail="La génération du classeur a échoué.") from error
    finally:
        export_lock.release()

    return FileResponse(
        output_path,
        filename="drafter_drafts.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=BackgroundTask(remove_file, output_path),
    )


frontend_dist = Path(__file__).resolve().parent / "web" / "dist"
if frontend_dist.is_dir():
    app.mount("/", StaticFiles(directory=frontend_dist, html=True), name="frontend")
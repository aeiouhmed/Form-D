"""FastAPI application providing health and file conversion endpoints."""

# Run with: uvicorn app.main:app --host 0.0.0.0 --port 8000

from datetime import datetime, timezone
from io import BytesIO
import logging
import traceback

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from .convert import convert_to_k1

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # replace with trusted origins for production deployments
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
async def health() -> dict[str, str]:
    """Simple health check endpoint."""
    return {"status": "ok"}


@app.post("/api/convert")
async def convert(
    file: UploadFile = File(...),
    uom_mode: str = Query("random"),
) -> StreamingResponse:
    """Convert uploaded spreadsheet into K1 import XLSX."""
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    uom_mode = (uom_mode or "random").lower()
    if uom_mode not in ("kgm", "random"):
        raise HTTPException(
            status_code=400, detail="uom_mode must be 'kgm' or 'random'"
        )

    try:
        converted_bytes = convert_to_k1(data, uom_mode=uom_mode)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(
            status_code=400,
            detail={"error": "VALIDATION", "detail": str(exc)},
        ) from exc
    except Exception as exc:  # noqa: BLE001
        logging.error("Unexpected error during conversion: %s", exc)
        logging.error(traceback.format_exc())
        raise HTTPException(
            status_code=500, detail="Conversion failed due to an unexpected error."
        ) from exc

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    filename = f'k1-import-{timestamp}.xlsx'
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
    }
    return StreamingResponse(
        BytesIO(converted_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


app.mount("/", StaticFiles(directory="../web", html=True), name="web")

"""
FastAPI application providing health and file conversion endpoints.

This script sets up a FastAPI server with Cross-Origin Resource Sharing (CORS) enabled,
allowing it to accept requests from any origin. It provides a health check endpoint
and a file conversion endpoint that processes uploaded spreadsheets.
"""

# To run the application, use the following command in the terminal:
# uvicorn app.main:app --host 0.0.0.0 --port 8000

from datetime import datetime, timezone
from io import BytesIO
import logging
import traceback

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from .convert import convert_to_k1

# Initialize the FastAPI application.
app = FastAPI()

# Add CORS middleware to allow cross-origin requests.
# This is configured to be open for development purposes.
# For production, 'allow_origins' should be restricted to trusted domains.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # !!!!! replace with trusted origins for production deployments
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
async def health() -> dict[str, str]:
    """
    Simple health check endpoint.

    Returns:
        A dictionary with the status "ok" to indicate the server is running.
    """
    return {"status": "ok"}


@app.post("/api/convert")
async def convert(
    file: UploadFile = File(...),
    country: str = Query("ID"),
) -> StreamingResponse:
    """
    Convert an uploaded spreadsheet into the K1 import XLSX format.

    Args:
        file: The uploaded spreadsheet file.
        country: The country code, defaults to "ID".

    Returns:
        A streaming response containing the converted XLSX file.

    Raises:
        HTTPException: If the file is empty, the country is missing,
                       or an error occurs during conversion.
    """
    # Read the uploaded file's content.
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    # Sanitize the country parameter.
    country = (country or "ID").strip()
    if not country:
        raise HTTPException(status_code=400, detail="country must be provided.")

    try:
        # Perform the conversion.
        converted_bytes = convert_to_k1(data, country=country)
    except FileNotFoundError as exc:
        # Handle cases where a required template file is not found.
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except ValueError as exc:
        # Handle validation errors during conversion.
        raise HTTPException(
            status_code=400,
            detail={"error": "VALIDATION", "detail": str(exc)},
        ) from exc
    except Exception as exc:  # noqa: BLE001
        # Catch any other unexpected errors.
        logging.error("Unexpected error during conversion: %s", exc)
        logging.error(traceback.format_exc())
        raise HTTPException(
            status_code=500, detail="Conversion failed due to an unexpected error."
        ) from exc

    # Generate a timestamped filename for the converted file.
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    filename = f'k1-import-{timestamp}.xlsx'

    # Set headers to prompt the browser to download the file.
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
    }

    # Return the converted file as a streaming response.
    return StreamingResponse(
        BytesIO(converted_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


# Mount the static files directory to serve the frontend application.
# This allows the server to deliver the HTML, CSS, and JavaScript files.
app.mount("/", StaticFiles(directory="../web", html=True), name="web")

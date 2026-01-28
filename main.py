import shutil
import logging
import tempfile
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
import uvicorn

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), ".."))

# Import main fuctions
from src.processor import process_document
from src.logger import setup_logging

# Setup Logging
setup_logging(level=logging.INFO)
logger = logging.getLogger("API")

app = FastAPI(title="Docx Formatter API")

def cleanup_files(paths: list[Path]):
    """
    Background task to remove temporary files after the response is sent.
    """
    for path in paths:
        try:
            if path.exists():
                path.unlink()
                logger.info("Cleaned up temp file: %s", path)
        except Exception as e:
            logger.exception("Failed to delete temp file %s: %s", path, e)

@app.post("/process-document/")
async def api_process_document(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """
    Endpoint to upload a docx, process it, and download the result.
    """
    # File Type Validation
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Invalid file type. Only .docx is supported.")

    # Create temporary directories/files
    # This is necessary to store the input and be able to access it
    temp_dir = Path(tempfile.gettempdir())
    
    input_tmp = temp_dir / f"upload_{file.filename}"
    output_tmp = temp_dir / f"processed_{file.filename}"

    try:
        # Save uploaded file to disk
        logger.info("Receiving file: %s", file.filename)
        with input_tmp.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Run the main processor function
        process_document(input_tmp, output_tmp)

        # Schedule the Cleanup
        # BackgroundTasks runs AFTER the response is sent.
        background_tasks.add_task(cleanup_files, [input_tmp, output_tmp])

        # Return the file
        return FileResponse(
            path=output_tmp,
            filename=f"processed_{file.filename}",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        # Clean up immediately if something failed before response
        cleanup_files([input_tmp, output_tmp])
        logger.error(f"API Error: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
    

if __name__ == "__main__":
    # reload = True, assuming it's a development version
    logger.info("Starting Uvicorn server...")
    uvicorn.run(
        "main:app", 
        host="127.0.0.1", 
        port=8000, 
        reload=True
    )
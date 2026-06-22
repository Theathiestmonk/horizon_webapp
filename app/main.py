"""
main.py — v5.1
Main FastAPI application for Horizon Export Document Generator.
"""

import logging
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)

from app.routers.documents import router as documents_router

app = FastAPI(
    title="Horizon Export Document Generator",
    version="5.1",
    description="Generates all 8 export documents from Google Sheets",
    docs_url="/api/docs",
    redoc_url="/api/redoc"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(documents_router, prefix="/api/v1/invoices", tags=["Documents"])

# Serve generated .docx files at /output/<filename> so GAS can download them
_output_dir = Path("output")
_output_dir.mkdir(parents=True, exist_ok=True)
app.mount("/output", StaticFiles(directory=str(_output_dir)), name="output")

@app.get("/health")
async def health():
    return {"status": "healthy", "version": "5.1", "service": "Horizon Export Document Generator"}

@app.get("/")
async def root():
    return {
        "message": "Horizon Export Document Generator v5.1",
        "docs": "/api/docs",
        "health": "/health"
    }

@app.on_event("startup")
async def startup_event():
    logger.info("Horizon v5.1 starting up...")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8080, workers=1, reload=False)

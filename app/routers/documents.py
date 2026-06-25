import logging
import os
from fastapi import APIRouter, HTTPException, Query, Body
from typing import List, Optional
from app.models.payload import HorizonPayload, validate_payload
from app.services.document_generator import generate_document

logger = logging.getLogger(__name__)
router = APIRouter()

BASE_DOWNLOAD_URL = os.getenv('BASE_DOWNLOAD_URL', 'http://localhost:8080/output')

DOCUMENT_MAP = {
    "proforma_invoice": "PI FORMAT.docx",
    "commercial_invoice": "Commercial_Invoice.docx",
    "packing_list": "Packing_List.docx",
    "annexure_1": "Annexure_1.docx",
    "tax_invoice": "Tax_Invoice.docx",
    "annexure_c": "Annexure_C.docx",
    "scomet": "SCOMET_Declaration.docx",
    "dbk": "DBK_Declaration.docx",
    "vintage": "Vintage_Car_Declaration.docx",
}


@router.post("/generate")
async def generate_documents(
    invoice_no: str = Query(..., description="Invoice number"),
    documents: Optional[List[str]] = Query(None, description="Documents to generate"),
    payload: HorizonPayload = Body(...),
):
    """
    Accept the full HorizonPayload from Google Apps Script and generate all documents.
    Returns generated_files list with download URLs so GAS showDownloadDialog works.
    """
    logger.info(f"[{invoice_no}] generate_documents called | mode={payload.mode} | vehicles={len(payload.vehicles)}")

    issues = validate_payload(payload.model_dump())
    errors = [i.message for i in issues if i.severity == "error"]
    if errors:
        raise HTTPException(status_code=400, detail={"errors": errors})

    docs_to_generate = documents if documents else list(DOCUMENT_MAP.keys())
    generated_files = []
    failed = []

    for doc_key in docs_to_generate:
        if doc_key not in DOCUMENT_MAP:
            failed.append({"document": doc_key, "error": "Unknown document type"})
            continue

        template_name = DOCUMENT_MAP[doc_key]
        result = generate_document(
            template_name=template_name,
            payload=payload,
            invoice_no=invoice_no,
        )

        if result.get("status") == "success":
            download_name = result["download_name"]
            file_url = result.get("download_url") or f"{BASE_DOWNLOAD_URL}/{download_name}"
            generated_files.append({
                "template": doc_key,
                "document": template_name,
                "download_name": download_name,
                # gcs_url key expected by GAS showDownloadDialog
                "gcs_url": file_url,
                "output_file": result.get("output_file", ""),
            })
        else:
            failed.append({"document": doc_key, "template": template_name, "error": result.get("error", "Unknown error")})

    status = "success" if not failed else ("partial" if generated_files else "failed")
    logger.info(f"[{invoice_no}] Done | generated={len(generated_files)} | failed={len(failed)}")

    return {
        "status": status,
        "invoice_no": invoice_no,
        "generated_files": generated_files,
        "failed": failed,
    }


@router.post("/generate/pi")
async def generate_pi_document(
    invoice_no: str = Query(..., description="Invoice number"),
    payload: HorizonPayload = Body(...),
):
    """Generate PI FORMAT (Proforma Invoice) only — called by GAS generatePIDocument()."""
    logger.info(f"[{invoice_no}] PI-only generation | vehicles={len(payload.vehicles)} | items={len(payload.items)}")

    for i, item in enumerate(payload.items):
        d = item.model_dump()
        logger.info(f"[{invoice_no}] item[{i}]: hsn={d.get('hsn_code')} qty={d.get('quantity')} desc_pi='{d.get('description_pi')}' desc='{d.get('description')}'")
    for i, v in enumerate(payload.vehicles):
        d = v.model_dump()
        logger.info(f"[{invoice_no}] vehicle[{i}]: chassis={d.get('chassis_no')} model='{d.get('model')}' price={d.get('unit_price_usd')}")

    result = generate_document(
        template_name="PI FORMAT.docx",
        payload=payload,
        invoice_no=invoice_no,
    )

    if result.get("status") != "success":
        raise HTTPException(status_code=500, detail={"error": result.get("error", "Generation failed"), "template": "PI FORMAT.docx"})

    download_name = result["download_name"]
    file_url = result.get("download_url") or f"{BASE_DOWNLOAD_URL}/{download_name}"

    return {
        "status": "success",
        "invoice_no": invoice_no,
        "generated_files": [{
            "template": "proforma_invoice",
            "document": "PI FORMAT.docx",
            "download_name": download_name,
            "gcs_url": file_url,
            "output_file": result.get("output_file", ""),
        }],
        "failed": [],
    }


@router.post("/{invoice_no}/cha-package")
async def cha_package(invoice_no: str):
    """
    Bundle all previously generated documents for a given invoice into a CHA package.
    Re-generates all documents if they don't exist, then returns download links.
    GAS exportCHAPackage() POSTs here and expects {drive_folder_url} or similar.
    """
    from pathlib import Path
    output_dir = Path("output")

    logger.info(f"[{invoice_no}] CHA package requested")

    # Collect any already-generated files for this invoice
    import re
    safe_inv = invoice_no.replace("/", "-").replace("\\", "-").replace(" ", "_")
    existing = sorted(output_dir.glob(f"{safe_inv}_*.docx"))

    if not existing:
        raise HTTPException(
            status_code=404,
            detail={
                "error": f"No documents found for invoice {invoice_no}. "
                         "Call /generate first, then /cha-package."
            }
        )

    package_files = [
        {
            "template": f.stem.replace(f"{safe_inv}_", ""),
            "document": f.name,
            "gcs_url": f"{BASE_DOWNLOAD_URL}/{f.name}",
        }
        for f in existing
    ]

    logger.info(f"[{invoice_no}] CHA package: {len(package_files)} files")

    return {
        "status": "success",
        "invoice_no": invoice_no,
        "package_files": package_files,
        "drive_folder_url": f"{BASE_DOWNLOAD_URL}/?invoice={invoice_no}",
        "folder_url": f"{BASE_DOWNLOAD_URL}/?invoice={invoice_no}",
        "generated_files": package_files,
    }


@router.get("/documents")
async def list_documents():
    return {"available": list(DOCUMENT_MAP.keys())}


@router.get("/health")
async def health():
    return {"status": "healthy", "version": "5.1"}

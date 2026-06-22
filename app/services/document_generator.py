import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict
from docxtpl import DocxTemplate
from num2words import num2words
from app.services.drive_uploader import upload_file, DRIVE_FOLDER_ID

logger = logging.getLogger(__name__)

TEMPLATE_DIR = Path("templates")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def sanitize_filename(name: str) -> str:
    return name.replace("/", "-").replace("\\", "-").replace(" ", "_")

def convert_to_words(amount: float, currency: str = "INR") -> str:
    try:
        major = int(amount)
        minor = int(round((amount - major) * 100))
        if currency == "INR":
            major_words = num2words(major, lang='en_IN').title()
            currency_name = "Rupees"
        else:
            major_words = num2words(major, lang='en').title()
            currency_name = "US Dollars"
        word_str = f"{major_words} {currency_name}"
        if minor > 0:
            minor_words = num2words(minor, lang='en_IN' if currency == "INR" else 'en').title()
            word_str += f" And {minor_words} {'Paise' if currency == 'INR' else 'Cents'}"
        return word_str + " Only"
    except Exception as e:
        logger.error(f"Words conversion error: {e}")
        return f"{amount} {currency}"

def build_context(payload: Any) -> Dict[str, Any]:
    if hasattr(payload, 'model_dump'):
        context = payload.model_dump()
    else:
        context = dict(payload) if isinstance(payload, dict) else payload

    # Flatten financials
    if isinstance(context.get('financials'), dict):
        financials = context.pop('financials')
        context.update({
            'quantity': financials.get('quantity'),
            'unit_price_usd': financials.get('unit_price_usd'),
            'fob_total_usd': financials.get('fob_total_usd'),
            'freight_usd': financials.get('freight_usd'),
            'insurance_usd': financials.get('insurance_usd'),
            'cif_total_usd': financials.get('cif_total_usd'),
            'exchange_rate': financials.get('exchange_rate'),
            'igst_rate': financials.get('igst_rate'),
            'taxable_value_inr': financials.get('taxable_value_inr'),
            'igst_amount_inr': financials.get('igst_amount_inr'),
            'total_value_inr': financials.get('total_value_inr'),
            'total_inr': financials.get('total_value_inr')
        })

    # Container No (from nested shipping)
    if isinstance(context.get('shipping'), dict):
        shipping = context.get('shipping', {})
        context['container_no'] = shipping.get('container_no', '')
    else:
        context['container_no'] = context.get('container_no', '')

    # Amount in Words fallback
    if not context.get('amount_usd_words'):
        cif = context.get('cif_total_usd') or 0
        context['amount_usd_words'] = convert_to_words(cif, currency="USD")

    if not context.get('amount_inr_words'):
        total_inr = context.get('total_value_inr') or context.get('total_inr') or 0
        context['amount_inr_words'] = convert_to_words(total_inr, currency="INR")

    context['generation_timestamp'] = datetime.utcnow().isoformat()
    
    logger.info(f"Context built with {len(context)} keys for invoice {context.get('invoice_no')}")
    return context

def generate_document(template_name: str, payload: Any, invoice_no: str) -> Dict[str, Any]:
    safe_invoice = sanitize_filename(invoice_no)
    logger.info(f"[{invoice_no}] Generating {template_name}")
    
    try:
        template_path = TEMPLATE_DIR / template_name
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        context = build_context(payload)
        
        logger.info(f"Rendering template {template_name} with {len(context)} context keys")
        logger.debug(f"Context keys: {list(context.keys())}")   # Helpful for debugging
        
        doc = DocxTemplate(str(template_path))
        doc.render(context) 
        
        output_filename = f"{safe_invoice}_{template_name}"
        output_path = OUTPUT_DIR / output_filename
        doc.save(str(output_path))

        # Upload to Google Drive and return the Drive URL
        drive_url = ''
        try:
            drive_result = upload_file(
                local_path=str(output_path),
                filename=output_filename,
                folder_id=DRIVE_FOLDER_ID,
            )
            drive_url = drive_result['view_url']
        except Exception as upload_err:
            logger.warning(f"[{invoice_no}] Drive upload failed for {template_name}: {upload_err}")

        logger.info(f"[{invoice_no}] {template_name} generated successfully")
        return {
            "status": "success",
            "template": template_name,
            "output_file": str(output_path),
            "download_name": output_filename,
            "drive_url": drive_url,
        }
  
    except Exception as e:
        logger.error(f"[{invoice_no}] {template_name} FAILED: {str(e)}", exc_info=True)
        return {
            "status": "error",
            "template": template_name,
            "error": str(e)
        }

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict
from docxtpl import DocxTemplate
from num2words import num2words

logger = logging.getLogger(__name__)

TEMPLATE_DIR = Path("templates")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Maps each DOCX template to the description field it should use from Item
TEMPLATE_DESC_FIELD = {
    "PI FORMAT.docx":              "description_pi",
    "Commercial_Invoice.docx":     "description_commercial",
    "Packing_List.docx":           "description_packing",
    "Tax_Invoice.docx":            "description_tax",
    "SCOMET_Declaration.docx":     "description_scomet",
    "Annexure_C.docx":             "description_scomet",
    "Annexure_1.docx":             "description_annexure1",
}

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

def build_context(payload: Any, template_name: str = "") -> Dict[str, Any]:
    if hasattr(payload, 'model_dump'):
        context = payload.model_dump()
    else:
        context = dict(payload) if isinstance(payload, dict) else payload

    # ── Flatten financials ──────────────────────────────────────────────────
    if isinstance(context.get('financials'), dict):
        financials = context.pop('financials')
        context.update({
            'quantity':          financials.get('quantity', 0),
            'unit_price_usd':    financials.get('unit_price_usd', 0),
            'fob_total_usd':     financials.get('fob_total_usd', 0),
            'freight_usd':       financials.get('freight_usd', 0),
            'insurance_usd':     financials.get('insurance_usd', 0),
            'cif_total_usd':     financials.get('cif_total_usd', 0),
            'exchange_rate':     financials.get('exchange_rate', 0),
            'igst_rate':         financials.get('igst_rate', 0),
            'taxable_value_inr': financials.get('taxable_value_inr', 0),
            'igst_amount_inr':   financials.get('igst_amount_inr', 0),
            'total_value_inr':   financials.get('total_value_inr', 0),
            'total_inr':         financials.get('total_value_inr', 0),
        })

    # ── Flatten weights ─────────────────────────────────────────────────────
    if isinstance(context.get('weights'), dict):
        weights = context.get('weights', {})
        context['total_packages']     = weights.get('total_packages', 0)
        context['total_net_weight']   = weights.get('net_weight_kg', 0)
        context['total_gross_weight'] = weights.get('gross_weight_kg', 0)
        context['net_weight_kg']      = weights.get('net_weight_kg', 0)
        context['gross_weight_kg']    = weights.get('gross_weight_kg', 0)

    # ── Flatten shipping (kept as nested dict too for PI FORMAT shipping.* refs) ──
    if isinstance(context.get('shipping'), dict):
        shipping = context.get('shipping', {})
        context['container_no']           = shipping.get('container_no', '')
        context['pre_carriage_by']        = shipping.get('pre_carriage_by', '')
        context['mode_of_transport']      = shipping.get('mode_of_transport', '')
        context['country_of_origin']      = shipping.get('country_of_origin', 'INDIA')
        context['port_of_loading']        = shipping.get('port_of_loading', '')
        context['port_of_discharge']      = shipping.get('port_of_discharge', '')
        context['final_destination']      = shipping.get('final_destination', '')
        context['country_of_destination'] = shipping.get('country_of_destination', '')
        context['place_of_receipt']  = context.get('place_of_receipt')  or shipping.get('port_of_loading', '')
        context['delivery_terms']    = context.get('delivery_terms')    or shipping.get('mode_of_transport', '')
    else:
        context['container_no'] = context.get('container_no', '')

    # ── Cross-template name aliases ─────────────────────────────────────────
    context['total_cif_usd']     = context.get('cif_total_usd') or 0
    context['total_fob_usd']     = context.get('total_fob_usd') or context.get('fob_total_usd') or 0
    context['total_taxable_inr'] = context.get('taxable_value_inr') or 0
    context['igst_percent']      = context.get('igst_rate') or 0

    # ── Swap item.description per template ──────────────────────────────────
    desc_field = TEMPLATE_DESC_FIELD.get(template_name, "")
    if desc_field and isinstance(context.get('items'), list):
        for it in context['items']:
            if isinstance(it, dict) and it.get(desc_field):
                it['description'] = it[desc_field]

    # ── Add aliases + sr_no + unit + package range on each item ────────────
    _running_item = 0
    for idx, it in enumerate(context.get('items') or []):
        if isinstance(it, dict):
            it['sr_no']   = idx + 1
            it.setdefault('unit',       'Nos')
            it.setdefault('unit_price', it.get('rate_per_unit', 0))
            it.setdefault('total',      it.get('amount_usd', 0))
            _qty = it.get('quantity', 0)
            it['sr_start'] = _running_item + 1
            it['sr_end']   = _running_item + _qty
            _running_item += _qty

    # ── total_quantity across all items ─────────────────────────────────────
    context['total_quantity'] = sum(
        it.get('quantity', 0) for it in (context.get('items') or [])
        if isinstance(it, dict)
    )
    # PI FORMAT uses top-level {{ quantity }} — override with actual vehicle count
    if context['total_quantity'] > 0:
        context['quantity'] = context['total_quantity']

    # ── item singular for PI FORMAT — aggregate all items into one summary row
    if isinstance(context.get('items'), list) and context['items']:
        first = dict(context['items'][0])
        first['quantity']   = context['total_quantity']
        first['amount_usd'] = sum(it.get('amount_usd', 0) for it in context['items'] if isinstance(it, dict))
        first['total']      = first['amount_usd']
        context['item'] = first

    # ── vin_list for Annexure_1 ({% for v in vin_list %}) ──────────────────
    if isinstance(context.get('vehicles'), list):
        context['vin_list'] = [
            {
                'sr_no':      i + 1,
                'chassis_no': v.get('chassis_no', ''),
                'engine_no':  v.get('engine_no', ''),
                'model':      v.get('model', ''),
                'color':      v.get('color', ''),
            }
            for i, v in enumerate(context['vehicles'])
            if isinstance(v, dict)
        ]
        # vehicle_models_list for Annexure C (item 15 "Vehicles" section)
        context['vehicle_models_list'] = ', '.join(
            v.get('model', '') for v in context['vehicles']
            if isinstance(v, dict) and v.get('model')
        )

    # ── notify_1 fallback → buyer name + address for Annexure C consignee ──
    if not context.get('notify_1'):
        buyer = context.get('buyer') or {}
        if isinstance(buyer, dict):
            parts = [buyer.get('name', ''), buyer.get('address', ''), buyer.get('country', '')]
            context['notify_1'] = ', '.join(p for p in parts if p)

    # ── packages for Packing_List ({% for p in packages %}) ────────────────
    if isinstance(context.get('items'), list):
        marks = context.get('marks_and_numbers', '')
        _running_pkg = 0
        _pkg_list = []
        for it in context['items']:
            if isinstance(it, dict):
                _qty = it.get('quantity', 0)
                _pkg_list.append({
                    'marks':       marks,
                    'description': it.get('description', ''),
                    'hsn_code':    it.get('hsn_code', ''),
                    'quantity':    _qty,
                    'unit':        it.get('unit', 'Nos'),
                    'sr_start':    _running_pkg + 1,
                    'sr_end':      _running_pkg + _qty,
                })
                _running_pkg += _qty
        context['packages'] = _pkg_list

    # ── generation_date fallback ────────────────────────────────────────────
    if not context.get('generation_date'):
        context['generation_date'] = datetime.utcnow().strftime('%d.%m.%Y')

    # ── scomet_product_desc fallback ────────────────────────────────────────
    if not context.get('scomet_product_desc') and isinstance(context.get('items'), list) and context['items']:
        context['scomet_product_desc'] = context['items'][0].get('description_scomet', '') or context['items'][0].get('description', '')

    # ── insurance_ref_no fallback ───────────────────────────────────────────
    if not context.get('insurance_ref_no'):
        context['insurance_ref_no'] = context.get('lc_number', '')

    # ── Amount in words ─────────────────────────────────────────────────────
    if not context.get('amount_usd_words'):
        context['amount_usd_words'] = convert_to_words(context.get('cif_total_usd') or 0, currency="USD")
    if not context.get('amount_inr_words'):
        context['amount_inr_words'] = convert_to_words(
            context.get('total_value_inr') or context.get('total_inr') or 0, currency="INR"
        )

    context['generation_timestamp'] = datetime.utcnow().isoformat()
    logger.info(f"[{template_name}] Context built: {len(context)} keys | invoice={context.get('invoice_no')} | items={len(context.get('items') or [])} | vehicles={len(context.get('vehicles') or [])}")
    return context

def generate_document(template_name: str, payload: Any, invoice_no: str) -> Dict[str, Any]:
    safe_invoice = sanitize_filename(invoice_no)
    logger.info(f"[{invoice_no}] Generating {template_name}")
    
    try:
        template_path = TEMPLATE_DIR / template_name
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        context = build_context(payload, template_name=template_name)
        
        logger.info(f"Rendering template {template_name} with {len(context)} context keys")
        logger.debug(f"Context keys: {list(context.keys())}")   # Helpful for debugging
        
        doc = DocxTemplate(str(template_path))
        doc.render(context) 
        
        output_filename = f"{safe_invoice}_{template_name}"
        output_path = OUTPUT_DIR / output_filename
        doc.save(str(output_path))

        logger.info(f"[{invoice_no}] {template_name} generated locally at {output_path}")
        return {
            "status": "success",
            "template": template_name,
            "output_file": str(output_path),
            "download_name": output_filename,
            "download_url": f"/output/{output_filename}",
        }
  
    except Exception as e:
        logger.error(f"[{invoice_no}] {template_name} FAILED: {str(e)}", exc_info=True)
        return {
            "status": "error",
            "template": template_name,
            "error": str(e)
        }

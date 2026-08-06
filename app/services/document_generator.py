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
BASE_DOWNLOAD_URL = os.getenv("BASE_DOWNLOAD_URL", "http://localhost:8080/output").rstrip("/")

# Maps each DOCX template to the description field it should use from Item
TEMPLATE_DESC_FIELD = {
    "PI FORMAT.docx":              "description_pi",
    "Commercial_Invoice.docx":     "description_commercial",
    "Packing_List.docx":           "description_packing",
    "Tax_Invoice.docx":            "description_tax",
    "SCOMET_Declaration.docx":     "description_scomet",
    "Annexure_C.docx":             "description_scomet",
    "Annexure_1.docx":             "description_annexure1",
    "CHA TI.docx":                 "description_cha_ti",
    "CHA PL.docx":                 "description_cha_pl",
    "CHA CI.docx":                 "description_cha_ci",
}

# Maps each DOCX template to which per-item HSN field it should print.
# Customer-facing documents (CI, PI, Tax Invoice, Packing List) use the
# customer's own country's HSN code; everything else — SCOMET/Annexure/DBK/
# Vintage and every CHA document — uses the Indian HSN code, since those are
# Indian customs/export filings. See the HSN resolution block in
# buildPayload() (script.gs).
TEMPLATE_HSN_FIELD = {
    "PI FORMAT.docx":              "hsn_code_pi",
    "Commercial_Invoice.docx":     "hsn_code_user_country",
    "Tax_Invoice.docx":            "hsn_code_user_country",
    "Packing_List.docx":           "hsn_code_user_country",
    "SCOMET_Declaration.docx":     "hsn_code_india",
    "Annexure_C.docx":             "hsn_code_india",
    "Annexure_1.docx":             "hsn_code_india",
    "DBK_Declaration.docx":        "hsn_code_india",
    "Vintage_Car_Declaration.docx":"hsn_code_india",
    "CHA TI.docx":                 "hsn_code_india",
    "CHA PL.docx":                 "hsn_code_india",
    "CHA CI.docx":                 "hsn_code_india",
}

def sanitize_filename(name: str) -> str:
    return name.replace("/", "-").replace("\\", "-").replace(" ", "_")

# Monetary + rate fields — exchange_rate keeps its decimals (genuine currency
# conversion precision), everything else here is always a whole number in
# practice (GST slabs are 0/5/12/18/28) so it gets rounded to an int so
# templates never print "1500.0" or "18.0%".
_MONEY_KEYS_TOP = (
    'unit_price_usd', 'fob_total_usd', 'freight_usd', 'insurance_usd', 'cif_total_usd',
    'taxable_value_inr', 'igst_amount_inr', 'total_value_inr', 'total_inr',
    'total_cif_usd', 'total_fob_usd', 'total_taxable_inr',
    'igst_rate', 'igst_percent',
)
_MONEY_KEYS_ITEM = ('rate_per_unit', 'amount_usd', 'unit_price', 'total')
_MONEY_KEYS_VEHICLE = ('unit_price_usd',)

# Genuine rates/measurements — real decimals (e.g. 83.45 exchange rate, 600.5
# kg) must survive untouched, but a value that happens to be a whole number
# (e.g. 94.0, 3600.0) should still print as "94" / "3600", not "94.0" /
# "3600.0". So: strip the ".0" only when there's no actual fraction.
_RATE_KEYS_TOP = (
    'exchange_rate', 'total_net_weight', 'total_gross_weight', 'net_weight_kg', 'gross_weight_kg'
)

def _clean_rate(v):
    return int(v) if isinstance(v, float) and v == int(v) else v

def _round_money(context: Dict[str, Any]) -> None:
    for key in _MONEY_KEYS_TOP:
        if isinstance(context.get(key), float):
            context[key] = int(round(context[key]))
    for key in _RATE_KEYS_TOP:
        if key in context:
            context[key] = _clean_rate(context[key])
    for it in (context.get('items') or []):
        if isinstance(it, dict):
            for key in _MONEY_KEYS_ITEM:
                if isinstance(it.get(key), float):
                    it[key] = int(round(it[key]))
    item_singular = context.get('item')
    if isinstance(item_singular, dict):
        for key in _MONEY_KEYS_ITEM:
            if isinstance(item_singular.get(key), float):
                item_singular[key] = int(round(item_singular[key]))
    for v in (context.get('vehicles') or []):
        if isinstance(v, dict):
            for key in _MONEY_KEYS_VEHICLE:
                if isinstance(v.get(key), float):
                    v[key] = int(round(v[key]))

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
    # Always assign (not "only if truthy") — an intentionally blank desc_field
    # (Invoice_Descriptions collapsing every item but the first down to '' so
    # a concatenating template prints the text exactly once) must make
    # `description` blank too, not silently fall back to whatever value it
    # already held from an earlier per-template pass.
    desc_field = TEMPLATE_DESC_FIELD.get(template_name, "")
    if desc_field and isinstance(context.get('items'), list):
        for it in context['items']:
            if isinstance(it, dict):
                # Always set 'description', even if the source field is missing (use fallback to description_pi)
                if desc_field in it:
                    it['description'] = it[desc_field]
                elif 'description_pi' in it:
                    # Fallback to description_pi if the template-specific field doesn't exist
                    it['description'] = it['description_pi']
                else:
                    # Last resort: use empty string
                    it['description'] = ''

    # ── Swap item.hsn_code per template ──────────────────────────────────────
    # hsn_code_user_country feeds Commercial_Invoice/Tax_Invoice/Packing_List,
    # hsn_code_india feeds everything else (SCOMET/Annexure/CHA), hsn_code_pi
    # feeds PI FORMAT. Each of those 3 fields already falls back to the
    # existing model-level HSN in buildPayload() (script.gs) when a Stock tab
    # override is blank, so this only ever changes output once an override
    # is actually filled in. Runs before context['item'] (singular, below) is
    # copied from items[0], so that also picks up the already-swapped value.
    hsn_field = TEMPLATE_HSN_FIELD.get(template_name, "")
    if hsn_field and isinstance(context.get('items'), list):
        for it in context['items']:
            if isinstance(it, dict) and it.get(hsn_field):
                it['hsn_code'] = it[hsn_field]

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
            # Ensure all description fields exist (fallback to empty string if missing)
            for desc_key in ['description', 'description_commercial', 'description_scomet',
                            'description_packing', 'description_tax', 'description_pi',
                            'description_annexure1', 'description_cha_ti', 'description_cha_pl', 'description_cha_ci']:
                it.setdefault(desc_key, '')
            # Ensure HSN fields exist
            for hsn_key in ['hsn_code', 'hsn_code_pi', 'hsn_code_india', 'hsn_code_user_country']:
                it.setdefault(hsn_key, '')

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

    # PI FORMAT.docx and Tax_Invoice.docx both use a native docxtpl row-loop
    # ({%tr for item in items %}, referencing item.hsn_code/description/
    # quantity/rate_per_unit/amount_usd — Tax_Invoice also uses item.sr_start/
    # sr_end) — they were already built to print one row per model, same as
    # Commercial Invoice / CHA CI/TI/PL. A previous version of this function
    # collapsed context['items'] into one merged summary row here, which
    # fought the template's own design and silently dropped every model but
    # the first from the description. Removed — context['items'] (with each
    # item's sr_no/sr_start/sr_end already set above) now renders as-is, one
    # row per model, exactly like the other multi-model documents.

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

    # PI FORMAT's amount cell is a bare {{ amount_usd_words }} with no static
    # label (unlike CI/CHA CI/CHA TI, which already print their own "AMOUNT
    # CHARGEABLE..." label before the placeholder) — so PI FORMAT needs the
    # full phrase baked into the value itself.
    if template_name == "PI FORMAT.docx":
        for key in ("amount_usd_words", "amount_inr_words"):
            val = context.get(key) or ""
            if val and not val.upper().startswith("AMOUNT CHARGEABLE"):
                context[key] = f"AMOUNT CHARGEABLE IN {val}"

    # ── Whole-number money — strip ".0" from every USD/INR amount ──────────
    _round_money(context)

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
        logger.info(f"Items in context: {len(context.get('items', []))} | Vehicles: {len(context.get('vehicles', []))}")
        if context.get('items'):
            for idx, item in enumerate(context['items'][:2]):  # Log first 2 items for debugging
                if isinstance(item, dict):
                    logger.info(f"  Item[{idx}]: hsn={item.get('hsn_code')} qty={item.get('quantity')} desc_comm={bool(item.get('description_commercial'))}")

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
            "download_url": f"{BASE_DOWNLOAD_URL}/{output_filename}",
        }
  
    except Exception as e:
        logger.error(f"[{invoice_no}] {template_name} FAILED: {str(e)}", exc_info=True)
        return {
            "status": "error",
            "template": template_name,
            "error": str(e)
        }

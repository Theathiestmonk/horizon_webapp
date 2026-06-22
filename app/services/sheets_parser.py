import json
import logging
import os
from typing import Dict, Any, List
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def get_sheets_service():
    try:
        raw = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
        if not raw:
            raise EnvironmentError('GOOGLE_SERVICE_ACCOUNT_JSON env var is not set')
        info = json.loads(raw)
        credentials = Credentials.from_service_account_info(info, scopes=SCOPES)
        return build('sheets', 'v4', credentials=credentials)
    except Exception as e:
        logger.error(f"Sheets service error: {e}")
        raise

def safe_extract_dicts(sheet_data: List[List[str]]) -> List[Dict[str, Any]]:
    if not sheet_data or len(sheet_data) < 2:
        return []
    headers = [h.strip() if h else "" for h in sheet_data[0]]
    dicts = []
    for row in sheet_data[1:]:
        normalized_row = row + [''] * (len(headers) - len(row))
        dicts.append(dict(zip(headers, normalized_row)))
    return dicts

def clean_float(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).replace('$', '').replace('₹', '').replace(',', '').strip()
        return float(s) if s else 0.0
    except (ValueError, TypeError):
        return 0.0

def clean_int(v: Any) -> int:
    try:
        return int(clean_float(v))
    except (ValueError, TypeError):
        return 0

def get_shipment_from_sheets(spreadsheet_id: str, invoice_no: str, service=None) -> Dict[str, Any]:
    if service is None:
        service = get_sheets_service()
    
    ranges = ['CONTROL!A1:Z100', 'Stock!A1:J2000', 'Company!A1:B30']
    
    try:
        batch_result = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges).execute()
    except Exception as e:
        logger.error(f"Sheets fetch failed: {e}")
        raise
    
    value_ranges = batch_result.get('valueRanges', [])
    control_list = safe_extract_dicts(value_ranges[0].get('values', []))
    stock_list = safe_extract_dicts(value_ranges[1].get('values', []))
    company_dict = {r[0]: r[1] if len(r) > 1 else "" for r in value_ranges[2].get('values', [])[1:] if r}
    
    invoice_row = next((r for r in control_list if r.get('Invoice No') == invoice_no), None)
    if not invoice_row:
        raise ValueError(f"Invoice {invoice_no} not found")
    
    assigned_vehicles = [v for v in stock_list if v.get('Assigned To') == invoice_no]
    
    return {
        "invoice_no": invoice_no,
        "invoice_date": invoice_row.get('Invoice Date', ''),
        "mode": invoice_row.get('Mode', 'FINAL'),
        "exporter": {
            "company_name": company_dict.get('Company Name', ''),
            "address": company_dict.get('Address', ''),
            "iec": company_dict.get('IEC No', ''),
            "gstin": company_dict.get('GSTIN', ''),
        },
        "buyer": {
            "name": invoice_row.get('Buyer Name', ''),
            "address": invoice_row.get('Buyer Address', ''),
            "country": invoice_row.get('Buyer Country', 'Kenya'),
        },
        "shipping": {
            "port_of_discharge": invoice_row.get('Port of Discharge', 'Mombasa'),
            "container_no": invoice_row.get('Container No', ''),
        },
        "financials": {
            "quantity": clean_int(invoice_row.get('Quantity', 0)),
            "unit_price_usd": clean_float(invoice_row.get('Unit Price USD', 0)),
            "cif_total_usd": clean_float(invoice_row.get('CIF Total USD', 0)),
            "exchange_rate": clean_float(invoice_row.get('Exchange Rate', 0)),
        },
        "bank": {
            "bank_name": company_dict.get('Bank Name', ''),
            "swift": company_dict.get('Swift Code', ''),
        },
        "weights": {
            "gross_weight_kg": clean_float(invoice_row.get('Gross Weight KG', 0)),
            "net_weight_kg": clean_float(invoice_row.get('Net Weight KG', 0)),
        },
        "vehicles": [
            {
                "chassis_no": v.get('Chassis No', ''),
                "engine_no": v.get('Engine No', ''),
                "model": v.get('Model', ''),
                "color": v.get('Color', ''),
                "year": v.get('Year', ''),
                "unit_price_usd": clean_float(v.get('Price', 0)),
            }
            for v in assigned_vehicles
        ],
    }

from typing import List, Optional, Any, Dict
from pydantic import BaseModel, field_validator, model_validator, ConfigDict
import logging

logger = logging.getLogger(__name__)

class Item(BaseModel):
    model_config = ConfigDict(extra='allow')
    hsn_code: str = ""
    description: str = ""
    quantity: int = 0
    rate_per_unit: float = 0.0
    amount_usd: float = 0.0

class Vehicle(BaseModel):
    model_config = ConfigDict(extra='allow')
    chassis_no: str = ""
    engine_no: str = ""
    model: str = ""
    color: str = ""
    year: str = ""
    unit_price_usd: float = 0.0
    
    @field_validator('unit_price_usd', mode='before')
    @classmethod
    def clean_price(cls, v):
        if not v or v == "":
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        try:
            cleaned = str(v).replace('$', '').replace(',', '').strip()
            return float(cleaned)
        except ValueError:
            logger.error(f"COERCION_FAILED | value={v}")
            raise ValueError(f"Invalid price: {v}")

class Shipping(BaseModel):
    model_config = ConfigDict(extra='allow')
    pre_carriage_by: str = ""
    mode_of_transport: str = ""
    country_of_origin: str = "INDIA"
    country_of_destination: str = ""
    port_of_loading: str = ""
    port_of_discharge: str = ""
    final_destination: str = ""
    container_no: Optional[str] = ""

class Financials(BaseModel):
    model_config = ConfigDict(extra='allow')
    quantity: int = 0
    unit_price_usd: float = 0.0
    fob_total_usd: float = 0.0
    freight_usd: float = 0.0
    insurance_usd: float = 0.0
    cif_total_usd: float = 0.0
    exchange_rate: float = 0.0
    igst_rate: float = 0.0
    taxable_value_inr: float = 0.0
    igst_amount_inr: float = 0.0
    total_value_inr: float = 0.0

class Weights(BaseModel):
    model_config = ConfigDict(extra='allow')
    net_weight_kg: float = 0.0
    gross_weight_kg: float = 0.0
    total_packages: int = 0

class Bank(BaseModel):
    model_config = ConfigDict(extra='allow')
    bank_name: str = ""
    account_no: str = ""
    swift: str = ""
    branch: str = ""

class Exporter(BaseModel):
    model_config = ConfigDict(extra='allow')
    company_name: str = ""
    address: str = ""
    phone: str = ""
    iec: str = ""
    pan: str = ""
    gstin: str = ""
    signatory: str = ""
    cha: str = ""
    customs_office: str = ""

class Buyer(BaseModel):
    model_config = ConfigDict(extra='allow')
    name: str = ""
    address: str = ""
    country: str = ""

class HorizonPayload(BaseModel):
    model_config = ConfigDict(extra='allow')
    invoice_no: str
    invoice_date: str = ""
    mode: str = "FINAL"
    generation_date: str = ""
    exporter: Exporter
    buyer: Buyer
    shipping: Shipping
    financials: Financials
    bank: Bank
    weights: Weights
    vehicles: List[Vehicle] = []
    items: List[Item] = []
    lc_number: str = ""
    buyers_order_no: str = ""
    notify_1: str = ""
    notify_2: str = ""
    terms_of_payment: str = ""
    company_seal_no: str = ""
    shipping_line_seal_no: str = ""
    marks_and_numbers: str = ""
    scomet_product_desc: str = ""
    amount_usd_words: str = ""
    amount_inr_words: str = ""
    total_fob_usd: float = 0.0
    total_inr: float = 0.0

    @model_validator(mode='after')
    def calculate_totals(self) -> 'HorizonPayload':
        if self.vehicles:
            self.total_fob_usd = sum(v.unit_price_usd for v in self.vehicles)
        elif self.financials and self.financials.fob_total_usd:
            self.total_fob_usd = self.financials.fob_total_usd
        if self.financials and self.financials.total_value_inr:
            self.total_inr = self.financials.total_value_inr
        elif self.financials:
            ex_rate = self.financials.exchange_rate or 0
            try:
                rate = float(str(ex_rate).replace(',', ''))
                self.total_inr = self.financials.cif_total_usd * rate
            except (ValueError, TypeError):
                self.total_inr = 0.0
        return self

class ValidationIssue(BaseModel):
    severity: str = "error"
    field: str = ""
    message: str = ""

def validate_payload(payload_dict: Dict[str, Any]) -> List[ValidationIssue]:
    issues = []
    if not payload_dict.get('invoice_no'):
        issues.append(ValidationIssue(severity="error", field="invoice_no", message="Invoice number required"))
    if not payload_dict.get('invoice_date'):
        issues.append(ValidationIssue(severity="error", field="invoice_date", message="Invoice date required"))
    exporter = payload_dict.get('exporter', {})
    if not exporter.get('company_name'):
        issues.append(ValidationIssue(severity="error", field="exporter.company_name", message="Exporter company name required"))
    buyer = payload_dict.get('buyer', {})
    if not buyer.get('name'):
        issues.append(ValidationIssue(severity="error", field="buyer.name", message="Buyer name required"))
    if payload_dict.get('mode') == 'FINAL':
        if not payload_dict.get('vehicles') or len(payload_dict.get('vehicles', [])) == 0:
            issues.append(ValidationIssue(severity="error", field="vehicles", message="FINAL mode requires vehicles"))
    return issues

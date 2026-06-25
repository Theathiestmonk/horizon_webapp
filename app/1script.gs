const CFG = {
  invoiceNoCell:      'C8',
  dateCell:           'F8',
  customerCell:       'C12',
  buyerAddressCell:   'C13',
  buyerCountryCell:   'H22',
  stockCell:          'C17',
  lcCell:             'C24',
  qtyCell:            'C26',
  unitPriceCell:      'F26',
  containerCell:      'C37',
  portLoadCell:       'C21',
  portDischargeCell:  'F21',
  finalDestinationCell: 'H22',
  modeCell:           'F3',
  exchangeRateCell:   'C29',
  igstRateCell:       'F29',
  taxableValueCell:   'C30',
  igstAmountCell:     'F30',
  totalInrCell:       'H30',
  netWeightCell:      'C33',   // CONTROL B33/C33 = Net Weight
  grossWeightCell:    'F33',   // CONTROL E33/F33 = Gross Weight
  totalPackagesCell:  'C34',
  hsnCodeCell:        'C50',
  itemDescriptionCell:'C51',
  districtOriginCell: 'C52',
  stateOriginCell:    'C53',
  preCarriageCell:    'F22',   // F22 = BY ROAD (pre-carriage by road to port)
  modeTransportCell:  'F22',  // F22 also used for delivery_terms
  countryOriginCell:  'F23',   // CONTROL F23 = INDIA (C23 is terms of payment!)
  termsOfPaymentCell: 'C23',   // CONTROL C23 = payment terms (e.g. "CIF, 100% TT ADVANCE")
  notifyCell1:        'C14',
  notifyCell2:        'C15',
  webhookUrl: 'https://drawing-cosmetic-losses-nancy.trycloudflare.com/api/v1/invoices/',
};

function AMOUNTWORDS(n, curr) {
  curr = curr || 'USD';
  if (!n || n === 0) return '';
  const ones = ['','ONE','TWO','THREE','FOUR','FIVE','SIX','SEVEN','EIGHT','NINE','TEN',
    'ELEVEN','TWELVE','THIRTEEN','FOURTEEN','FIFTEEN','SIXTEEN','SEVENTEEN','EIGHTEEN','NINETEEN'];
  const tens = ['','','TWENTY','THIRTY','FORTY','FIFTY','SIXTY','SEVENTY','EIGHTY','NINETY'];
  function words(x) {
    if (x < 20)       return ones[x] || '';
    if (x < 100)      return tens[Math.floor(x/10)] + (x%10 ? ' '+ones[x%10] : '');
    if (x < 1000)     return ones[Math.floor(x/100)] + ' HUNDRED' + (x%100 ? ' AND '+words(x%100) : '');
    if (x < 100000)   return words(Math.floor(x/1000)) + ' THOUSAND' + (x%1000 ? ' '+words(x%1000) : '');
    if (x < 10000000) return words(Math.floor(x/100000)) + ' LAKH' + (x%100000 ? ' '+words(x%100000) : '');
    return words(Math.floor(x/10000000)) + ' CRORE' + (x%10000000 ? ' '+words(x%10000000) : '');
  }
  const whole = Math.floor(Math.abs(n));
  const cents = Math.round((Math.abs(n) - whole) * 100);
  const currText = curr === 'INR' ? 'INR Rupees' : 'US Dollar';
  let result = 'AMOUNT CHARGEABLE IN ' + currText + ' ' + words(whole) + ' ONLY';
  if (cents > 0) result += ' AND PAISE ' + words(cents) + ' ONLY';
  return result;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Horizon')
    .addItem('📄 Generate All Documents',    'generateDocuments')
    .addItem('📋 Generate PI Only',          'generatePIDocument')
    .addItem('✅ Validate Shipment',         'validateAndReport')
    .addSeparator()
    .addItem('🚗 Assign Vehicles (Sidebar)', 'showVehicleSidebar')
    .addItem('🔗 Bulk Assign by Model',      'bulkAssignByModel')
    .addItem('⚡ Quick Add Multiple Products','quickMultiProductEntry')
    .addSeparator()
    .addItem('📋 Clone Last Shipment',       'cloneLastShipment')
    .addItem('📊 Recent Shipments',          'showRecentShipments')
    .addSeparator()
    .addItem('📦 Export CHA Package',        'exportCHAPackage')
    .addSeparator()
    .addItem('🚢 Set Port Dropdowns',         'setupPortDropdowns')
    .addItem('👁 Assigned Vehicles Panel',    'showAssignedVehiclesPanel')
    .addItem('📅 Build Monthly Stock Tabs',   'buildMonthlyStockTabs')
    .addItem('✏️ Edit Exporter / Bank',        'showExporterBankEditor')
    .addSeparator()
    .addItem('🔄 Refresh Chassis Dropdown',    'updateChassisDropdown')
    .addItem('👤 Set Customer Dropdowns (Stock Q/R)', 'setupCustomerDropdowns')
    .addToUi();
  updateChassisDropdown();
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const name  = sheet.getName();
  const row   = e.range.getRow();
  const col   = e.range.getColumn();

  if (name === 'Stock' && col === 8 && row >= 4)
    sheet.getRange(row, 7).setValue(e.range.getValue() ? 'RESERVED' : 'AVAILABLE');
  if (name === 'Stock' && col === 9 && row >= 4 && e.range.getValue())
    sheet.getRange(row, 7).setValue('SHIPPED');

  // Refresh C17 chassis dropdown whenever invoice number in C8 changes
  if (name === 'CONTROL' && e.range.getA1Notation() === 'C8')
    updateChassisDropdown();
  
  // Auto-split customer dropdown in Stock col Q → contact_name (Q) + company_name (R)
if (name === 'Stock' && col === 17 && row >= 4 && e.value) {
  var val = String(e.value).trim();
  var sep = val.indexOf(' — ');
  if (sep !== -1) {
    sheet.getRange(row, 17).setValue(val.substring(0, sep).trim());   // Q = contact_name
    sheet.getRange(row, 18).setValue(val.substring(sep + 3).trim());  // R = company_name
  }
}
}

// Populates C17 with chassis entries where Stock col G = 'RESERVED' AND col H = current invoice
// Format: "chassis - engine - model - color" to match the existing dropdown display style
function updateChassisDropdown() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var control = ss.getSheetByName('CONTROL');
  var stock   = ss.getSheetByName('Stock');
  if (!control || !stock) return;

  var invoiceNo    = String(control.getRange('C8').getValue()).trim();
  var dropdownCell = control.getRange('C17');

  if (!invoiceNo) {
    dropdownCell.clearDataValidations();
    return;
  }

  var assignedChassis = stock.getRange('A4:H2000').getValues()
    .filter(function(r) {
      return r[0] && String(r[6]).trim() === 'RESERVED' && String(r[7]).trim() === invoiceNo;
    })
    .map(function(r) {
      var parts = [String(r[0])];
      if (r[1]) parts.push(String(r[1]));
      if (r[2]) parts.push(String(r[2]));
      if (r[3]) parts.push(String(r[3]));
      return parts.join(' - ');
    });

  if (assignedChassis.length === 0) {
    dropdownCell.clearDataValidations();
    dropdownCell.setNote('No RESERVED vehicles found for invoice: ' + invoiceNo);
    return;
  }

  dropdownCell.clearNote();
  dropdownCell.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(assignedChassis, true)
      .setAllowInvalid(true)
      .setHelpText('Reserved vehicles for invoice ' + invoiceNo + ' (' + assignedChassis.length + ' found)')
      .build()
  );
}

function validate() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl  = ss.getSheetByName('CONTROL');
  const stock = ss.getSheetByName('Stock');
  const errors = [], warnings = [];

  const inv      = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const customer = ctrl.getRange(CFG.customerCell).getValue();
  const stockRef = ctrl.getRange(CFG.stockCell).getValue();
  const qty      = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;
  const price    = Number(ctrl.getRange(CFG.unitPriceCell).getValue()) || 0;
  const mode     = ctrl.getRange(CFG.modeCell).getValue();
  const lc       = ctrl.getRange(CFG.lcCell).getValue();
  const container = ctrl.getRange(CFG.containerCell).getValue();

  if (!inv)       errors.push('① Invoice number missing          → fix: C8');
  if (!customer)  errors.push('② Customer name missing           → fix: C12');
  if (!stockRef)  errors.push('③ Stock / Chassis Ref missing     → fix: C17');
  if (qty <= 0)   errors.push('④ Quantity must be > 0            → fix: C26');
  if (price <= 0) errors.push('⑤ Unit price must be > 0          → fix: F26');

  if (!lc)        warnings.push('⚠ LC Number is empty              → cell: C24');
  if (!container) warnings.push('⚠ Container number is empty → cell: C37 (Optional except for Annexure C)');

  if (mode === 'FINAL') {
    const assigned = stock.getRange('A4:H2000').getValues().filter(function(r) { return r[7] === inv; });
    if (assigned.length === 0)
      errors.push('⑥ No vehicles assigned to this invoice — use sidebar or Bulk Assign');
    else if (assigned.length !== qty)
      errors.push('⑥ Assigned vehicles (' + assigned.length + ') ≠ quantity (' + qty + ') — mismatch will cause document errors');
  }

  if (mode !== 'FINAL' && mode !== 'DRAFT' && mode !== 'PROFORMA')
    warnings.push('⚠ Mode "' + mode + '" is not DRAFT, PROFORMA or FINAL  → fix: F3');

  return { errors: errors, warnings: warnings };
}

function validateAndReport() {
  const res      = validate();
  const errors   = res.errors;
  const warnings = res.warnings;
  const ui = SpreadsheetApp.getUi();

  if (errors.length > 0) {
    const errBlock  = '❌ ERRORS (must fix before generating):\n' + errors.join('\n');
    const warnBlock = warnings.length > 0 ? '\n\n⚠ WARNINGS (optional but check):\n' + warnings.join('\n') : '';
    ui.alert('Validation Failed', errBlock + warnBlock, ui.ButtonSet.OK);
    return;
  }
  if (warnings.length > 0) {
    const btn = ui.alert('Warnings Only',
      '⚠ No blocking errors, but review these:\n\n' + warnings.join('\n') + '\n\nYou can still generate documents.',
      ui.ButtonSet.OK_CANCEL);
    if (btn === ui.Button.OK) ui.alert('✅ Validation passed — ready to generate documents!');
    return;
  }
  ui.alert('✅ All Clear', 'No errors or warnings — ready to generate documents!', ui.ButtonSet.OK);
}

function buildPayload(ss, ctrl, inv) {
  const stock   = ss.getSheetByName('Stock');
  const company = ss.getSheetByName('Company');

  Logger.log('🔍 DEBUG: buildPayload v5.3 starting for invoice: ' + inv);

  const cif_usd   = Number(ctrl.getRange('C28').getValue()) || 0;
  const total_inr = Number(ctrl.getRange(CFG.totalInrCell).getValue()) || 0;

  const vehicles = stock.getRange('A4:J2000').getValues()
    .filter(function(r) { return r[7] === inv; })
    .map(function(r) {
      return {
        chassis_no:     r[0] || '',
        engine_no:      r[1] || '',
        model:          r[2] || '',
        color:          r[3] || '',
        year:           r[4] || '',
        unit_price_usd: Number(r[9]) || 0
      };
    });

  Logger.log('🚗 VEHICLES FOUND: ' + vehicles.length);
  vehicles.forEach(function(v, i) {
    Logger.log('  [' + i + '] chassis=' + v.chassis_no + ' | model="' + v.model + '" | price=' + v.unit_price_usd + ' | engine=' + v.engine_no + ' | color=' + v.color);
  });

  // Build product lookup from Products tab
  // Col A=ID, B=product_name (match key), C=HSN Code (8-digit), D=Commercial desc, E=SCOMET desc,
  // O(14)=Packing List desc, P(15)=Tax Invoice desc, Q(16)=PI Format desc
  var productMap = {};
  var productSheet = ss.getSheetByName('Products');
  if (productSheet) {
    productSheet.getRange('A3:R2000').getValues().forEach(function(row) {  // A3 skips header; R = desc_annexure1
      var modelKey = String(row[1]).toUpperCase().trim();  // col B = product_name (match key)
      if (modelKey) productMap[modelKey] = row;
    });
    Logger.log('📦 PRODUCT TAB keys (' + Object.keys(productMap).length + '): ' + Object.keys(productMap).join(', '));
  } else {
    Logger.log('⚠️ PRODUCT TAB not found — using defaults for all descriptions');
  }

  var defaultDesc     = ctrl.getRange(CFG.itemDescriptionCell || 'C51').getValue() || 'Motorcycles';
  var defaultHsn      = ctrl.getRange(CFG.hsnCodeCell || 'C50').getValue() || '8711';
  var defaultDistrict = ctrl.getRange(CFG.districtOriginCell || 'C52').getValue() || '';
  var defaultState    = ctrl.getRange(CFG.stateOriginCell    || 'C53').getValue() || '29';
  Logger.log('📋 DEFAULTS: desc="' + defaultDesc + '" | hsn="' + defaultHsn + '" | district="' + defaultDistrict + '" | state="' + defaultState + '"');

  // Group by model+price so each model gets its own HSN/description row
  var itemsObj = {};
  var controlUnitPrice = Number(ctrl.getRange(CFG.unitPriceCell).getValue()) || 0;
  vehicles.forEach(function(v) {
    var modelKey = String(v.model || '').toUpperCase().trim();

    var prod        = productMap[modelKey] || null;
    Logger.log('  🔎 model="' + v.model + '" → key="' + modelKey + '" → productMatch=' + (prod ? 'YES' : 'NO (using defaults)'));

    var hsnCode       = prod ? (String(prod[2]).trim()  || defaultHsn)      : defaultHsn;   // col C = HSN code
    var productName   = prod ? (String(prod[1]).trim()  || '')               : '';            // col B = product_name
    var descComm      = prod ? (String(prod[3]).trim()  || defaultDesc)      : defaultDesc;   // col D
    var descScomet    = prod ? (String(prod[4]).trim()  || defaultDesc)      : defaultDesc;   // col E
    var descPacking   = prod ? (String(prod[14]).trim() || defaultDesc)      : defaultDesc;   // col O
    var descTax       = prod ? (String(prod[15]).trim() || defaultDesc)      : defaultDesc;   // col P
    var descPi        = prod ? (String(prod[16]).trim() || defaultDesc)      : defaultDesc;   // col Q
    var descAnnexure1 = prod ? (String(prod[17]).trim() || descComm)         : descComm;      // col R
    var productPrice  = prod ? (Number(prod[12]) || 0)                       : 0;             // col M = default_price_usd
    var districtCode  = prod ? (String(prod[6]).trim()  || defaultDistrict)  : defaultDistrict;  // col G
    var stateCode     = prod ? (String(prod[7]).trim()  || defaultState)     : defaultState;     // col H
    var price    = v.unit_price_usd || productPrice || controlUnitPrice;
    var groupKey = modelKey + '||' + price.toFixed(2);

    Logger.log('    hsn="' + hsnCode + '" | product_name(B)="' + productName + '" | productPrice(M)=' + productPrice);
    Logger.log('    desc_pi(Q)="' + descPi + '" | desc_comm(D)="' + descComm + '" | desc_scomet(E)="' + descScomet + '"');
    Logger.log('    desc_pack(O)="' + descPacking + '" | desc_tax(P)="' + descTax + '" | desc_ann1(R)="' + descAnnexure1 + '"');
    Logger.log('    district(G)="' + districtCode + '" | state(H)="' + stateCode + '"');

    // Append product_name (col B) and model name to all descriptions
    var modelDisplay = v.model || modelKey;
    function withModel(base) {
      var parts = [base];
      if (productName && productName !== base) parts.push(productName);
      if (modelDisplay && parts.indexOf(modelDisplay) === -1) parts.push(modelDisplay);
      return parts.filter(Boolean).join(' ');
    }

    if (!itemsObj[groupKey]) {
      itemsObj[groupKey] = {
        hsn_code:               hsnCode,
        description:            withModel(descPi),  // default = PI FORMAT desc; backend overrides per-template
        description_commercial: withModel(descComm),
        description_scomet:     withModel(descScomet),
        description_packing:    withModel(descPacking),
        description_tax:        withModel(descTax),
        description_pi:         withModel(descPi),
        description_annexure1:  withModel(descAnnexure1),
        quantity:               0,
        rate_per_unit:          price,
        amount_usd:             0,
        district_origin_code:   districtCode,
        state_origin_code:      stateCode
      };
      Logger.log('    ✅ NEW item group: key="' + groupKey + '" | desc_pi="' + withModel(descPi) + '"');
    }
    itemsObj[groupKey].quantity   += 1;
    itemsObj[groupKey].amount_usd += price;
  });
  var items = [];
  Object.keys(itemsObj).forEach(function(k) { items.push(itemsObj[k]); });
  Logger.log('📊 ITEMS BUILT: ' + items.length + ' group(s)');
  items.forEach(function(it, i) {
    Logger.log('  [' + i + '] hsn=' + it.hsn_code + ' | qty=' + it.quantity + ' | rate=' + it.rate_per_unit + ' | amount=' + it.amount_usd);
    Logger.log('       desc_pi="' + it.description_pi + '"');
    Logger.log('       desc_commercial="' + it.description_commercial + '"');
    Logger.log('       desc_scomet="' + it.description_scomet + '"');
  });

  // Resolve buyer details: CONTROL C12 = smart dropdown "CONTACT — COMPANY"
  // Look up in Customers sheet (col L = smart_dropdown, B = company_name, C/D/E = address, F = country)
  var customerDropdownVal = String(ctrl.getRange(CFG.customerCell).getValue() || '').trim();
  var buyerName    = customerDropdownVal;
  var buyerAddress = String(ctrl.getRange(CFG.buyerAddressCell).getValue() || '').trim();
  var buyerCountry = String(ctrl.getRange(CFG.buyerCountryCell).getValue() || '').trim();
  var customersSheet = ss.getSheetByName('Customers');
  if (customersSheet && customerDropdownVal) {
    var custRows = customersSheet.getRange('A4:L2000').getValues();
    for (var ci = 0; ci < custRows.length; ci++) {
      var cr = custRows[ci];
      if (!cr[0] && !cr[11]) break;
      var smartVal = String(cr[11]).trim();  // col L = smart_dropdown "CONTACT — COMPANY"
      if (smartVal && smartVal === customerDropdownVal) {
        buyerName    = String(cr[1]).trim() || buyerName;   // col B = company_name
        var addr1    = String(cr[2]).trim();                 // col C = address_line1
        var addr2    = String(cr[3]).trim();                 // col D = address_line2
        var custCity = String(cr[4]).trim();                 // col E = city
        buyerAddress = buyerAddress || [addr1, addr2, custCity].filter(Boolean).join(', ');
        buyerCountry = buyerCountry || String(cr[5]).trim(); // col F = country
        Logger.log('👤 Customer lookup matched: name="' + buyerName + '" | addr="' + buyerAddress + '" | country="' + buyerCountry + '"');
        break;
      }
    }
  }
  if (!customersSheet || buyerName === customerDropdownVal) {
    Logger.log('⚠️ Customer lookup: no match for "' + customerDropdownVal + '" — using raw dropdown value');
  }

  const payload = {
    invoice_no:      inv,
    invoice_date:    Utilities.formatDate(ctrl.getRange(CFG.dateCell).getValue(), 'GMT+5:30', 'dd.MM.yyyy') || '',
    mode:            ctrl.getRange(CFG.modeCell).getValue() || 'FINAL',
    generation_date: Utilities.formatDate(new Date(), 'GMT+5:30', 'dd.MM.yyyy'),

    exporter: {
      company_name:   company.getRange('B2').getValue() || '',
      address:        company.getRange('B3').getValue() || '',
      phone:          company.getRange('B4').getValue() || '',
      iec:            company.getRange('B5').getValue() || '',
      pan:            company.getRange('B6').getValue() || '',
      gstin:          company.getRange('B7').getValue() || '',
      signatory:      company.getRange('B12').getValue() || '',
      cha:            company.getRange('B13').getValue() || '',
      customs_office: company.getRange('B14').getValue() || ''
    },

    buyer: {
      name:    buyerName,
      address: buyerAddress,
      country: buyerCountry
    },

    shipping: {
      pre_carriage_by:        ctrl.getRange(CFG.preCarriageCell || 'F20').getValue() || '',
      mode_of_transport:      ctrl.getRange(CFG.modeTransportCell).getValue() || '',
      country_of_origin:      ctrl.getRange(CFG.countryOriginCell || 'F23').getValue() || 'INDIA',
      country_of_destination: ctrl.getRange(CFG.buyerCountryCell).getValue() || '',
      port_of_loading:        ctrl.getRange(CFG.portLoadCell).getValue() || '',
      port_of_discharge:      ctrl.getRange(CFG.portDischargeCell).getValue() || '',
      final_destination:      ctrl.getRange(CFG.finalDestinationCell).getValue() || '',
      container_no:           ctrl.getRange(CFG.containerCell).getValue() || ''
    },

    financials: {
      quantity:          Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0,
      unit_price_usd:    Number(ctrl.getRange(CFG.unitPriceCell).getValue()) || 0,
      fob_total_usd:     Number(ctrl.getRange('C27').getValue()) || 0,
      freight_usd:       Number(ctrl.getRange('F27').getValue()) || 0,
      insurance_usd:     Number(ctrl.getRange('H27').getValue()) || 0,
      cif_total_usd:     cif_usd,
      exchange_rate:     Number(ctrl.getRange(CFG.exchangeRateCell).getValue()) || 0,
      igst_rate:         Number(ctrl.getRange(CFG.igstRateCell).getValue()) || 0,
      taxable_value_inr: Number(ctrl.getRange(CFG.taxableValueCell).getValue()) || 0,
      igst_amount_inr:   Number(ctrl.getRange(CFG.igstAmountCell).getValue()) || 0,
      total_value_inr:   total_inr
    },

    bank: {
      bank_name:  company.getRange('B8').getValue() || '',
      account_no: company.getRange('B9').getValue() || '',
      swift:      company.getRange('B10').getValue() || '',
      branch:     company.getRange('B11').getValue() || ''
    },

    weights: {
      net_weight_kg:   Number(ctrl.getRange(CFG.netWeightCell || 'C33').getValue()) || 0,
      gross_weight_kg: Number(ctrl.getRange(CFG.grossWeightCell || 'F33').getValue()) || 0,
      total_packages:  Number(ctrl.getRange(CFG.totalPackagesCell || 'C34').getValue()) || vehicles.length
    },

    vehicles: vehicles,
    items:    items,

    lc_number:             String(ctrl.getRange(CFG.lcCell).getValue() || ''),
    buyers_order_no:       String(ctrl.getRange('C15').getValue() || ''),
    notify_1:              String(ctrl.getRange(CFG.notifyCell1).getValue() || ''),
    notify_2:              String(ctrl.getRange(CFG.notifyCell2).getValue() || ''),
    terms_of_payment:      String(ctrl.getRange(CFG.termsOfPaymentCell || 'C23').getValue() || ''),
    company_seal_no:       '',
    shipping_line_seal_no: '',
    marks_and_numbers:     String(ctrl.getRange('C16').getValue() || ''),
    scomet_product_desc:   items.length > 0 ? (items[0].description_scomet || '') : '',
    amount_usd_words:      AMOUNTWORDS(cif_usd, 'USD'),
    amount_inr_words:      AMOUNTWORDS(total_inr, 'INR'),

    // PI FORMAT fields
    insurance_ref_no: String(ctrl.getRange(CFG.lcCell).getValue() || ''),
    delivery_terms:   String(ctrl.getRange(CFG.modeTransportCell).getValue() || ''),
    place_of_receipt: String(ctrl.getRange(CFG.portLoadCell).getValue() || '')
  };

  Logger.log('✅ Payload built | Vehicles: ' + vehicles.length + ' | Items: ' + items.length);
  Logger.log('💰 Financials: cif_usd=' + cif_usd + ' | total_inr=' + total_inr + ' | exchange_rate=' + payload.financials.exchange_rate);
  Logger.log('🚢 Shipping: port_load="' + payload.shipping.port_of_loading + '" | port_discharge="' + payload.shipping.port_of_discharge + '" | dest="' + payload.shipping.final_destination + '"');
  Logger.log('📄 PI fields: lc_number="' + payload.lc_number + '" | insurance_ref_no="' + payload.insurance_ref_no + '" | delivery_terms="' + payload.delivery_terms + '"');
  return payload;
}

// ── Generate PI FORMAT only ───────────────────────────────────────────────────
function generatePIDocument() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName('CONTROL');
  const inv  = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const ui   = SpreadsheetApp.getUi();

  if (!inv) { ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK); return; }

  Logger.log('📄 PI GENERATION starting for invoice: ' + inv);

  const payload    = buildPayload(ss, ctrl, inv);
  const encodedInv = encodeURIComponent(inv);
  // Pass documents=proforma_invoice so backend generates PI FORMAT only
  const fullUrl    = CFG.webhookUrl + 'generate?invoice_no=' + encodedInv + '&documents=proforma_invoice';

  Logger.log('📤 PI URL: ' + fullUrl);
  Logger.log('📦 PI payload items: ' + JSON.stringify(payload.items));
  Logger.log('📦 PI payload vehicles: ' + JSON.stringify(payload.vehicles));

  try {
    const res  = UrlFetchApp.fetch(fullUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText();

    Logger.log('📥 PI Response HTTP ' + code + ': ' + body.substring(0, 500));

    if (code === 200 || code === 201) {
      logAudit(inv, 'PI_GENERATED', 'PI FORMAT only');
      const parsed = JSON.parse(body);
      // Show file link directly
      const files = parsed.generated_files || [];
      if (files.length === 0) {
        ui.alert('⚠ PI Generated', 'Document was processed but no download link returned.\nCheck backend logs.', ui.ButtonSet.OK);
        return;
      }
      const driveUrl = files[0].gcs_url || '';
      const fileName = files[0].download_name || ('PI_FORMAT_' + inv + '.docx');
      const linkHtml = driveUrl
        ? '<p style="margin:12px 0"><a href="' + driveUrl + '" target="_blank" style="background:#1e3a5f;color:white;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:14px">⬇️ Download PI FORMAT</a></p>'
          + '<p style="font-size:11px;color:#666;margin-top:8px;word-break:break-all">' + driveUrl + '</p>'
        : '<p style="color:#dc2626">No download link returned — check backend logs.</p>';
      SpreadsheetApp.getUi().showModalDialog(
        HtmlService.createHtmlOutput(
          '<style>body{font-family:Arial,sans-serif;padding:20px;text-align:center}</style>' +
          '<h3 style="color:#1e3a5f">✅ PI FORMAT Generated</h3>' +
          '<p><strong>Invoice:</strong> ' + inv + '</p>' +
          '<p><strong>File:</strong> ' + fileName + '</p>' +
          linkHtml
        ).setWidth(480).setHeight(200),
        'PI FORMAT — ' + inv
      );
    } else {
      logAudit(inv, 'PI_FAILED', 'HTTP ' + code);
      ui.alert('❌ PI Generation Failed (HTTP ' + code + ')', body.substring(0, 800), ui.ButtonSet.OK);
    }
  } catch(err) {
    Logger.log('❌ PI Error: ' + err);
    ui.alert('❌ Connection Failed', 'Could not reach backend.\n\nError: ' + String(err), ui.ButtonSet.OK);
  }
}

function generateDocuments() {
  const res      = validate();
  const errors   = res.errors;
  const warnings = res.warnings;
  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert('❌ Fix these errors first:\n\n' + errors.join('\n'));
    return;
  }

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName('CONTROL');
  const inv  = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const mode = ctrl.getRange(CFG.modeCell).getValue();
  const qty  = ctrl.getRange(CFG.qtyCell).getValue();
  const cif  = ctrl.getRange('C28').getValue();
  const ui   = SpreadsheetApp.getUi();

  if (mode === 'FINAL') {
    const warnNote = warnings.length > 0 ? '\n\n⚠ Warnings:\n' + warnings.join('\n') : '';
    const summary = '📋 GENERATION SUMMARY\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
                    'Invoice : ' + inv + '\nMode    : ' + mode + '\n' +
                    'Customer: ' + ctrl.getRange(CFG.customerCell).getValue() + '\n' +
                    'Qty     : ' + qty + ' vehicles\nCIF     : USD ' + cif + '\n' +
                    'LC No.  : ' + (ctrl.getRange(CFG.lcCell).getValue() || '(none)') +
                    warnNote + '\n\nProceed?';
    if (ui.alert('Confirm', summary, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }

  const payload    = buildPayload(ss, ctrl, inv);
  const encodedInv = encodeURIComponent(inv);
  const fullUrl    = CFG.webhookUrl + 'generate?invoice_no=' + encodedInv;

  try {
    const res2 = UrlFetchApp.fetch(fullUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const code = res2.getResponseCode();
    const body = res2.getContentText();

    if (code === 200 || code === 201) {
      logAudit(inv, 'GENERATED', 'Mode:' + mode);
      showDownloadDialog(JSON.parse(body), inv);
    } else {
      logAudit(inv, 'FAILED', 'HTTP ' + code);
      ui.alert('❌ Backend Error (HTTP ' + code + ')', body.substring(0, 800), ui.ButtonSet.OK);
    }
  } catch(err) {
    Logger.log('Error: ' + err);
    ui.alert('❌ Connection Failed', String(err), ui.ButtonSet.OK);
  }
}

function showDownloadDialog(result, invoiceNo) {
  const files      = result.generated_files || [];
  const failed     = result.failed || [];
  const folderUrl  = result.drive_folder_url || '';

  let html =
    '<style>body{font-family:Arial,sans-serif;padding:20px;background:#f8fafc}h2{color:#1e3a5f;text-align:center;margin-bottom:4px}' +
    '.sub{text-align:center;color:#64748b;font-size:13px;margin-bottom:14px}' +
    '.folder-btn{display:block;width:100%;padding:12px;background:#1e3a5f;color:white;text-align:center;text-decoration:none;font-size:14px;font-weight:bold;border-radius:8px;margin-bottom:14px}' +
    '.file-item{padding:10px 12px;margin:6px 0;background:white;border-radius:8px;display:flex;justify-content:space-between;align-items:center;box-shadow:0 1px 3px rgba(0,0,0,.1)}' +
    '.file-name{font-size:12px;color:#334155;flex:1}' +
    '.download-btn{background:#0d9488;color:white;padding:6px 14px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:12px;white-space:nowrap;margin-left:10px}' +
    '.failed-item{padding:8px 12px;margin:6px 0;background:#fee2e2;border-radius:8px;font-size:12px;color:#dc2626}' +
    '.section{font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.5px;margin:12px 0 6px}' +
    '</style>' +
    '<h2>✅ Documents Generated!</h2>' +
    '<p class="sub"><strong>Invoice:</strong> ' + invoiceNo + ' &nbsp;·&nbsp; ' + files.length + ' file(s) ready</p>';

  if (folderUrl)
    html += '<a href="' + folderUrl + '" target="_blank" class="folder-btn">📂 Open Drive Folder</a>';

  if (files.length > 0) {
    html += '<div class="section">Downloads</div>';
    files.forEach(function(file) {
      const docName = file.document || file.download_name || file.template || 'Document';
      const url     = file.gcs_url || file.drive_url || '';
      if (url)
        html += '<div class="file-item"><span class="file-name">' + docName + '</span><a href="' + url + '" target="_blank" class="download-btn">⬇️ Download</a></div>';
    });
  }

  if (failed.length > 0) {
    html += '<div class="section">Failed</div>';
    failed.forEach(function(f) {
      html += '<div class="failed-item">❌ ' + (f.template || f.document || 'Unknown') + ': ' + (f.error || 'error') + '</div>';
    });
  }

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(560).setHeight(560),
    'Documents — ' + invoiceNo
  );
}

function logAudit(inv, action, notes) {
  const audit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit_Trail');
  if (!audit) return;
  const ts   = Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail() || 'unknown';
  audit.appendRow([ts, inv, action, user, 'v5.2', notes || '']);
}

function showVehicleSidebar() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl    = ss.getSheetByName('CONTROL');
  const stock   = ss.getSheetByName('Stock');
  const invoice = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const qty     = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;

  if (!invoice) { SpreadsheetApp.getUi().alert('❌ No invoice number found in C8.'); return; }

  const data    = stock.getRange('A4:H2000').getValues();
  const visible = data.filter(function(r) { return r[0] && (r[6] === 'AVAILABLE' || r[7] === invoice); });
  const alreadyAssigned = data.filter(function(r) { return r[0] && r[7] === invoice; }).length;

  const rows = visible.map(function(r) {
    const isAssigned  = r[7] === invoice;
    const statusColor = r[6] === 'AVAILABLE' ? '#16a34a' : '#d97706';
    const assignedTo  = r[7] && r[7] !== invoice
      ? '<span style="color:#dc2626;font-size:10px">Taken: ' + r[7] + '</span>'
      : r[7] === invoice ? '<span style="color:#d97706;font-size:10px">This invoice</span>' : '';
    return '<tr class="vrow" style="background:' + (isAssigned ? '#FEF9C3' : 'white') + '">' +
      '<td style="text-align:center"><input type="checkbox" class="vchk" data-chassis="' + r[0] + '" ' + (isAssigned ? 'checked' : '') + '></td>' +
      '<td style="font-family:monospace;font-size:11px">' + r[0] + '</td>' +
      '<td>' + r[2] + '</td><td>' + r[3] + '</td>' +
      '<td style="font-weight:600;color:' + statusColor + '">' + r[6] + '</td>' +
      '<td>' + assignedTo + '</td></tr>';
  }).join('');

  const sidebarHtml = HtmlService.createHtmlOutput(
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:12px;background:#f8fafc}' +
    '.header{background:#1e3a5f;color:white;padding:10px 14px;border-radius:8px;margin-bottom:12px}' +
    '.header h3{font-size:14px;font-weight:600}.header p{font-size:11px;opacity:.8;margin-top:3px}' +
    '.counter-bar{display:flex;align-items:center;gap:10px;background:white;border:1px solid #e2e8f0;border-radius:8px;padding:8px 12px;margin-bottom:10px}' +
    '.counter-text{font-size:13px;font-weight:600;flex:1}' +
    '.counter-text.ok{color:#16a34a}.counter-text.over{color:#dc2626}.counter-text.under{color:#d97706}' +
    '.progress-wrap{height:6px;background:#e2e8f0;border-radius:99px;width:120px;overflow:hidden}' +
    '.progress-fill{height:100%;background:#0d9488;border-radius:99px;transition:width .3s}' +
    'table{width:100%;border-collapse:collapse;background:white;border-radius:8px;border:1px solid #e2e8f0}' +
    'thead th{background:#1e3a5f;color:white;padding:7px 8px;font-size:11px;font-weight:600;text-align:left}' +
    'tbody td{padding:6px 8px;border-bottom:1px solid #f1f5f9;font-size:12px}' +
    'tbody tr:last-child td{border-bottom:none}tbody tr:hover{background:#f0f9ff!important}' +
    'input[type=checkbox]{width:15px;height:15px;cursor:pointer;accent-color:#0d9488}' +
    '.btn{width:100%;padding:11px;background:#0d9488;color:white;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;margin-top:10px}' +
    '.btn:hover:not(:disabled){background:#0f766e}.btn:disabled{opacity:.45;cursor:not-allowed}' +
    '#msg{margin-top:10px;padding:8px 12px;border-radius:6px;font-size:12px;font-weight:600;display:none}' +
    '#msg.success{background:#dcfce7;color:#15803d;display:block}#msg.error{background:#fee2e2;color:#dc2626;display:block}' +
    '</style>' +
    '<div class="header"><h3>🚗 Assign Vehicles</h3>' +
    '<p>Invoice: <strong>' + invoice + '</strong> &nbsp;·&nbsp; Need: <strong>' + qty + '</strong> &nbsp;·&nbsp; Showing: <strong>' + visible.length + '</strong> available</p></div>' +
    '<div class="counter-bar"><span class="counter-text under" id="counterText">Select vehicles below</span>' +
    '<div class="progress-wrap"><div class="progress-fill" id="progressFill" style="width:' + Math.min(100,alreadyAssigned/Math.max(qty,1)*100) + '%"></div></div></div>' +
    '<table><thead><tr><th>✓</th><th>Chassis No.</th><th>Model</th><th>Colour</th><th>Status</th><th>Assignment</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>' +
    '<button class="btn" id="assignBtn" onclick="assignSelected()" disabled>Assign Selected Vehicles</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'var QTY=' + qty + ';' +
    'function updateCounter(){' +
    'var n=document.querySelectorAll(".vchk:checked").length;' +
    'var txt=document.getElementById("counterText"),bar=document.getElementById("progressFill"),btn=document.getElementById("assignBtn");' +
    'bar.style.width=Math.min(100,n/Math.max(QTY,1)*100)+"%";btn.disabled=n===0;' +
    'if(n===0){txt.textContent="Select vehicles below";txt.className="counter-text under";bar.style.background="#0d9488";}' +
    'else if(n<QTY){txt.textContent=n+" selected — need "+(QTY-n)+" more";txt.className="counter-text under";bar.style.background="#d97706";}' +
    'else if(n===QTY){txt.textContent="✅ "+n+" selected — perfect match!";txt.className="counter-text ok";bar.style.background="#16a34a";}' +
    'else{txt.textContent="⚠ "+n+" selected — "+(n-QTY)+" over";txt.className="counter-text over";bar.style.background="#dc2626";}}' +
    'document.querySelectorAll(".vchk").forEach(function(c){c.addEventListener("change",function(){this.closest("tr").style.background=this.checked?"#FEF9C3":"";updateCounter();});});' +
    'updateCounter();' +
    'function assignSelected(){' +
    'var sel=[].slice.call(document.querySelectorAll(".vchk:checked")).map(function(c){return c.dataset.chassis;});' +
    'var btn=document.getElementById("assignBtn"),msg=document.getElementById("msg");' +
    'btn.disabled=true;btn.textContent="Saving…";msg.className="";msg.style.display="none";' +
    'google.script.run' +
    '.withSuccessHandler(function(r){msg.textContent="✅ "+r;msg.className="success";btn.textContent="Assign Selected Vehicles";btn.disabled=false;})' +
    '.withFailureHandler(function(e){msg.textContent="❌ Error: "+e.message;msg.className="error";btn.textContent="Assign Selected Vehicles";btn.disabled=false;})' +
    '.assignVehiclesFromSidebar(sel,"' + invoice + '");}' +
    '<\/script>'
  ).setTitle('Assign Vehicles — ' + invoice).setWidth(780);
  SpreadsheetApp.getUi().showSidebar(sidebarHtml);
}

function assignVehiclesFromSidebar(chassisList, invoiceNo) {
  const stock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock');
  const data  = stock.getRange('A4:H2000').getValues();
  let assigned = 0, released = 0;
  data.forEach(function(row, i) {
    if (!row[0]) return;
    const chassis = String(row[0]);
    if (chassisList.indexOf(chassis) !== -1) {
      stock.getRange(i+4, 7).setValue('RESERVED');
      stock.getRange(i+4, 8).setValue(invoiceNo);
      assigned++;
    } else if (row[7] === invoiceNo) {
      stock.getRange(i+4, 7).setValue('AVAILABLE');
      stock.getRange(i+4, 8).setValue('');
      released++;
    }
  });
  logAudit(invoiceNo, 'VEHICLES_ASSIGNED', assigned + ' assigned, ' + released + ' released');
  updateChassisDropdown();
  return assigned + ' vehicles assigned to ' + invoiceNo + (released ? ' | ' + released + ' released.' : '.');
}

function bulkAssignByModel() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl  = ss.getSheetByName('CONTROL');
  const stock = ss.getSheetByName('Stock');

  const inv = ctrl.getRange(CFG.invoiceNoCell).getValue();
  if (!inv) return ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK);

  const qty = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;
  if (qty <= 0) return ui.alert('❌ Quantity is Zero', 'Update quantity in C26 first.', ui.ButtonSet.OK);

  const alreadyAssigned = stock.getRange('A4:H2000').getValues()
    .filter(function(r) { return r[7] === inv; }).length;
  if (alreadyAssigned >= qty)
    return ui.alert('✅ Already Fully Assigned',
      'Invoice ' + inv + ' already has all ' + qty + ' vehicles assigned.\n\nUse the sidebar to review.', ui.ButtonSet.OK);

  const remaining = qty - alreadyAssigned;
  const resp = ui.prompt('🔗 Bulk Assign by Model — ' + inv,
    'Need ' + remaining + ' more vehicle(s)  (' + alreadyAssigned + ' of ' + qty + ' already assigned).\n\n' +
    'Enter model name — partial match works (e.g. "PULSAR" matches "PULSAR NS200"):',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const filter = resp.getResponseText().toUpperCase().trim();
  if (!filter) return ui.alert('❌ Empty Input', 'Please enter a model name.', ui.ButtonSet.OK);

  const allData = stock.getRange('A4:H2000').getValues();
  const matches = [];
  for (var i = 0; i < allData.length; i++) {
    const row = allData[i];
    if (!row[0] || row[6] !== 'AVAILABLE') continue;
    if (String(row[2]).toUpperCase().indexOf(filter) === -1) continue;
    matches.push({ rowIndex: i, chassis: row[0], model: row[2], color: row[3] });
  }

  if (matches.length === 0)
    return ui.alert('⚠ No Vehicles Found',
      'No AVAILABLE vehicles matched "' + filter + '".\n\n' +
      'Tips:\n• Try a shorter term\n• Check the Stock sheet for the exact model name\n• All matching vehicles may already be RESERVED or SHIPPED',
      ui.ButtonSet.OK);

  const toAssign   = matches.slice(0, remaining);
  const willAssign = toAssign.length;
  const isPartial  = willAssign < remaining;

  let preview = toAssign.slice(0, 8).map(function(v, i) {
    return (i+1)+'. '+v.chassis+'  |  '+v.model+'  |  '+v.color;
  }).join('\n');
  if (willAssign > 8) preview += '\n... and ' + (willAssign-8) + ' more.';

  let msg = '📋 ASSIGNMENT PREVIEW\n' + Array(37).join('─') + '\n' +
            'Invoice  : ' + inv + '\nModel    : ' + filter + '\nAssigning: ' + willAssign + ' vehicle(s)\n' +
            Array(37).join('─') + '\n\n' + preview + '\n\n';
  if (isPartial) msg += '⚠ Only ' + willAssign + ' matching vehicles available.\nStill need ' + (remaining-willAssign) + ' more after this.\n\n';
  msg += 'Proceed with assignment?';

  if (ui.alert('Confirm Bulk Assignment', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES)
    return ui.alert('Cancelled', 'No changes were made.', ui.ButtonSet.OK);

  toAssign.forEach(function(v) {
    stock.getRange(v.rowIndex+4, 7).setValue('RESERVED');
    stock.getRange(v.rowIndex+4, 8).setValue(inv);
  });

  logAudit(inv, 'BULK_ASSIGN', willAssign + ' "' + filter + '" assigned' +
    (isPartial ? ' (PARTIAL — ' + (remaining-willAssign) + ' still needed)' : ''));
  updateChassisDropdown();

  const total     = alreadyAssigned + willAssign;
  const stillNeed = qty - total;
  ui.alert(isPartial ? '⚠ Partial Assignment' : '✅ Assignment Complete',
    '✅ ' + willAssign + ' vehicle(s) assigned to ' + inv + '\n\nProgress : ' + total + ' / ' + qty + ' assigned\n' +
    (stillNeed > 0 ? 'Still need: ' + stillNeed + ' more\n\n→ Run Bulk Assign again or use the sidebar.' : '\n🎉 Fully assigned — ready to generate documents!'),
    ui.ButtonSet.OK);
}

function cloneLastShipment() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl  = ss.getSheetByName('CONTROL');
  const stock = ss.getSheetByName('Stock');
  const audit = ss.getSheetByName('Audit_Trail');

  const currentInv = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const currentQty = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;

  if (currentInv) {
    const reserved = stock.getRange('A4:H2000').getValues()
      .filter(function(r) { return r[7] === currentInv; }).length;
    if (reserved > 0 && currentQty > 0 && reserved < currentQty) {
      if (ui.alert('⚠ Incomplete Assignment',
        'Invoice ' + currentInv + ' only has ' + reserved + ' of ' + currentQty + ' vehicles assigned.\n\nAre you sure you want to clone?\n(Current vehicles stay reserved — only CONTROL fields are copied.)',
        ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    }
  }

  const resp = ui.prompt('📋 Clone Shipment',
    'Current invoice: ' + (currentInv || '(none)') + '\n\nNew invoice number:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const newInv = resp.getResponseText().trim();
  if (!newInv) { ui.alert('❌ Cancelled', 'No invoice number entered.', ui.ButtonSet.OK); return; }
  if (newInv === currentInv) { ui.alert('❌ Same Invoice', 'New invoice number is identical. Nothing changed.', ui.ButtonSet.OK); return; }

  if (audit) {
    const alreadyExists = audit.getRange('A4:B2000').getValues()
      .some(function(r) { return String(r[1]) === newInv; });
    if (alreadyExists) {
      if (ui.alert('⚠ Invoice Already Exists',
        'Invoice "' + newInv + '" already appears in the Audit Trail.\n\nProceed anyway?',
        ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    }
  }

  ctrl.getRange(CFG.invoiceNoCell).setValue(newInv);
  ctrl.getRange(CFG.dateCell).setValue(Utilities.formatDate(new Date(), 'GMT+5:30', 'dd.MM.yyyy'));
  logAudit(newInv, 'CLONED', 'Cloned from ' + (currentInv || 'blank'));

  ui.alert('✅ Shipment Cloned',
    'Invoice set to: ' + newInv + '\nDate updated to today.\n\n' +
    'Next steps:\n1. Update quantity in C26 if different\n2. Update customer / LC / container details as needed\n3. Use Assign Vehicles sidebar to pick the vehicles',
    ui.ButtonSet.OK);
}

function showRecentShipments() {
  const audit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit_Trail');
  if (!audit) {
    SpreadsheetApp.getUi().alert('⚠ No Audit Trail', 'The Audit_Trail sheet was not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const data = audit.getRange('A4:F2000').getValues().filter(function(r) { return r[0]; });
  if (data.length === 0) {
    SpreadsheetApp.getUi().alert('📊 No Data', 'The audit trail is empty — no shipments recorded yet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const recent = data.slice(-30).reverse();
  const badgeColors = {
    'GENERATED':        { bg: '#dcfce7', color: '#15803d' },
    'BULK_ASSIGN':      { bg: '#dbeafe', color: '#1d4ed8' },
    'VEHICLES_ASSIGNED':{ bg: '#dbeafe', color: '#1d4ed8' },
    'CLONED':           { bg: '#fef9c3', color: '#92400e' },
    'FAILED':           { bg: '#fee2e2', color: '#dc2626' },
    'DEFAULT':          { bg: '#f1f5f9', color: '#475569' }
  };

  const rows = recent.map(function(r) {
    const action = String(r[2]);
    const style  = badgeColors[action] || badgeColors['DEFAULT'];
    return '<tr>' +
      '<td style="color:#64748b;font-size:11px;white-space:nowrap">' + r[0] + '</td>' +
      '<td style="font-weight:700;font-family:monospace;font-size:11px">' + r[1] + '</td>' +
      '<td><span style="background:' + style.bg + ';color:' + style.color + ';padding:2px 7px;border-radius:99px;font-size:10px;font-weight:700">' + action + '</span></td>' +
      '<td style="color:#64748b;font-size:11px">' + r[3] + '</td>' +
      '<td style="color:#475569;font-size:11px">' + r[5] + '</td></tr>';
  }).join('');

  SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(
    '<style>*{box-sizing:border-box}body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:14px;background:#f8fafc}' +
    'h3{color:#1e3a5f;margin:0 0 12px;font-size:15px}table{width:100%;border-collapse:collapse;background:white;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}' +
    'thead th{background:#1e3a5f;color:white;padding:8px 10px;font-size:11px;font-weight:600;text-align:left}' +
    'tbody td{padding:7px 10px;border-bottom:1px solid #f1f5f9;vertical-align:middle}tbody tr:last-child td{border-bottom:none}tbody tr:hover{background:#f0f9ff}</style>' +
    '<h3>📊 Recent Shipments <span style="font-size:11px;font-weight:400;color:#94a3b8">(last 30 actions)</span></h3>' +
    '<table><thead><tr><th>Timestamp</th><th>Invoice</th><th>Action</th><th>User</th><th>Notes</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>'
  ).setTitle('Recent Shipments').setWidth(740).setHeight(480), 'Recent Shipments');
}

function exportCHAPackage() {
  const ctrl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL');
  const inv  = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const ui   = SpreadsheetApp.getUi();
  if (!inv) { ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK); return; }

  const fullUrl = CFG.webhookUrl + encodeURIComponent(inv) + '/cha-package';
  Logger.log('🚀 Calling CHA Package endpoint: ' + fullUrl);

  try {
    const res  = UrlFetchApp.fetch(fullUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ invoice_no: inv }), muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code === 200 || code === 201) {
      logAudit(inv, 'CHA_PACKAGE', 'Success');
      let folderUrl = '';
      try { const j = JSON.parse(body); folderUrl = j.drive_folder_url || j.folder_url || ''; } catch(e) {}
      ui.alert('🎉 CHA Package Ready',
        '✅ CHA Package Created!\n\nInvoice: ' + inv + '\n' + (folderUrl ? '🔗 Folder link opening...' : 'Check Google Drive for CHA_' + inv),
        ui.ButtonSet.OK);
      if (folderUrl)
        SpreadsheetApp.getUi().showModelessDialog(
          HtmlService.createHtmlOutput('<script>window.open("' + folderUrl + '","_blank");<\/script>'),
          'Opening Drive Folder...');
    } else {
      logAudit(inv, 'CHA_FAILED', 'HTTP ' + code);
      ui.alert('❌ Backend Error (HTTP ' + code + ')', body.substring(0, 500), ui.ButtonSet.OK);
    }
  } catch(err) {
    logAudit(inv, 'CHA_FAILED', String(err));
    ui.alert('❌ Connection Failed', 'Could not reach backend.\n\nError: ' + err, ui.ButtonSet.OK);
  }
}

function quickMultiProductEntry() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName('CONTROL');
  const inv  = ctrl.getRange(CFG.invoiceNoCell).getValue();

  if (!inv) {
    SpreadsheetApp.getUi().alert('❌ No Invoice', 'Set an invoice number in C8 before adding products.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const html = HtmlService.createHtmlOutput(
    '<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:14px;background:#f8fafc}' +
    '.header{background:#1e3a5f;color:white;padding:10px 14px;border-radius:8px;margin-bottom:12px}.header h3{font-size:14px;font-weight:600}.header p{font-size:11px;opacity:.8;margin-top:3px}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:8px}thead th{background:#e2e8f0;color:#475569;padding:6px 8px;font-size:11px;font-weight:600;text-align:left}tbody td{padding:4px}' +
    'input[type=text],input[type=number]{width:100%;padding:6px 8px;border:1px solid #cbd5e1;border-radius:5px;font-size:12px;background:white}input:focus{outline:none;border-color:#0d9488}' +
    '.del-btn{background:none;border:none;color:#dc2626;font-size:16px;cursor:pointer;padding:4px 8px}' +
    '.totals{background:white;border:1px solid #e2e8f0;border-radius:8px;padding:8px 12px;margin-bottom:10px;font-size:12px;color:#475569}.totals strong{color:#1e3a5f;font-size:14px}' +
    '.btn-add{width:100%;padding:9px;background:white;color:#0d9488;border:2px solid #0d9488;border-radius:7px;font-weight:700;font-size:12px;cursor:pointer;margin-bottom:8px}' +
    '.btn-save{width:100%;padding:11px;background:#0d9488;color:white;border:none;border-radius:7px;font-weight:700;font-size:13px;cursor:pointer}.btn-save:disabled{opacity:.45;cursor:not-allowed}' +
    '#msg{margin-top:8px;padding:8px 12px;border-radius:6px;font-size:12px;font-weight:600;display:none}#msg.err{background:#fee2e2;color:#dc2626;display:block}' +
    '</style>' +
    '<div class="header"><h3>⚡ Quick Add Multiple Products</h3><p>Invoice: <strong>' + inv + '</strong> — vehicles will be reserved immediately</p></div>' +
    '<table><thead><tr><th style="width:42%">Model Name</th><th style="width:18%">Qty</th><th style="width:28%">Price (USD)</th><th style="width:12%"></th></tr></thead>' +
    '<tbody id="rows"><tr>' +
    '<td><input type="text" class="model" placeholder="e.g. PULSAR NS200" oninput="recalc()"></td>' +
    '<td><input type="number" class="qty" value="1" min="1" oninput="recalc()"></td>' +
    '<td><input type="number" class="price" value="1250" min="0" step="0.01" oninput="recalc()"></td>' +
    '<td><button class="del-btn" onclick="delRow(this)">✕</button></td>' +
    '</tr></tbody></table>' +
    '<div class="totals" id="totals">Total: <strong>0 vehicles</strong> &nbsp;·&nbsp; USD <strong>0.00</strong></div>' +
    '<button class="btn-add" onclick="addRow()">＋ Add Another Product</button>' +
    '<button class="btn-save" id="saveBtn" onclick="saveProducts()">Save &amp; Assign to Invoice</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'function addRow(){var tbody=document.getElementById("rows"),tr=document.createElement("tr");' +
    'tr.innerHTML=\'<td><input type="text" class="model" placeholder="e.g. AVENGER 220" oninput="recalc()"></td>\'' +
    '+\'<td><input type="number" class="qty" value="1" min="1" oninput="recalc()"></td>\'' +
    '+\'<td><input type="number" class="price" value="1250" min="0" step="0.01" oninput="recalc()"></td>\'' +
    '+\'<td><button class="del-btn" onclick="delRow(this)">✕</button></td>\';' +
    'tbody.appendChild(tr);tr.querySelector(".model").focus();recalc();}' +
    'function delRow(btn){var rows=document.querySelectorAll("#rows tr");if(rows.length===1){alert("At least one product row is required.");return;}btn.closest("tr").remove();recalc();}' +
    'function recalc(){' +
    'var qtys=[].slice.call(document.querySelectorAll(".qty")).map(function(i){return parseInt(i.value)||0;});' +
    'var prices=[].slice.call(document.querySelectorAll(".price")).map(function(i){return parseFloat(i.value)||0;});' +
    'var tv=qtys.reduce(function(a,b){return a+b;},0);' +
    'var val=qtys.reduce(function(s,q,i){return s+q*prices[i];},0);' +
    'document.getElementById("totals").innerHTML="Total: <strong>"+tv+" vehicle(s)</strong> &nbsp;·&nbsp; USD <strong>"+val.toFixed(2)+"</strong>";}' +
    'function saveProducts(){' +
    'var models=[].slice.call(document.querySelectorAll(".model")).map(function(i){return i.value.trim();});' +
    'var qtys=[].slice.call(document.querySelectorAll(".qty")).map(function(i){return parseInt(i.value)||0;});' +
    'var prices=[].slice.call(document.querySelectorAll(".price")).map(function(i){return parseFloat(i.value)||0;});' +
    'var msg=document.getElementById("msg"),btn=document.getElementById("saveBtn");' +
    'if(models.some(function(m){return !m;})){msg.textContent="❌ All model name fields must be filled.";msg.className="err";return;}' +
    'if(qtys.some(function(q){return q<=0;})){msg.textContent="❌ All quantities must be greater than 0.";msg.className="err";return;}' +
    'if(prices.some(function(p){return p<=0;})){msg.textContent="❌ All prices must be greater than 0.";msg.className="err";return;}' +
    'msg.className="";msg.style.display="none";btn.disabled=true;btn.textContent="Saving…";' +
    'google.script.run' +
    '.withSuccessHandler(function(r){btn.textContent="Save & Assign to Invoice";btn.disabled=false;alert("✅ "+r);google.script.host.close();})' +
    '.withFailureHandler(function(e){btn.textContent="Save & Assign to Invoice";btn.disabled=false;msg.textContent="❌ Error: "+e.message;msg.className="err";})' +
    '.quickAddProducts(models,qtys,prices);}' +
    '<\/script>'
  ).setWidth(540).setHeight(480);

  SpreadsheetApp.getUi().showModalDialog(html, 'Quick Add Multiple Products — ' + inv);
}

function quickAddProducts(models, qtys, prices) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl  = ss.getSheetByName('CONTROL');
  const stock = ss.getSheetByName('Stock');
  const inv   = ctrl.getRange(CFG.invoiceNoCell).getValue();
  if (!inv) throw new Error('No invoice number found in C8.');

  let totalAdded = 0;
  for (var i = 0; i < models.length; i++) {
    if (!models[i] || qtys[i] <= 0) continue;
    const startRow = stock.getLastRow() + 1;
    for (var j = 0; j < qtys[i]; j++) {
      const row = startRow + j;
      stock.getRange(row, 1).setValue('MD2' + Utilities.formatString('%07d', Math.floor(Math.random() * 9999999)));
      stock.getRange(row, 3).setValue(models[i]);
      stock.getRange(row, 7).setValue('RESERVED');
      stock.getRange(row, 8).setValue(inv);
      stock.getRange(row, 10).setValue(prices[i]);
      totalAdded++;
    }
  }
  logAudit(inv, 'QUICK_MULTI_PRODUCT', totalAdded + ' vehicles added across ' + models.length + ' model(s)');
  updateChassisDropdown();
  return totalAdded + ' vehicle(s) added and reserved under invoice ' + inv;
}

function setupPortDropdowns() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName('CONTROL');
  const ui   = SpreadsheetApp.getUi();

  const PORTS = ['KARACHI, PAKISTAN','CHITTAGONG, BANGLADESH','COLOMBO, SRI LANKA','DUBAI, UAE',
    'MOMBASA, KENYA','DAR ES SALAAM, TANZANIA','DJIBOUTI','LAGOS, NIGERIA','TEMA, GHANA',
    'PORT LOUIS, MAURITIUS','ANTANANARIVO, MADAGASCAR','LUSAKA, ZAMBIA','HARARE, ZIMBABWE',
    'KAMPALA, UGANDA','NAIROBI, KENYA'];

  const DESTINATIONS = ['KARACHI, PAKISTAN','DHAKA, BANGLADESH','COLOMBO, SRI LANKA','DUBAI, UAE',
    'MOMBASA, KENYA','DAR ES SALAAM, TANZANIA','DJIBOUTI','LAGOS, NIGERIA','ACCRA, GHANA',
    'PORT LOUIS, MAURITIUS','LUSAKA, ZAMBIA','HARARE, ZIMBABWE','KAMPALA, UGANDA','NAIROBI, KENYA'];

  ctrl.getRange('F21').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PORTS, true).setAllowInvalid(true)
      .setHelpText('Select a port or type a custom value').build());
  ctrl.getRange('H22').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(DESTINATIONS, true).setAllowInvalid(true)
      .setHelpText('Select a destination or type a custom value').build());

  ui.alert('✅ Dropdowns Installed',
    'Port of Discharge (F21) and Final Destination (H22) now have dropdown lists.\n\n' +
    'You can still type any custom value — the dropdown is a suggestion, not a lock.',
    ui.ButtonSet.OK);
}
function setupCustomerDropdowns() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('Stock');
  const custSheet  = ss.getSheetByName('Customers');
  const ui         = SpreadsheetApp.getUi();

  if (!stockSheet || !custSheet) {
    ui.alert('❌ Sheet not found', 'Could not find Stock or Customers sheet.', ui.ButtonSet.OK);
    return;
  }

  // Collect smart_dropdown values from Customers (col L = index 11)
  const custData   = custSheet.getRange('A2:L2000').getValues();
  const dropValues = [];
  for (var i = 0; i < custData.length; i++) {
    var val = String(custData[i][11]).trim();  // col L = smart_dropdown
    if (val) dropValues.push(val);
  }

  if (dropValues.length === 0) {
    ui.alert('⚠ No Customers', 'No smart_dropdown values found in Customers sheet (col L).', ui.ButtonSet.OK);
    return;
  }

  // Set headers in Q1 and R1 if empty
  if (!stockSheet.getRange(1, 17).getValue()) stockSheet.getRange(1, 17).setValue('customer_name');
  if (!stockSheet.getRange(1, 18).getValue()) stockSheet.getRange(1, 18).setValue('company_name');

  // Apply dropdown to Q column for all existing stock rows
  const lastRow = Math.max(stockSheet.getLastRow(), 4);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(dropValues, true)
    .setAllowInvalid(true)
    .setHelpText('Pick customer — name and company auto-fill into Q and R')
    .build();

  stockSheet.getRange(4, 17, lastRow - 3, 1).setDataValidation(rule);

  ui.alert(
    '✅ Customer Dropdowns Ready',
    dropValues.length + ' customer(s) loaded into column Q.\n\n' +
    'Picking from the dropdown auto-fills:\n• Q = Contact Name\n• R = Company Name',
    ui.ButtonSet.OK
  );
}

function showAssignedVehiclesPanel() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl     = ss.getSheetByName('CONTROL');
  const stock    = ss.getSheetByName('Stock');
  const ui       = SpreadsheetApp.getUi();
  const invoice  = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const customer = ctrl.getRange(CFG.customerCell).getValue();

  if (!invoice) { ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK); return; }

  const allData  = stock.getRange('A4:J2000').getValues();
  const assigned = allData.filter(function(r) { return r[0] && r[7] === invoice; });

  if (assigned.length === 0) {
    ui.alert('📋 No Vehicles Assigned',
      'No vehicles are currently assigned to invoice ' + invoice + '.\n\nUse "Assign Vehicles (Sidebar)" or "Bulk Assign by Model" to assign.',
      ui.ButtonSet.OK);
    return;
  }

  const modelSummary = {};
  assigned.forEach(function(r) {
    const model = r[2] || 'Unknown';
    if (!modelSummary[model]) modelSummary[model] = { count: 0, prices: [] };
    modelSummary[model].count++;
    if (r[9]) modelSummary[model].prices.push(Number(r[9]));
  });

  const summaryRows = Object.keys(modelSummary).map(function(model) {
    const info = modelSummary[model];
    const priceDisplay = info.prices.length > 0 ? 'USD ' + info.prices.filter(function(v,i,a){return a.indexOf(v)===i;}).join(' / ') : 'default price';
    return '<tr><td style="font-weight:700;padding:6px 10px">' + model + '</td><td style="padding:6px 10px;text-align:center">' + info.count + '</td><td style="padding:6px 10px;color:#64748b;font-size:11px">' + priceDisplay + '</td></tr>';
  }).join('');

  const vehicleRows = assigned.map(function(r, i) {
    return '<tr style="background:' + (i%2===0?'white':'#f8fafc') + '">' +
      '<td style="font-family:monospace;font-size:11px;padding:5px 8px">' + r[0] + '</td>' +
      '<td style="padding:5px 8px;font-size:12px">' + r[2] + '</td>' +
      '<td style="padding:5px 8px;font-size:11px;color:#64748b">' + (r[3]||'') + '</td>' +
      '<td style="padding:5px 8px;font-size:11px;color:#64748b">' + (r[4]||'') + '</td>' +
      '<td style="padding:5px 8px;font-size:11px;text-align:right;color:#0f766e;font-weight:600">' + (r[9]?'USD '+r[9]:'') + '</td></tr>';
  }).join('');

  ui.showModelessDialog(HtmlService.createHtmlOutput(
    '<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:12px;background:#f8fafc}' +
    '.header{background:#1e3a5f;color:white;padding:10px 14px;border-radius:8px;margin-bottom:10px}.header h3{font-size:14px;font-weight:600}' +
    '.section-title{font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 4px}' +
    'table{width:100%;border-collapse:collapse;background:white;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;margin-bottom:10px}' +
    'thead th{background:#e2e8f0;color:#475569;padding:6px 8px;font-size:11px;font-weight:600;text-align:left}' +
    '.total-bar{background:#0d9488;color:white;padding:8px 14px;border-radius:8px;font-size:13px;font-weight:700;text-align:center}' +
    '</style>' +
    '<div class="header"><h3>📋 Assigned Vehicles</h3><p style="font-size:11px;opacity:.8;margin-top:2px">Invoice: <strong>' + invoice + '</strong> &nbsp;·&nbsp; Buyer: <strong>' + (customer||'not set') + '</strong></p></div>' +
    '<div class="section-title">Summary by Model</div>' +
    '<table><thead><tr><th>Model</th><th style="text-align:center">Qty</th><th>Price</th></tr></thead><tbody>' + summaryRows + '</tbody></table>' +
    '<div class="section-title">All Assigned Chassis</div>' +
    '<table><thead><tr><th>Chassis No.</th><th>Model</th><th>Colour</th><th>Year</th><th style="text-align:right">Unit Price</th></tr></thead><tbody>' + vehicleRows + '</tbody></table>' +
    '<div class="total-bar">Total: ' + assigned.length + ' vehicle(s) assigned to ' + invoice + '</div>'
  ).setTitle('Assigned Vehicles — ' + invoice).setWidth(700).setHeight(520), 'Assigned Vehicles — ' + invoice);
}

function buildMonthlyStockTabs() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stock = ss.getSheetByName('Stock');
  const ui    = SpreadsheetApp.getUi();

  if (!stock) { ui.alert('❌ No Stock Sheet', 'Could not find the Stock sheet.', ui.ButtonSet.OK); return; }

  const data = stock.getRange('A4:J2000').getValues().filter(function(r) { return r[0]; });
  if (data.length === 0) { ui.alert('⚠ No Data', 'Stock sheet has no vehicle rows.', ui.ButtonSet.OK); return; }

  const groups = {};
  data.forEach(function(row) {
    const shipDate = row[8];
    let monthKey;
    if (shipDate && shipDate instanceof Date)
      monthKey = Utilities.formatDate(shipDate, 'GMT+5:30', 'MMM-yyyy');
    else if (shipDate && String(shipDate).trim())
      monthKey = String(shipDate).substring(0, 7);
    else
      monthKey = 'Unshipped';
    if (!groups[monthKey]) groups[monthKey] = [];
    groups[monthKey].push(row);
  });

  let tabsCreated = 0, tabsUpdated = 0;
  Object.keys(groups).forEach(function(monthKey) {
    const rows    = groups[monthKey];
    const tabName = 'SR_' + monthKey;
    let sheet = ss.getSheetByName(tabName);
    if (sheet) { sheet.clearContents(); tabsUpdated++; }
    else { sheet = ss.insertSheet(tabName); tabsCreated++; }

    const headerRange = sheet.getRange(1, 1, 1, 10);
    headerRange.setValues([['Chassis No.','Engine No.','Model','Colour','Year','Status','Reserved For','Invoice','Ship Date','Unit Price USD']]);
    headerRange.setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 10).setValues(rows);
    const summaryRow = rows.length + 3;
    sheet.getRange(summaryRow, 1).setValue('Total Vehicles:');
    sheet.getRange(summaryRow, 2).setValue(rows.length);
    sheet.getRange(summaryRow, 1, 1, 2).setFontWeight('bold');
    sheet.autoResizeColumns(1, 10);
  });

  logAudit('SYSTEM', 'MONTHLY_TABS_BUILT', tabsCreated + ' created, ' + tabsUpdated + ' updated across ' + Object.keys(groups).length + ' months');
  ui.alert('✅ Stock Register Built',
    Object.keys(groups).length + ' monthly tab(s) created/updated:\n\n' +
    Object.keys(groups).map(function(k) { return '• SR_' + k + '  (' + groups[k].length + ' vehicles)'; }).join('\n') +
    '\n\nRun this again anytime to refresh all tabs.',
    ui.ButtonSet.OK);
}

function showExporterBankEditor() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const company = ss.getSheetByName('Company');
  if (!company) {
    SpreadsheetApp.getUi().alert('❌ No Company Sheet', 'Could not find the Company sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const vals = company.getRange('B2:B14').getValues().map(function(r) { return String(r[0] || ''); });
  const labels = ['Company Name','Address','Phone','IEC Code','PAN','GSTIN','Bank Name','Account Number','SWIFT Code','Branch','Signatory Name','CHA Name','Customs Office'];
  const ids    = ['companyName','address','phone','iec','pan','gstin','bankName','accountNo','swift','branch','signatory','cha','customsOffice'];
  const phs    = ['e.g. ABC EXPORTS PVT LTD','Full registered address','+91 XXXXX XXXXX','10-digit IEC','AAAAA0000A','15-digit GSTIN','e.g. HDFC Bank','','e.g. HDFCINBB','Branch name and address','Authorised signatory full name','Customs House Agent name','e.g. ICD TUGHLAKABAD'];
  const sectionAt = {0:'Exporter Details', 6:'Bank Details', 10:'Signatory & CHA'};

  let fieldsHtml = '';
  vals.forEach(function(val, i) {
    if (sectionAt[i]) fieldsHtml += '<div class="section">' + sectionAt[i] + '</div>';
    fieldsHtml += '<div class="field"><label>' + labels[i] + '</label><input type="text" id="' + ids[i] + '" value="' + val.replace(/"/g, '&quot;') + '" placeholder="' + phs[i] + '"></div>';
  });

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(
    '<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:14px;background:#f8fafc}' +
    '.header{background:#1e3a5f;color:white;padding:10px 14px;border-radius:8px;margin-bottom:14px}.header h3{font-size:14px;font-weight:600}.header p{font-size:11px;opacity:.8;margin-top:2px}' +
    '.section{font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.5px;margin:14px 0 6px;border-top:1px solid #e2e8f0;padding-top:10px}' +
    '.field{margin-bottom:8px}label{display:block;font-size:11px;font-weight:600;color:#475569;margin-bottom:3px}' +
    'input{width:100%;padding:7px 10px;border:1px solid #cbd5e1;border-radius:6px;font-size:12px;background:white}input:focus{outline:none;border-color:#0d9488}' +
    '.btn-save{width:100%;padding:11px;background:#0d9488;color:white;border:none;border-radius:7px;font-weight:700;font-size:13px;cursor:pointer;margin-top:14px}' +
    '#msg{margin-top:8px;padding:8px 12px;border-radius:6px;font-size:12px;font-weight:600;display:none}' +
    '#msg.ok{background:#dcfce7;color:#15803d;display:block}#msg.err{background:#fee2e2;color:#dc2626;display:block}' +
    '</style>' +
    '<div class="header"><h3>✏️ Edit Exporter & Bank Details</h3><p>Changes save directly to the Company sheet (B2:B14)</p></div>' +
    fieldsHtml +
    '<button class="btn-save" id="saveBtn" onclick="saveDetails()">Save Changes to Company Sheet</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'var IDS=["companyName","address","phone","iec","pan","gstin","bankName","accountNo","swift","branch","signatory","cha","customsOffice"];' +
    'function saveDetails(){' +
    'var btn=document.getElementById("saveBtn"),msg=document.getElementById("msg");' +
    'var vals=IDS.map(function(id){return document.getElementById(id).value.trim();});' +
    'if(!vals[0]){msg.textContent="❌ Company name cannot be empty.";msg.className="err";return;}' +
    'btn.disabled=true;btn.textContent="Saving…";msg.className="";msg.style.display="none";' +
    'google.script.run' +
    '.withSuccessHandler(function(){msg.textContent="✅ Saved successfully.";msg.className="ok";btn.textContent="Save Changes to Company Sheet";btn.disabled=false;})' +
    '.withFailureHandler(function(e){msg.textContent="❌ Error: "+e.message;msg.className="err";btn.textContent="Save Changes to Company Sheet";btn.disabled=false;})' +
    '.saveExporterBankDetails(vals);}' +
    '<\/script>'
  ).setTitle('Edit Exporter & Bank Details').setWidth(480));
}

function saveExporterBankDetails(vals) {
  const company = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Company');
  if (!company) throw new Error('Company sheet not found.');
  company.getRange('B2:B14').setValues(vals.map(function(v) { return [v]; }));
  logAudit('SYSTEM', 'EXPORTER_UPDATED', 'Company/bank details edited via sidebar');
}

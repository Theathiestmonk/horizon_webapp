const CFG = {
  invoiceNoCell:      'C8',
  dateCell:           'F8',
  customerCell:       'C12',
  buyerAddressCell:   'C13',
  buyerCountryCell:   'H22',
  stockCell:          'C17',
  lcCell:             'C24',
  lcDateCell:         'F24',  // NEW — next to LC No, was an unused FIELD/VALUE slot
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
  stockTabCell:       'C19',  // dropdown: which monthly Stock_YYYY_MM tab this invoice pulls vehicles from
  // NEW — section 10 "CUSTOMS / DESTINATION COMPLIANCE" (row 56-58), for the
  // Commercial Invoice / CHA CI/TI generated description block. All optional
  // — blank means that line is simply omitted from the generated text.
  tinNoCell:              'C57',
  vehicleTypeCell:        'F57',
  dealerCertNoCell:       'C58',
  dealerCertDateCell:     'F58',
  yearOfManufactureCell:  'C59',
  webhookUrl: 'https://hayes-latin-theorem-remedies.trycloudflare.com/api/v1/invoices/',
};

// ── Descriptions tab ────────────────────────────────────────────────────────
// Every generated document has exactly ONE description placeholder in its
// table (verified directly against all 7 templates) — the 9-way split lives
// entirely in the Products tab's desc_* columns, which get auto-appended with
// product_name/model via withModel_() in buildPayload(). Two problems with
// that: (1) most of those columns are usually left blank and silently fall
// back to a generic default, and (2) even a filled-in column never prints
// verbatim — withModel_() always tacks on extra text.
// This tab is a purely additive override: same 9 columns, keyed by
// product_name (must match Products!B exactly, case-insensitive). Whatever
// is typed here prints EXACTLY as typed — no fallback default, no appending.
// If this tab doesn't exist, or a model/column here is blank, buildPayload()
// falls through to the original Products-tab + withModel_() behavior
// unchanged — nothing about existing generation can regress from adding this.
var DESCRIPTIONS_TAB_NAME_ = 'Descriptions';

// ── Monthly Stock tabs ──────────────────────────────────────────────────────
// Stock is split across tabs named Stock_YYYY_MM (one per intake month) instead
// of a single ever-growing "Stock" sheet. Every monthly tab mirrors the exact
// row/column layout of the original "Stock" sheet (title row 1, instructions
// row 2, headers row 3, data row 4+) so none of the existing A4:S2000-style
// ranges elsewhere in this file need to change — only *which sheet* they run
// against changes, via getSelectedStockSheet_().
var STOCK_TAB_REGEX_ = /^Stock_\d{4}_\d{2}$/;

function isStockTabName_(name) {
  return name === 'Stock' || STOCK_TAB_REGEX_.test(name);
}

// Invoice numbers get compared between a Stock cell (r[invCol]) and CONTROL!C8
// in a dozen places. A raw ​=== fails silently on any stray whitespace or a
// numeric vs. string mismatch (e.g. a cell that auto-trimmed on one side but
// not the other) — that's what caused "0 vehicles found" for an invoice that
// visibly had vehicles assigned in the Stock tab. Always compare through this.
function normInvoice_(v) {
  return String(v || '')
    .replace(/[\u200B\u200C\u200D\uFEFF\u00A0]/g, '')  // strip zero-width/NBSP chars (copy-paste artifacts)
    .trim()
    .toUpperCase();
}

function sameInvoice_(a, b) {
  return normInvoice_(a) === normInvoice_(b);
}

function rowMatchesInvoice_(row, invoiceNo) {
  if (!row) return false;
  return sameInvoice_(row[7], invoiceNo) || sameInvoice_(row[18], invoiceNo);
}

function listStockTabNames_(ss) {
  return ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return STOCK_TAB_REGEX_.test(n); })
    .sort();
}

// Resolves the Stock sheet to operate on for the invoice currently loaded in
// CONTROL: whatever is picked in CFG.stockTabCell, falling back to the legacy
// single 'Stock' sheet (pre-migration) or the only monthly tab if just one
// exists. Throws a descriptive Error otherwise — callers decide how to surface it.
function getSelectedStockSheet_(ss, ctrl) {
  var picked = String(ctrl.getRange(CFG.stockTabCell).getValue() || '').trim();
  if (picked) {
    var sheet = ss.getSheetByName(picked);
    if (sheet) return sheet;
    throw new Error('Selected Stock Tab "' + picked + '" (CONTROL!' + CFG.stockTabCell + ') was not found. Pick a valid tab from the dropdown.');
  }
  var legacy = ss.getSheetByName('Stock');
  if (legacy) return legacy;
  var monthly = listStockTabNames_(ss);
  if (monthly.length === 1) return ss.getSheetByName(monthly[0]);
  throw new Error('No Stock Tab selected. Pick a month from the "Stock Tab" dropdown in CONTROL!' + CFG.stockTabCell + '.');
}

// Menu-triggered functions call this instead of duplicating try/catch —
// resolves the stock sheet or shows an alert and returns null.
function resolveStockOrAlert_(ss, ctrl, ui) {
  try {
    return getSelectedStockSheet_(ss, ctrl);
  } catch (err) {
    ui.alert('❌ Stock Tab Not Set', err.message, ui.ButtonSet.OK);
    return null;
  }
}

// Keeps CONTROL!C19's dropdown list (and its label in B19) in sync with
// whatever Stock_YYYY_MM tabs currently exist.
function refreshStockTabDropdown_(ss) {
  var ctrl = ss.getSheetByName('CONTROL');
  if (!ctrl) return;
  var labelCell = ctrl.getRange('B19');
  if (!String(labelCell.getValue()).trim()) labelCell.setValue('Stock Tab (Month)  ●');
  var names = listStockTabNames_(ss);
  var cell = ctrl.getRange(CFG.stockTabCell);
  if (names.length === 0) {
    cell.clearDataValidations();
    cell.setNote('No monthly Stock tabs found yet — use ⚡ Horizon → Create Monthly Stock Tab.');
    return;
  }
  cell.clearNote();
  cell.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(names, true)
      .setAllowInvalid(false)
      .setHelpText('Pick which month\'s Stock tab this invoice should pull vehicles from.')
      .build()
  );
}

// Labels the new section-10 CONTROL cells (TIN, dealer certificate, vehicle
// type/year, LC date) the first time they're needed — only ever sets a label
// if that cell is currently blank, so it can't clobber anything you've
// already typed there. Values themselves are always left for you to fill in.
function ensureCustomsFieldsLabels_(ss) {
  var ctrl = ss.getSheetByName('CONTROL');
  if (!ctrl) return;
  function setIfBlank(cellA1, text) {
    var cell = ctrl.getRange(cellA1);
    if (!String(cell.getValue()).trim()) cell.setValue(text);
  }
  setIfBlank('E24', 'LC Date');
  setIfBlank('A56', '  ⑩  CUSTOMS / DESTINATION COMPLIANCE  (optional — used by Commercial Invoice / CHA CI, TI generation)');
  setIfBlank('B57', 'TIN No.');
  setIfBlank('E57', 'Vehicle Type');
  setIfBlank('B58', 'Dealer Cert No.');
  setIfBlank('E58', 'Dealer Cert Date');
  setIfBlank('B59', 'Year of Manufacture');
}

function parseDateFlexible_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  var m = /^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/.exec(s);  // dd.mm.yyyy / dd/mm/yyyy
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}

// Per-vehicle columns T-W, added for the full chassis/certificate-level
// description generation used by Commercial Invoice / CHA CI (see
// buildVehicleDetailBlock_ / buildChaCiVehicleBlock_ below). All 4 are
// optional per row: district/state origin fall back to the Products-tab
// per-model value when blank, and eic_cert_no / first_registration_date
// simply don't print a line if left blank. Nothing reading only A:S breaks —
// these are pure additions past the existing S column.
var STOCK_EXTRA_COL_START_ = 20;  // col T, 1-based
var STOCK_EXTRA_COL_END_   = 23;  // col W, 1-based
var STOCK_EXTRA_HEADERS_ = [
  'eic_cert_no', 'first_registration_date', 'district_origin_code', 'state_origin_code'
];

// Ensures columns T-W have their header labels in row 3 — safe to call
// repeatedly, only ever touches header cells that are currently blank.
function ensureStockExtraColumns_(sheet) {
  var existing = sheet.getRange(3, STOCK_EXTRA_COL_START_, 1, STOCK_EXTRA_HEADERS_.length).getValues()[0];
  var needsHeaders = existing.some(function(v) { return !String(v || '').trim(); });
  if (!needsHeaders) return;
  var range = sheet.getRange(3, STOCK_EXTRA_COL_START_, 1, STOCK_EXTRA_HEADERS_.length);
  range.setValues([STOCK_EXTRA_HEADERS_]);
  range.setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
}

// Creates (or returns) Stock_<yyyy>_<mm>, copying the header layout (rows 1-3)
// from the legacy 'Stock' sheet or, failing that, the most recent monthly tab —
// so every monthly tab keeps the same structure the rest of the script assumes.
function ensureMonthlyStockTab_(ss, yyyy, mm) {
  var tabName = 'Stock_' + yyyy + '_' + mm;
  var existing = ss.getSheetByName(tabName);
  if (existing) return existing;

  var template = ss.getSheetByName('Stock');
  if (!template) {
    var names = listStockTabNames_(ss);
    if (names.length > 0) template = ss.getSheetByName(names[names.length - 1]);
  }

  var sheet = ss.insertSheet(tabName);
  if (template) {
    var lastCol = Math.max(template.getLastColumn(), STOCK_EXTRA_COL_END_);  // keep at least A:W
    var headerRange = template.getRange(1, 1, 3, lastCol);
    var headerRows = headerRange.getValues();
    headerRows[0][0] = 'VEHICLE STOCK REGISTER — ' + Utilities.formatDate(new Date(Number(yyyy), Number(mm) - 1, 1), 'GMT+5:30', 'MMMM yyyy');
    // Copy the template's own header formatting (fonts, borders, merges, etc.)
    // before overwriting values, so a hand-styled header row isn't reduced to
    // just the three manual style calls below.
    headerRange.copyTo(sheet.getRange(1, 1, 3, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheet.getRange(1, 1, 3, lastCol).setValues(headerRows);
    sheet.getRange(3, 1, 1, lastCol).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');

    // Copy formulas, cell formatting (number formats, colors, borders),
    // conditional formatting and data validation (status/customer dropdowns,
    // etc.) from the template's data rows — copying values alone left new
    // tabs looking right but missing any formula-driven columns (e.g.
    // model_color / assign_display / Smart Fill) and cell-level formatting
    // the original Stock sheet had set up on its data rows. PASTE_FORMULA
    // only touches cells that actually contain a formula, so this never
    // duplicates the template's literal vehicle data into the new tab.
    var dataRowCount = 1997;  // matches the A4:S2000 range used everywhere else
    var srcData  = template.getRange(4, 1, dataRowCount, lastCol);
    var destData = sheet.getRange(4, 1, dataRowCount, lastCol);
    srcData.copyTo(destData, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    srcData.copyTo(destData, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    srcData.copyTo(destData, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    srcData.copyTo(destData, SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false);

    // Column widths are a per-column (not per-cell) attribute — copy them
    // explicitly since none of the range-based copyTo calls above carry them.
    for (var col = 1; col <= lastCol; col++) {
      sheet.setColumnWidth(col, template.getColumnWidth(col));
    }
  } else {
    sheet.getRange(1, 1).setValue('VEHICLE STOCK REGISTER — ' + yyyy + '-' + mm);
    sheet.getRange(3, 1, 1, 16).setValues([['chassis_no','engine_no','model','color','date_received','supplier','status','assigned_to','visibility_filter','Selling price','purchase_inr','remarks','visible_for_assign','model_color','assign_display','Smart Fill']]);
    sheet.getRange(3, 1, 1, 16).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  }
  // Ensure the 4 per-vehicle columns T-W exist even if the template predates
  // them (e.g. old 'Stock' sheet or an older monthly tab) — see
  // STOCK_EXTRA_HEADERS_ note above. Safe to call on every creation; only
  // touches header row 3, never data rows.
  ensureStockExtraColumns_(sheet);
  sheet.setFrozenRows(3);

  // Apply the customer picker dropdown (col Q) directly too, in case the
  // template sheet never had it set up yet (copyTo above only carries over
  // validations that already existed on the template).
  var built = buildCustomerDropdownRule_(ss.getSheetByName('Customers'));
  if (built) applyCustomerDropdownToSheet_(sheet, built.rule);

  return sheet;
}

// Column layout for the 'Descriptions' tab — same header labels as the
// matching Products tab columns (D, E, O, P, Q, R, S, T, U, V) so anyone
// already familiar with Products immediately recognizes them here.
// Columns 12-14 (engine_cc, make, accessories) feed the per-vehicle CI /
// CHA CI block generator (buildVehicleDetailBlock_ / buildChaCiVehicleBlock_)
// — they are NOT picked verbatim like the desc_* columns above; the
// generator composes them into the exact multi-line block your customs
// paperwork needs, since that block includes per-vehicle chassis/certificate
// data no single static cell could hold.
var DESCRIPTIONS_HEADERS_ = [
  'product_id', 'product_name',
  'desc_commercial_invoice', 'desc_scomet', 'desc_packing_list', 'desc_tax_invoice',
  'desc_PI', 'desc_annexure1', 'PI HSN Code',
  'CHA TI Description', 'CHA PL Description', 'CHA CI Description',
  'engine_cc', 'make', 'accessories'
];

// Creates (or returns) the 'Descriptions' tab if it doesn't already exist.
// Safe to call repeatedly — never touches Products or any other sheet.
function ensureDescriptionsTab_(ss) {
  var sheet = ss.getSheetByName(DESCRIPTIONS_TAB_NAME_);
  if (sheet) return sheet;

  sheet = ss.insertSheet(DESCRIPTIONS_TAB_NAME_);
  var lastCol = DESCRIPTIONS_HEADERS_.length;

  sheet.getRange(1, 1).setValue(
    'ITEM DESCRIPTIONS — exact text per document type, used as-is (no auto-appending)');
  sheet.getRange(2, 1).setValue(
    'Fill only what you want to override. product_name must match Products!B exactly (case-insensitive). ' +
    'Leave a cell blank to keep using the Products tab default for that document.');
  sheet.getRange(3, 1, 1, lastCol).setValues([DESCRIPTIONS_HEADERS_]);
  sheet.getRange(3, 1, 1, lastCol).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  sheet.setFrozenRows(3);
  sheet.autoResizeColumns(1, lastCol);

  // Mirror the product_name dropdown from Products (col B), if available, so
  // product_name here can only ever be picked from real Products entries —
  // this is what keeps the modelKey match in buildPayload() reliable.
  var productSheet = ss.getSheetByName('Products');
  if (productSheet) {
    var names = productSheet.getRange('B4:B2000').getValues()
      .map(function(r) { return String(r[0] || '').trim(); })
      .filter(Boolean);
    if (names.length > 0) {
      sheet.getRange(4, 2, 1997, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(names, true)
          .setAllowInvalid(true)
          .setHelpText('Pick a product_name from the Products tab — must match exactly for the override to apply.')
          .build()
      );
    }
  }

  return sheet;
}

// Menu-triggered: create the tab (if missing) and confirm.
function setupDescriptionsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existed = !!ss.getSheetByName(DESCRIPTIONS_TAB_NAME_);
  var sheet = ensureDescriptionsTab_(ss);
  logAudit('SYSTEM', 'DESCRIPTIONS_TAB_SETUP', existed ? 'Already existed' : 'Created');
  SpreadsheetApp.getUi().alert(
    existed ? '✅ Descriptions Tab Already Exists' : '✅ Descriptions Tab Created',
    'Sheet "' + sheet.getName() + '" is ready.\n\n' +
    'Fill in product_name (col B, dropdown matches Products) plus whichever document description columns ' +
    'you want to override with exact text. Leave the rest blank — those documents keep using the Products tab as before.\n\n' +
    'Nothing here affects existing generation until you actually fill in a row.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Menu-triggered: seed two sample rows (matching PROD001/PROD002 in the
// reference Products sheet) with distinct, realistic text per document type,
// purely so the override path can be tested end-to-end before real data goes
// in. Never touches Products — only adds/overwrites rows in Descriptions.
function seedDescriptionsDummyData() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert('🧪 Seed Test Description Data',
    'This adds 2 sample rows to the Descriptions tab (for BAJAJ PULSAR NS 250 and BAJAJ AVENGER 160) ' +
    'with distinct made-up text per document, purely for testing that each document shows its own wording.\n\n' +
    'Safe to run — only affects the Descriptions tab. Continue?',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureDescriptionsTab_(ss);

  var dummyRows = [
    [
      'PROD001', 'BAJAJ PULSAR NS 250',
      '[TEST] Commercial Invoice text — BAJAJ PULSAR NS 250 motorcycle',
      '[TEST] SCOMET text — Motorcycle, cylinder capacity >250cc <=500cc',
      '[TEST] Packing List text — 1 motorcycle per wooden crate, PULSAR NS 250',
      '[TEST] Tax Invoice text — Motor Cycle (PULSAR NS 250) as per HSN 87112019',
      '[TEST] PI FORMAT text — Motorcycle PULSAR NS 250 for export',
      '[TEST] Annexure-1 text — PULSAR NS 250 unit declaration',
      '87112019',
      '[TEST] CHA Tax Invoice text — PULSAR NS 250',
      '[TEST] CHA Packing List text — PULSAR NS 250',
      '[TEST] CHA Commercial Invoice text — PULSAR NS 250'
    ],
    [
      'PROD002', 'BAJAJ AVENGER 160',
      '[TEST] Commercial Invoice text — BAJAJ AVENGER 160 motorcycle',
      '[TEST] SCOMET text — Motorcycle, cylinder capacity >75cc <=250cc',
      '[TEST] Packing List text — 1 motorcycle per wooden crate, AVENGER 160',
      '[TEST] Tax Invoice text — Motor Cycle (AVENGER 160) as per HSN 87112019',
      '[TEST] PI FORMAT text — Motorcycle AVENGER 160 for export',
      '[TEST] Annexure-1 text — AVENGER 160 unit declaration',
      '87112019',
      '[TEST] CHA Tax Invoice text — AVENGER 160',
      '[TEST] CHA Packing List text — AVENGER 160',
      '[TEST] CHA Commercial Invoice text — AVENGER 160'
    ]
  ];

  var existing = sheet.getRange(4, 1, 1997, 2).getValues();
  dummyRows.forEach(function(newRow) {
    // Reuse the row if this model_name already has one (re-running the seed
    // just refreshes the test text), otherwise take the first blank row.
    var targetRow = -1;
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][1]).trim().toUpperCase() === newRow[1].toUpperCase()) { targetRow = i + 4; break; }
    }
    if (targetRow === -1) {
      for (var j = 0; j < existing.length; j++) {
        if (!existing[j][0] && !existing[j][1]) { targetRow = j + 4; break; }
      }
    }
    if (targetRow === -1) targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
  });

  logAudit('SYSTEM', 'DESCRIPTIONS_DUMMY_SEEDED', '2 test rows added/updated');
  ui.alert('✅ Test Data Added',
    '2 sample rows added to the Descriptions tab.\n\n' +
    'To test: assign a PULSAR NS 250 or AVENGER 160 vehicle to an invoice and generate documents — ' +
    'each document should now show its own "[TEST] ..." wording instead of the Products-tab default.\n\n' +
    'Delete these rows (or clear the cells) once you\'ve confirmed it works and are ready to enter real text.',
    ui.ButtonSet.OK);
}

// ── Invoice_Descriptions tab ────────────────────────────────────────────────
// Simplest possible override, one row per shipment (not per model): type the
// complete, exact final text for each document directly, once. No per-vehicle
// chassis blocks, no auto-appending, no Stock/CONTROL schema to fill in —
// just literal text in, literal text out.
// Checked FIRST in buildPayload(), before the per-model Descriptions tab and
// before the auto-generated CI/CHA CI/CHA TI/CHA PL blocks — if a row exists
// here for the current invoice, it wins outright and everything else (the
// per-vehicle generation, the model-keyed Descriptions tab) is skipped for
// that document. Leave a cell blank to fall through to those instead.
//
// invoice_no vs pi_invoice_no: a shipment's PROFORMA number and its eventual
// FINAL invoice number are often different values typed into the same
// CONTROL!C8 cell at different stages. One row here can serve both stages —
// fill in whichever number(s) apply. When CONTROL's Mode is PROFORMA, the
// lookup matches against pi_invoice_no; for FINAL/DRAFT it matches
// invoice_no. Fill in just one if you only need one stage, or both if the
// same row should answer to either number as the shipment progresses.
var INVOICE_DESC_TAB_NAME_ = 'Invoice_Descriptions';
var INVOICE_DESC_HEADERS_ = [
  'invoice_no', 'pi_invoice_no',
  'desc_commercial_invoice', 'desc_scomet', 'desc_packing_list', 'desc_tax_invoice',
  'desc_PI', 'desc_annexure1',
  'CHA TI Description', 'CHA PL Description', 'CHA CI Description'
];

function ensureInvoiceDescriptionsTab_(ss) {
  var sheet = ss.getSheetByName(INVOICE_DESC_TAB_NAME_);
  if (sheet) return sheet;

  sheet = ss.insertSheet(INVOICE_DESC_TAB_NAME_);
  var lastCol = INVOICE_DESC_HEADERS_.length;

  sheet.getRange(1, 1).setValue(
    'INVOICE DESCRIPTIONS — one row per shipment, exact text per document, used as-is');
  sheet.getRange(2, 1).setValue(
    'Fill in invoice_no (FINAL/DRAFT number) and/or pi_invoice_no (PROFORMA number) exactly as they appear in ' +
    'CONTROL!' + CFG.invoiceNoCell + ' at each stage, then type the complete text you want printed on each ' +
    'document. Leave a description cell blank to fall back to the Descriptions tab (per-model) or Products tab ' +
    'instead. This tab always wins when filled in.');
  sheet.getRange(3, 1, 1, lastCol).setValues([INVOICE_DESC_HEADERS_]);
  sheet.getRange(3, 1, 1, lastCol).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  sheet.setFrozenRows(3);
  sheet.autoResizeColumns(1, lastCol);

  return sheet;
}

// ── Auto-generated description preview (read-only) ──────────────────────────
// Sits to the right of the manual override columns (A:K), starting at col M
// (col L left blank as a visual spacer) — a completely separate block that
// buildPayload() never reads back. Every generation run overwrites it with
// whatever text actually ended up in that invoice's documents (Products
// default, a Descriptions-tab override, the auto-generated VIN/CI block, or
// an Invoice_Descriptions override itself), so you can check what printed
// without opening the .docx. Safe to ignore — nothing here feeds back into
// generation, unlike columns A:K.
var DESC_PREVIEW_COL_START_ = 13;  // col M, 1-based
var DESC_PREVIEW_HEADERS_ = [
  'Generated: Commercial Invoice', 'Generated: SCOMET', 'Generated: Packing List',
  'Generated: Tax Invoice', 'Generated: PI FORMAT', 'Generated: Annexure-1',
  'Generated: CHA TI', 'Generated: CHA PL', 'Generated: CHA CI'
];
var DESC_PREVIEW_FIELDS_ = [
  'description_commercial', 'description_scomet', 'description_packing',
  'description_tax', 'description_pi', 'description_annexure1',
  'description_cha_ti', 'description_cha_pl', 'description_cha_ci'
];

function ensureDescriptionPreviewHeaders_(sheet) {
  var existing = sheet.getRange(3, DESC_PREVIEW_COL_START_, 1, DESC_PREVIEW_HEADERS_.length).getValues()[0];
  var needsHeaders = existing.some(function(v) { return !String(v || '').trim(); });
  if (!needsHeaders) return;
  var range = sheet.getRange(3, DESC_PREVIEW_COL_START_, 1, DESC_PREVIEW_HEADERS_.length);
  range.setValues([DESC_PREVIEW_HEADERS_]);
  range.setFontWeight('bold').setBackground('#64748b').setFontColor('#ffffff');
  sheet.getRange(2, DESC_PREVIEW_COL_START_).setValue(
    'AUTO-GENERATED PREVIEW — read-only, overwritten every time documents are generated. Shows the exact text that printed, whatever its source.');
}

// One invoice can have several model groups (different price/model
// combinations) — Commercial Invoice / Packing List / CHA* print different
// text per model, but this tab is one row per invoice, so multiple models'
// text gets stacked with a divider instead of only showing the first one.
function joinItemsField_(items, field) {
  var multiple = items.length > 1;
  var parts = items.map(function(it, i) {
    var text = String(it[field] || '').trim();
    if (!text) return '';
    return multiple ? ('── Model ' + (i + 1) + ' (' + (it.model_display || '') + ') ──\n' + text) : text;
  }).filter(Boolean);
  return parts.join('\n\n');
}

// Writes the final per-invoice description text into the preview block —
// called at the end of buildPayload(), after every override/auto-generation
// pass has already resolved each item's description_* fields, so this always
// reflects exactly what the generated documents actually contain.
//
// One shipment can be generated at two stages under two different numbers
// (PROFORMA's pi_invoice_no, then FINAL's invoice_no) — matching on only
// whichever number is active THIS run would create a second, duplicate row
// once the shipment moves to its other stage. So this matches (and then
// backfills) on EITHER number, consolidating both stages into one row —
// exactly the "one row can serve both stages" design the override columns
// (A:K) already document above.
function writeDescriptionPreview_(ss, invoiceNo, piInvoiceNo, items) {
  if (!items || items.length === 0) return;
  if (!invoiceNo && !piInvoiceNo) return;
  var sheet = ensureInvoiceDescriptionsTab_(ss);
  ensureDescriptionPreviewHeaders_(sheet);

  var data = sheet.getRange(4, 1, 1997, 2).getValues();
  var targetRow = -1;
  for (var i = 0; i < data.length; i++) {
    var rowInv = data[i][0], rowPi = data[i][1];
    if ((invoiceNo && rowInv && sameInvoice_(rowInv, invoiceNo)) ||
        (piInvoiceNo && rowPi && sameInvoice_(rowPi, piInvoiceNo))) {
      targetRow = i + 4; break;
    }
  }
  if (targetRow === -1) {
    for (var j = 0; j < data.length; j++) {
      if (!data[j][0] && !data[j][1]) { targetRow = j + 4; break; }
    }
    if (targetRow === -1) targetRow = sheet.getLastRow() + 1;
  }

  // Backfill whichever number(s) are known onto this row — never overwrites
  // a value already sitting there (e.g. one you typed manually), and every
  // override column (C:K) stays untouched, so this has zero effect on
  // generation (buildPayload() skips blank override text either way).
  if (invoiceNo && !String(sheet.getRange(targetRow, 1).getValue() || '').trim())
    sheet.getRange(targetRow, 1).setValue(invoiceNo);
  if (piInvoiceNo && !String(sheet.getRange(targetRow, 2).getValue() || '').trim())
    sheet.getRange(targetRow, 2).setValue(piInvoiceNo);

  var rowValues = DESC_PREVIEW_FIELDS_.map(function(f) { return joinItemsField_(items, f); });
  sheet.getRange(targetRow, DESC_PREVIEW_COL_START_, 1, rowValues.length).setValues([rowValues]);
}

function setupInvoiceDescriptionsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existed = !!ss.getSheetByName(INVOICE_DESC_TAB_NAME_);
  var sheet = ensureInvoiceDescriptionsTab_(ss);
  logAudit('SYSTEM', 'INVOICE_DESCRIPTIONS_TAB_SETUP', existed ? 'Already existed' : 'Created');
  SpreadsheetApp.getUi().alert(
    existed ? '✅ Invoice Descriptions Tab Already Exists' : '✅ Invoice Descriptions Tab Created',
    'Sheet "' + sheet.getName() + '" is ready.\n\n' +
    'Add a row per invoice: put the invoice number in column A, then type the exact text for each ' +
    'document in that row\'s columns. That\'s it — no other setup, no other tab, needed.\n\n' +
    'This always takes priority over the Descriptions tab and the auto-generated CI/CHA text when filled in.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createMonthlyStockTab() {
  var html = HtmlService.createHtmlOutput(
    '<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:16px;background:#f8fafc}' +
    'label{display:block;font-weight:600;margin-bottom:6px;color:#1e3a5f}' +
    'input[type=month]{width:100%;padding:8px 10px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px;margin-bottom:12px}' +
    '.btn{width:100%;padding:11px;background:#0d9488;color:white;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer}' +
    '.btn:disabled{opacity:.45;cursor:not-allowed}' +
    '#msg{margin-top:10px;padding:8px 12px;border-radius:6px;font-size:12px;font-weight:600;display:none}' +
    '#msg.success{background:#dcfce7;color:#15803d;display:block}#msg.error{background:#fee2e2;color:#dc2626;display:block}' +
    '</style>' +
    '<label for="m">Pick the stock month to create</label>' +
    '<input type="month" id="m">' +
    '<button class="btn" id="go" onclick="go()">Create Tab</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'function go(){' +
    'var v=document.getElementById("m").value;' +
    'var msg=document.getElementById("msg");' +
    'if(!v){msg.textContent="Pick a month first.";msg.className="error";return;}' +
    'var btn=document.getElementById("go");btn.disabled=true;btn.textContent="Creating…";' +
    'google.script.run' +
    '.withSuccessHandler(function(r){msg.textContent="✅ "+r;msg.className="success";btn.textContent="Create Tab";btn.disabled=false;})' +
    '.withFailureHandler(function(e){msg.textContent="❌ "+e.message;msg.className="error";btn.textContent="Create Tab";btn.disabled=false;})' +
    '.createMonthlyStockTab_serverSide(v);}' +
    '<\/script>'
  ).setWidth(360).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '📅 Create Monthly Stock Tab');
}

function createMonthlyStockTab_serverSide(monthValue) {
  var m = /^(\d{4})-(\d{2})$/.exec(String(monthValue || '').trim());
  if (!m) throw new Error('Invalid month value: ' + monthValue);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureMonthlyStockTab_(ss, m[1], m[2]);
  refreshStockTabDropdown_(ss);
  logAudit('SYSTEM', 'STOCK_TAB_CREATED', sheet.getName());
  return 'Tab "' + sheet.getName() + '" is ready — pick it from CONTROL!' + CFG.stockTabCell + ' when you\'re ready to assign vehicles from it.';
}

function migrateStockToMonthlyTabs() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var stock = ss.getSheetByName('Stock');
  if (!stock) { ui.alert('❌ No Stock Sheet', 'Could not find the Stock sheet — it may already be migrated.', ui.ButtonSet.OK); return; }

  var confirm = ui.alert('🔀 Migrate Stock to Monthly Tabs',
    'This splits every vehicle row in "Stock" into monthly tabs (Stock_YYYY_MM) based on Date Received, ' +
    'then renames "Stock" to a timestamped archive tab (kept as backup, not deleted).\n\n' +
    'This is meant to run once. Continue?',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  var lastCol = Math.max(stock.getLastColumn(), 19);
  var lastRow = stock.getLastRow();
  if (lastRow < 4) { ui.alert('⚠ No Data', 'Stock sheet has no vehicle rows to migrate.', ui.ButtonSet.OK); return; }

  var data = stock.getRange(4, 1, lastRow - 3, lastCol).getValues();
  var groups = {};
  data.forEach(function(row) {
    if (!row[0]) return;  // skip blank rows
    var d = parseDateFlexible_(row[4]);  // col E = date_received
    var key = d ? Utilities.formatDate(d, 'GMT+5:30', 'yyyy_MM') : 'unknown_00';
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  });

  var report = [];
  Object.keys(groups).sort().forEach(function(key) {
    var rows = groups[key];
    var parts = key.split('_');
    var sheet;
    if (parts[0] === 'unknown') {
      sheet = ss.getSheetByName('Stock_Unknown_Date') || ss.insertSheet('Stock_Unknown_Date');
      if (sheet.getLastRow() < 3) {
        var hdr = stock.getRange(1, 1, 3, lastCol).getValues();
        hdr[0][0] = 'VEHICLE STOCK REGISTER — UNKNOWN DATE RECEIVED';
        sheet.getRange(1, 1, 3, lastCol).setValues(hdr);
        sheet.setFrozenRows(3);
        stock.getRange(4, 1, 1997, lastCol)
          .copyTo(sheet.getRange(4, 1, 1997, lastCol), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      }
    } else {
      sheet = ensureMonthlyStockTab_(ss, parts[0], parts[1]);
    }
    if (sheet.getLastRow() > 3) sheet.getRange(4, 1, sheet.getLastRow() - 3, lastCol).clearContent();
    sheet.getRange(4, 1, rows.length, lastCol).setValues(rows);
    report.push('• ' + sheet.getName() + ': ' + rows.length + ' vehicle(s)');
  });

  var archiveName = 'Stock_ARCHIVE_' + Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyyMMdd_HHmmss');
  stock.setName(archiveName);

  refreshStockTabDropdown_(ss);
  logAudit('SYSTEM', 'STOCK_MIGRATED', Object.keys(groups).length + ' month(s), original archived as ' + archiveName);
  ui.alert('✅ Migration Complete',
    report.join('\n') + '\n\nOriginal sheet kept as backup: "' + archiveName + '"\n\n' +
    'Next: open CONTROL and pick a Stock Tab from the new dropdown (row 19) before assigning vehicles.',
    ui.ButtonSet.OK);
}

// wordsOnly=false (default) keeps the full "AMOUNT CHARGEABLE IN ... ONLY"
// sentence — used when this is called directly as a spreadsheet formula.
// wordsOnly=true returns just the currency + words + "ONLY" part, for
// buildPayload() below, since every docx template already has its own
// static "AMOUNT CHARGEABLE..." label before the {{ amount_*_words }} placeholder
// — using the full sentence there duplicated "AMOUNT CHARGEABLE IN" twice.
function AMOUNTWORDS(n, curr, wordsOnly) {
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
  let result = (wordsOnly ? '' : 'AMOUNT CHARGEABLE IN ') + currText + ' ' + words(whole) + ' ONLY';
  if (cents > 0) result += ' AND PAISE ' + words(cents) + ' ONLY';
  return result;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Horizon')
    .addItem('📄 Generate All Documents',    'generateDocuments')
    .addItem('📋 Generate PI Only',          'generatePIDocument')
    .addItem('🏦 Generate CHA Documents',    'generateCHADocuments')
    .addItem('📑 Generate Annexure C Only',  'generateAnnexureCDocument')
    .addItem('✅ Validate Shipment',         'validateAndReport')
    .addSeparator()
    .addItem('🚗 Assign Vehicles (Sidebar)', 'showVehicleSidebar')
    .addItem('🎯 Select Vehicles for Generation (C17)', 'showGenerationVehicleSelector')
    .addItem('🔍 Preview Descriptions (no documents)', 'generateDescriptionPreview')
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
    .addItem('📊 Legacy Ship-Date Report Tabs (SR_, read-only)', 'buildMonthlyStockTabs')
    .addItem('✏️ Edit Exporter / Bank',        'showExporterBankEditor')
    .addSeparator()
    .addItem('🆕 Create Monthly Stock Tab (Stock_YYYY_MM)',    'createMonthlyStockTab')
    .addItem('🔀 Migrate Stock → Monthly Tabs (one-time, do this first)', 'migrateStockToMonthlyTabs')
    .addSeparator()
    .addItem('🔄 Refresh Chassis Dropdown',    'updateChassisDropdown')
    .addItem('👤 Set Customer Dropdowns (Stock Q/R)', 'setupCustomerDropdowns')
    .addSeparator()
    .addItem('📋 Setup Invoice Descriptions Tab (simplest — one row per invoice)', 'setupInvoiceDescriptionsTab')
    .addItem('📝 Setup Descriptions Tab (per-model, advanced)', 'setupDescriptionsTab')
    .addItem('🧪 Seed Test Description Data (for testing only)', 'seedDescriptionsDummyData')
    .addToUi();
  updateChassisDropdown();
  refreshStockTabDropdown_(SpreadsheetApp.getActiveSpreadsheet());
  ensureCustomsFieldsLabels_(SpreadsheetApp.getActiveSpreadsheet());
}

function onEdit(e) {
  const ss    = e.source;
  const sheet = ss.getActiveSheet();
  const name  = sheet.getName();
  const row   = e.range.getRow();
  const col   = e.range.getColumn();
  const cell  = e.range.getA1Notation();

  if (isStockTabName_(name) && col === 8 && row >= 4)
    sheet.getRange(row, 7).setValue(e.range.getValue() ? 'RESERVED' : 'AVAILABLE');
  if (isStockTabName_(name) && col === 9 && row >= 4 && e.range.getValue())
    sheet.getRange(row, 7).setValue('SHIPPED');

  // Refresh C17 chassis dropdown whenever invoice number in C8 changes
  if (name === 'CONTROL' && cell === 'C8')
    updateChassisDropdown();

  // Refresh C17 chassis dropdown whenever the selected Stock Tab (C19) changes
  if (name === 'CONTROL' && cell === CFG.stockTabCell)
    updateChassisDropdown();

  // Customer autofill: selecting from C12 fills address and country in CONTROL
  if (name === 'CONTROL' && cell === 'C12')
    autoFillCustomerDetails_(sheet, ss, e.range.getValue());

  // Chassis autofill: selecting from C17 fills unit price from Stock record
  if (name === 'CONTROL' && cell === 'C17')
    autoFillChassisDetails_(sheet, ss, e.range.getValue());

  // Auto-split customer dropdown in any Stock tab's col Q → contact_name (Q) + company_name (R)
  if (isStockTabName_(name) && col === 17 && row >= 4 && e.value) {
    var val = String(e.value).trim();
    var sep = val.indexOf(' — ');
    if (sep !== -1) {
      sheet.getRange(row, 17).setValue(val.substring(0, sep).trim());
      sheet.getRange(row, 18).setValue(val.substring(sep + 3).trim());
    }
  }
}

// Fires the instant the user clicks into (or arrows onto) a cell — the
// closest thing Apps Script offers to "checkboxes appear when you open the
// dropdown" for C17, since Sheets has no scriptable multi-select dropdown.
// Auto-opens the checkbox sidebar only when C17 is the exact single-cell
// selection AND more than one vehicle is reserved for the current invoice —
// for a single reserved vehicle the plain dropdown from updateChassisDropdown_
// is enough, so this stays silent and doesn't interrupt normal navigation.
function onSelectionChange(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() !== 'CONTROL') return;
  if (range.getA1Notation() !== CFG.stockCell) return;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

  var ss = sheet.getParent();
  var invoiceNo = String(sheet.getRange(CFG.invoiceNoCell).getValue() || '').trim();
  if (!invoiceNo) return;

  var stock;
  try { stock = getSelectedStockSheet_(ss, sheet); } catch (err) { return; }

  var mode   = String(sheet.getRange(CFG.modeCell).getValue() || '').trim();
  var invCol = (mode === 'PROFORMA') ? 18 : 7;
  var reservedCount = stock.getRange('A4:S2000').getValues()
    .filter(function(r) { return r[0] && String(r[6]).trim() === 'RESERVED' && sameInvoice_(r[invCol], invoiceNo); }).length;

  if (reservedCount > 1) showGenerationVehicleSelector();
}

// ── CONTROL autofill helpers ──────────────────────────────────────────────────

// Called when C12 (customer dropdown) is changed in CONTROL.
// Looks up the selected value in Customers col L (smart_dropdown) and fills:
//   C13 = full address  |  C14 = notify party  |  H22 = country
function autoFillCustomerDetails_(ctrl, ss, dropdownVal) {
  if (!dropdownVal) return;
  var custSheet = ss.getSheetByName('Customers');
  if (!custSheet) return;

  var rows = custSheet.getRange('A2:L2000').getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0] && !r[11]) break;                          // empty row = end of data
    if (String(r[11]).trim() !== String(dropdownVal).trim()) continue;

    var companyName = String(r[1]  || '').trim();         // col B = company_name
    var addr1       = String(r[2]  || '').trim();         // col C = address_line1
    var addr2       = String(r[3]  || '').trim();         // col D = address_line2
    var city        = String(r[4]  || '').trim();         // col E = city
    var country     = String(r[5]  || '').trim();         // col F = country
    var phone       = String(r[6]  || '').trim();         // col G = phone (if present)
    var email       = String(r[7]  || '').trim();         // col H = email (if present)
    var contactName = String(r[7]  || '').trim();         // col H = contact_name

    // Full address block for C13
    var fullAddr = [addr1, addr2, city].filter(Boolean).join(', ');

    // Notify party: company name (or contact + company if both present)
    var notifyParty = companyName || contactName;

    ctrl.getRange('C12').setValue(companyName || dropdownVal);  // display company name
    if (fullAddr)    ctrl.getRange('C13').setValue(fullAddr);
    if (notifyParty) ctrl.getRange('C14').setValue(notifyParty);
    if (country)     ctrl.getRange('H22').setValue(country);

    Logger.log('✅ Customer autofill: "' + companyName + '" | addr="' + fullAddr + '" | country="' + country + '"');
    return;
  }
  Logger.log('⚠ Customer autofill: no match for "' + dropdownVal + '"');
}

// Called when C17 (chassis dropdown) is changed in CONTROL.
// Dropdown format: "chassis - engine - model - color"
// Looks up that chassis in Stock and fills F26 (unit price) if the stock record has one.
function autoFillChassisDetails_(ctrl, ss, dropdownVal) {
  if (!dropdownVal) return;
  // A comma means C17 holds a multi-vehicle selection from the "Select
  // Vehicles for Generation" sidebar, not a single legacy autofill value —
  // nothing single to autofill from in that case.
  if (String(dropdownVal).indexOf(',') !== -1) return;
  var chassisNo = String(dropdownVal).split(' - ')[0].trim();
  if (!chassisNo) return;

  var stock;
  try { stock = getSelectedStockSheet_(ss, ctrl); } catch (err) { Logger.log('⚠ Chassis autofill: ' + err.message); return; }

  var rows = stock.getRange('A4:J2000').getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    if (String(r[0]).trim() !== chassisNo) continue;

    var unitPrice = Number(r[9]) || 0;    // col J = unit_price_usd
    var model     = String(r[2] || '').trim();
    var netWt     = Number(r[5] || 0);    // col F = net weight (if stored)
    var grossWt   = Number(r[6] || 0);    // col G = status (not weight — skip)

    if (unitPrice > 0) ctrl.getRange('F26').setValue(unitPrice);
    Logger.log('✅ Chassis autofill: chassis="' + chassisNo + '" | model="' + model + '" | price=' + unitPrice);
    return;
  }
  Logger.log('⚠ Chassis autofill: no Stock row found for chassis "' + chassisNo + '"');
}

// Populates C17 with chassis entries where Stock col G = 'RESERVED' AND col H = current invoice
// Format: "chassis - engine - model - color" to match the existing dropdown display style
// ── ⑨ MULTI-PRODUCT SHIPMENT ITEMS table (CONTROL!A46:B53) ──────────────────
// This table already has live formulas in columns C-H (UNIT PRICE via
// VLOOKUP into Products, FOB SUBTOTAL, CI DESCRIPTION, VINs LINKED via
// COUNTIFS, STATUS) that key off column A's "PROD_ID — MODEL NAME" smart-
// dropdown format (Products!N) and column B's quantity — only A and B are
// meant to be typed/filled in, never the formula columns. This is exactly
// why a single flat Quantity/Unit Price in section ⑤ falls apart once a
// shipment has more than one model or price: this table is where the real
// per-model breakdown belongs, and each row's formulas already resolve its
// own price/description/status independently. Fills one row per distinct
// model among the vehicles actually assigned (or, if a "Select Vehicles for
// Generation" subset is active via C17, just that subset) — never touches
// columns C-H so their formulas keep working untouched.
var MULTI_ITEMS_FIRST_ROW_ = 46;
var MULTI_ITEMS_LAST_ROW_  = 53;

// Section ⑤'s FOB Total (C27) originally read "=IFERROR(C26*F26,0)" — one
// quantity times one price, which only ever worked for a single-model
// shipment. Once vehicles can span different models/prices, that formula
// breaks (blank price × qty = 0, silently zeroing CIF Total). Re-pointing it
// at the item table's own per-row FOB Subtotal column (which already
// multiplies each row's qty by its own VLOOKUP price) keeps FOB/CIF Total
// correct for any mix. Quantity (C26) is deliberately NOT touched here — it's
// the manually-typed TARGET quantity that bulkAssignByModel()/validate()/
// cloneLastShipment() compare actual assignments against ("remaining = qty -
// alreadyAssigned"); turning it into a readout of already-assigned vehicles
// would make it impossible to ever set a target before assigning anything.
// Idempotent — safe to call every refresh; setFormula() is a no-op in
// effect if the formula's unchanged.
function ensureFinancialsFormulas_(ctrl) {
  ctrl.getRange('C27').setFormula(
    '=IFERROR(SUM(D' + MULTI_ITEMS_FIRST_ROW_ + ':D' + MULTI_ITEMS_LAST_ROW_ + '),0)');
}

function refreshMultiProductItemsTable_(ss, ctrl) {
  ensureFinancialsFormulas_(ctrl);
  var capacity  = MULTI_ITEMS_LAST_ROW_ - MULTI_ITEMS_FIRST_ROW_ + 1;
  var dataRange = ctrl.getRange(MULTI_ITEMS_FIRST_ROW_, 1, capacity, 2);  // A:B only — never touch C:H formulas

  var inv = String(ctrl.getRange(CFG.invoiceNoCell).getValue() || '').trim();
  if (!inv) { dataRange.clearContent(); return; }

  var stock;
  try { stock = getSelectedStockSheet_(ss, ctrl); } catch (err) { dataRange.clearContent(); return; }

  var mode   = String(ctrl.getRange(CFG.modeCell).getValue() || '').trim();
  var invCol = (mode === 'PROFORMA') ? 18 : 7;
  var reservedRows = stock.getRange('A4:S2000').getValues()
    .filter(function(r) { return r[0] && String(r[6]).trim() === 'RESERVED' && sameInvoice_(r[invCol], inv); });

  // Same C17 multi-select filter as buildPayload() — only the checked
  // subset counts when a "Select Vehicles for Generation" selection is active.
  var c17Raw = String(ctrl.getRange(CFG.stockCell).getValue() || '').trim();
  var c17SelectionActive = c17Raw.indexOf(',') !== -1;
  if (c17SelectionActive) {
    var selectedSet = {};
    c17Raw.split(',').map(function(s) { return s.trim().toUpperCase(); }).filter(Boolean)
      .forEach(function(c) { selectedSet[c] = true; });
    var filtered = reservedRows.filter(function(r) { return selectedSet[String(r[0]).trim().toUpperCase()]; });
    if (filtered.length > 0) reservedRows = filtered;
  }

  // Quantity (C26) and Total Packages (C34) only follow the selection when
  // the C17 checkbox picker is actually in use — at that point assignment
  // is already finished, so this can't collide with bulkAssignByModel()'s
  // "manually-typed target" use of C26 (that workflow runs before C17 ever
  // holds a comma-separated value), and it can't clobber a genuine manual
  // packing count either (e.g. multiple vehicles per crate) outside this
  // flow. buildPayload()'s total_packages is "Number(C34) || vehicles.length"
  // — since C34 already holds a real (non-zero) number from the full
  // invoice, that fallback never fires on its own, which is exactly why this
  // needs to be set explicitly instead of left to buildPayload()'s fallback.
  if (c17SelectionActive) {
    ctrl.getRange(CFG.qtyCell).setValue(reservedRows.length);
    ctrl.getRange(CFG.totalPackagesCell).setValue(reservedRows.length);
  }

  // VINs Assigned (F34) has no "manually-typed target" role anywhere else
  // in the script (unlike C26) — nothing reads or writes it except H34's
  // own "VIN match" comparison against C26. So unlike C26, this can always
  // track the real count: every reserved vehicle normally, or just the C17-
  // selected subset when that's active — keeping H34 accurate in both cases
  // instead of comparing a stale full-invoice count against a narrowed C26.
  ctrl.getRange('F34').setValue(reservedRows.length);

  if (reservedRows.length === 0) { dataRange.clearContent(); return; }

  // product_id lookup so column A can hold "PROD_ID — MODEL NAME" —
  // matches Products!N's own smart_dropdown formula (=A&" — "&B) exactly,
  // which is what column C-H's VLOOKUP/LEFT/FIND formulas expect.
  var productSheet = ss.getSheetByName('Products');
  var idByModel = {};
  if (productSheet) {
    productSheet.getRange('A4:B2000').getValues().forEach(function(row) {
      var name = String(row[1] || '').trim();
      if (name) idByModel[name.toUpperCase()] = String(row[0] || '').trim();
    });
  }

  var qtyByModel = {};
  var order = [];
  reservedRows.forEach(function(r) {
    var model = String(r[2] || '').trim();
    if (!model) return;
    if (!(model in qtyByModel)) { qtyByModel[model] = 0; order.push(model); }
    qtyByModel[model]++;
  });

  if (order.length > capacity) {
    Logger.log('⚠ MULTI-PRODUCT ITEMS: ' + order.length + ' distinct model(s) found but the table only has ' +
      capacity + ' row(s) — showing the first ' + capacity + '.');
    order = order.slice(0, capacity);
  }

  var rows = order.map(function(model) {
    var pid = idByModel[model.toUpperCase()] || '';
    return [pid ? (pid + ' — ' + model) : model, qtyByModel[model]];
  });
  while (rows.length < capacity) rows.push(['', '']);
  dataRange.setValues(rows);
}

function updateChassisDropdown() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var control = ss.getSheetByName('CONTROL');
  if (!control) return;
  refreshMultiProductItemsTable_(ss, control);
  var stock;
  try { stock = getSelectedStockSheet_(ss, control); } catch (err) {
    control.getRange('C17').clearDataValidations().setNote(err.message);
    return;
  }

  var invoiceNo    = String(control.getRange('C8').getValue()).trim();
  var mode_        = String(control.getRange(CFG.modeCell).getValue()).trim();
  var invCol_      = (mode_ === 'PROFORMA') ? 18 : 7;  // 0-based: col S(18) for PROFORMA, col H(7) for FINAL/DRAFT
  var dropdownCell = control.getRange('C17');

  if (!invoiceNo) {
    dropdownCell.clearDataValidations();
    return;
  }

  var reservedRows = stock.getRange('A4:S2000').getValues()
    .filter(function(r) { return r[0] && String(r[6]).trim() === 'RESERVED' && sameInvoice_(r[invCol_], invoiceNo); });

  if (reservedRows.length === 0) {
    dropdownCell.clearDataValidations();
    dropdownCell.setNote('No RESERVED vehicles found for invoice: ' + invoiceNo);
    return;
  }

  // More than one reserved vehicle means this is exactly the case the
  // "Select Vehicles for Generation" checkbox sidebar exists for. Google
  // Sheets has no scriptable way to turn a plain dropdown into an in-cell
  // multi-select (the "allow multiple selections" chip UI is a manual,
  // per-cell toggle with no Apps Script API — and even if set by hand,
  // rebuilding the validation here would silently strip it back to
  // single-select). So instead of a single-pick list that only ever shows
  // (and risks overwriting) one chassis, clear the dropdown entirely and
  // let onEdit's C17 handler auto-open the real checkbox picker.
  if (reservedRows.length > 1) {
    dropdownCell.clearDataValidations();
    dropdownCell.setNote('Multiple vehicles reserved (' + reservedRows.length + ') — click this cell to pick which ones to include in this generation.');
    return;
  }

  var assignedChassis = reservedRows.map(function(r) {
    var parts = [String(r[0])];
    if (r[1]) parts.push(String(r[1]));
    if (r[2]) parts.push(String(r[2]));
    if (r[3]) parts.push(String(r[3]));
    return parts.join(' - ');
  });

  dropdownCell.clearNote();
  dropdownCell.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(assignedChassis, true)
      .setAllowInvalid(true)
      .setHelpText('Reserved vehicle for invoice ' + invoiceNo)
      .build()
  );
}

function validate() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl  = ss.getSheetByName('CONTROL');
  const errors = [], warnings = [];
  let stock = null;
  try { stock = getSelectedStockSheet_(ss, ctrl); } catch (err) { errors.push('⑥ ' + err.message); }

  const inv      = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const customer = ctrl.getRange(CFG.customerCell).getValue();
  const stockRef = ctrl.getRange(CFG.stockCell).getValue();
  const qty      = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;
  // FOB Total (C27) now sums the item table's own per-model VLOOKUP prices
  // (see ensureFinancialsFormulas_), so it's correct for both single- and
  // multi-price shipments — checking it here instead of the flat F26 Unit
  // Price cell means F26 is no longer a hard requirement for generation.
  const fobTotal = Number(ctrl.getRange('C27').getValue()) || 0;
  const mode     = ctrl.getRange(CFG.modeCell).getValue();
  const lc       = ctrl.getRange(CFG.lcCell).getValue();
  const container = ctrl.getRange(CFG.containerCell).getValue();

  if (!inv)       errors.push('① Invoice number missing          → fix: C8');
  if (!customer)  errors.push('② Customer name missing           → fix: C12');
  if (!stockRef)  errors.push('③ Stock / Chassis Ref missing     → fix: C17');
  if (qty <= 0)   errors.push('④ Quantity must be > 0            → fix: C26');
  if (fobTotal <= 0) errors.push('⑤ FOB Total is zero — check item table has prices → fix: rows 46-53 (via Products catalog) or C27');

  if (!lc)        warnings.push('⚠ LC Number is empty              → cell: C24');
  if (!container) warnings.push('⚠ Container number is empty → cell: C37 (Optional except for Annexure C)');

  if (stock && (mode === 'FINAL' || mode === 'PROFORMA')) {
    let assigned = stock.getRange('A4:S2000').getValues().filter(function(r) { return rowMatchesInvoice_(r, inv); });

    // If a C17 "Select Vehicles for Generation" subset is active, qty (C26)
    // already reflects just that subset (see refreshMultiProductItemsTable_'s
    // c17SelectionActive sync) — compare against the same subset here, not
    // every reserved vehicle, otherwise a deliberate partial-shipment
    // selection always fails validation even though it's fully consistent.
    const stockRefRaw = String(stockRef || '').trim();
    if (stockRefRaw.indexOf(',') !== -1) {
      const selectedSet = {};
      stockRefRaw.split(',').map(function(s) { return s.trim().toUpperCase(); }).filter(Boolean)
        .forEach(function(c) { selectedSet[c] = true; });
      const filtered = assigned.filter(function(r) { return selectedSet[String(r[0]).trim().toUpperCase()]; });
      if (filtered.length > 0) assigned = filtered;
    }

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

// ── Per-vehicle description generators (Commercial Invoice / CHA CI/TI/PL) ──
// These build the multi-line, chassis-level goods description your real CI
// format needs — no static cell could hold this since it must repeat once
// per physical vehicle (chassis, export inspection certificate, first
// registration date), not once per model. Fed by the new Stock columns T-W
// (STOCK_EXTRA_HEADERS_) and the Descriptions tab's engine_cc/make columns.

// Formats an 8-digit HSN as customs-dotted notation (87112019 -> 8711.20.19).
// Anything that isn't exactly 8 digits is returned unchanged — safer than
// guessing a split point for a code we don't recognize.
function formatHsCodeDotted_(raw) {
  var digits = String(raw || '').replace(/[^0-9]/g, '');
  if (digits.length !== 8) return String(raw || '').trim();
  return digits.slice(0, 4) + '.' + digits.slice(4, 6) + '.' + digits.slice(6, 8);
}

// Commercial Invoice style block for one model group:
//   MODEL CC (HS.CODE: x.xx.xx)
//   CHASSIS NO:
//   <chassis1>
//   NO OF THE EXPORT INSPECTION CERTIFICATE: <cert1>
//   YEAR/MONTH OF THE FIRST REGISTRATION: <reg1>
//   <chassis2>                                    <- label NOT repeated
//   ...
//   MAKE:   <make>
//   MODEL:  <model>
// "CHASSIS NO:" appears once per model group, not once per vehicle — matches
// the exact format supplied. EIC-certificate / registration lines are
// omitted per-vehicle if that vehicle's cell is blank, so partially-filled
// data doesn't print empty labels.
function buildVehicleDetailBlock_(modelDisplay, engineCc, hsCodeDotted, make, vehicleGroup) {
  var lines = [];
  var ccPart = engineCc ? (String(engineCc).trim() + ' ') : '';
  lines.push(modelDisplay + ' ' + ccPart + '(HS.CODE: ' + hsCodeDotted + ')');
  lines.push('CHASSIS NO: ');
  vehicleGroup.forEach(function(v) {
    lines.push(v.chassis_no || '');
    if (v.eic_cert_no) lines.push('NO OF THE EXPORT INSPECTION CERTIFICATE: ' + v.eic_cert_no);
    if (v.first_registration_date) lines.push('YEAR/MONTH OF THE FIRST REGISTRATION: ' + v.first_registration_date);
  });
  if (make) lines.push('MAKE:   ' + make);
  lines.push('MODEL:  ' + modelDisplay);
  return lines.join('\n');
}

// Packing List style block for one model group:
//   <base packing description — Products/Descriptions-tab text, model-tagged>
//   VIN 1: <chassis1> ENG. No. <engine1> Colour: <color1>
//   VIN 2: <chassis2> ENG. No. <engine2> Colour: <color2>
//   ...
// One "VIN n:" line per vehicle actually in this model group — vehicleGroup
// is already the post-C17-selection subset, so this reflects only the
// vehicles the user picked for this generation run, not every reserved one.
function buildPackingListVehicleBlock_(baseDesc, vehicleGroup) {
  var lines = [baseDesc];
  vehicleGroup.forEach(function(v, i) {
    lines.push('VIN ' + (i + 1) + ': ' + (v.chassis_no || '') + ' ENG. No. ' + (v.engine_no || '') + ' Colour: ' + (v.color || ''));
  });
  return lines.join('\n');
}

// Shipment-level trailer appended once, after every model block, on the
// Commercial Invoice only: manufacture year/type, accessories, LC/TIN/dealer
// certificate references, proforma-conformity certification line. Every
// piece is optional — a blank CONTROL cell just omits that line rather than
// printing "LC NO.  DT. ".
function buildCommercialInvoiceTrailer_(ctrl, accessories, invoiceNo, invoiceDate, countryOfOrigin, fallbackYear) {
  var lines = [];
  var year = String(ctrl.getRange(CFG.yearOfManufactureCell).getValue() || fallbackYear || '').trim();
  var vehicleType = String(ctrl.getRange(CFG.vehicleTypeCell).getValue() || '').trim();
  var lcNumber = String(ctrl.getRange(CFG.lcCell).getValue() || '').trim();
  var lcDate   = String(ctrl.getRange(CFG.lcDateCell).getValue() || '').trim();
  var tinNo    = String(ctrl.getRange(CFG.tinNoCell).getValue() || '').trim();
  var dealerCertNo   = String(ctrl.getRange(CFG.dealerCertNoCell).getValue() || '').trim();
  var dealerCertDate = String(ctrl.getRange(CFG.dealerCertDateCell).getValue() || '').trim();

  if (year) lines.push('YEAR OF MANUFACTURE: ' + year);
  if (vehicleType) lines.push('TYPE OF VEHICLE: ' + vehicleType);
  if (accessories) {
    lines.push('FOLLOWING ACCESSORIES ARE INSTALLED IN THE VEHICLE:');
    lines.push(accessories);
  }
  if (lcNumber) lines.push('LC NO. ' + lcNumber + (lcDate ? (' DT. ' + lcDate) : ''));
  if (tinNo) lines.push('TIN NO.: ' + tinNo);
  if (dealerCertNo) {
    lines.push('DEPARTMENT OF MOTOR TRAFFIC OF ' + (countryOfOrigin === 'INDIA' ? String(countryOfOrigin) : String(countryOfOrigin || '')) +
      ' DEALER CERTIFICATE NO.' + dealerCertNo + (dealerCertDate ? (' DATED ' + dealerCertDate) : ''));
  }
  if (invoiceNo) {
    lines.push('CERTIFYING THAT SHIPMENT IS IN CONFORMITY WITH PROFORMA');
    lines.push('INVOICE NO. ' + invoiceNo + (invoiceDate ? (' DT ' + invoiceDate) : ''));
  }
  return lines.join('\n');
}

// CHA CI style block for one model group:
//   MOTOR VEHICLES WITH HSN CODE AND CHASSIS NO AS PER LC DOCUMENT NO. x DT. y
//   CHASSIS NO:    <chassis1>          <- label repeated per vehicle here,
//   CHASSIS NO:    <chassis2>             unlike the plain CI block above
//   District Origin Code: <code1, code2, ...>
//   State Origin Code: <code1, code2, ...>
// District/State codes come from each vehicle's own Stock columns (V/W) if
// filled in, else fall back to the model's single Products-tab code —
// matches your example where two units of the same model had two different
// origin codes ("102, 315").
function buildChaCiVehicleBlock_(lcLine, vehicleGroup, defaultDistrict, defaultState) {
  var lines = [lcLine];
  var districts = [], states = [];
  vehicleGroup.forEach(function(v) {
    lines.push('CHASSIS NO:    ' + (v.chassis_no || ''));
    districts.push(v.district_origin_code || defaultDistrict || '');
    states.push(v.state_origin_code || defaultState || '');
  });
  var districtsClean = districts.filter(Boolean);
  var statesClean = states.filter(Boolean);
  if (districtsClean.length) lines.push('District Origin Code: ' + districtsClean.join(', '));
  if (statesClean.length) lines.push('State Origin Code: ' + statesClean.join(', '));
  return lines.join('\n');
}

// Trailer appended once at the end of the CHA CI description, after every
// model block — SQC / preferential trade / invoice code come straight from
// the Products tab columns L/I/J (sqc_code / pref_trade_code / invoice_code),
// which already existed but weren't being used anywhere until now.
function buildChaCiTrailer_(sqcCode, prefTradeCode, invoiceCode, countryOfOrigin) {
  var lines = [];
  if (sqcCode) lines.push('SQC (Standard Quantity Code) – ' + sqcCode);
  if (prefTradeCode) lines.push('Preferential Trade Agreement Code: ' + prefTradeCode);
  if (invoiceCode) lines.push('Invoice Code: ' + invoiceCode);
  if (countryOfOrigin) lines.push((countryOfOrigin === 'INDIA' ? 'INDIAN' : String(countryOfOrigin).toUpperCase()) + ' ORIGIN');
  return lines.join('\n');
}

// Simple one-line reference used by CHA TI / CHA PL — "AS PER LC DOCUMENT".
function buildLcReferenceLine_(lcNumber, lcDate) {
  if (!lcNumber) return '';
  return 'MOTOR VEHICLES WITH HSN CODE AND CHASSIS NO AS PER LC DOCUMENT NO. ' + lcNumber + (lcDate ? (' DT. ' + lcDate) : '');
}

function buildPayload(ss, ctrl, inv) {
  const stock   = getSelectedStockSheet_(ss, ctrl);
  const company = ss.getSheetByName('Company');
  // Backfills T-W headers on THIS invoice's Stock tab if it predates them —
  // scoped to generation (not every UI call) to avoid needless writes.
  ensureStockExtraColumns_(stock);

  Logger.log('🔍 DEBUG: buildPayload v5.3 starting for invoice: ' + inv);

  let cif_usd     = Number(ctrl.getRange('C28').getValue()) || 0;
  const total_inr = Number(ctrl.getRange(CFG.totalInrCell).getValue()) || 0;
  let fobTotalUsd = Number(ctrl.getRange('C27').getValue()) || 0;
  let quantityForPayload = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;
  let generationSelectionActive = false;

  var payloadMode = ctrl.getRange(CFG.modeCell).getValue() || 'FINAL';
  let rawMatchedRows = stock.getRange('A4:W2000').getValues()
    .filter(function(r) { return rowMatchesInvoice_(r, inv); });
  let vehicles = rawMatchedRows.map(function(r) {
      return {
        chassis_no:     r[0] || '',
        engine_no:      r[1] || '',
        model:          r[2] || '',
        color:          r[3] || '',
        year:           r[4] || '',
        unit_price_usd: Number(r[9]) || 0,
        // Per-vehicle fields for the CI / CHA CI detailed description block
        // (cols T-W, see STOCK_EXTRA_HEADERS_) — all optional.
        eic_cert_no:            String(r[19] || '').trim(),
        first_registration_date: String(r[20] || '').trim(),
        district_origin_code:  String(r[21] || '').trim(),
        state_origin_code:     String(r[22] || '').trim()
      };
    });

  Logger.log('🚗 VEHICLES FOUND: ' + vehicles.length);
  if (vehicles.length === 0) {
    // Diagnostic: if this is still 0 after normInvoice_, dump the exact
    // char codes of C8 vs. every non-blank assigned_to value in this sheet's
    // invCol0 column, so a mismatched character shows up directly in the log
    // instead of guessing again.
    Logger.log('🔍 DEBUG: 0 vehicles — inv="' + inv + '" charCodes=' +
      String(inv).split('').map(function(c) { return c.charCodeAt(0); }).join(','));
    var debugInvCol0 = (String(payloadMode).trim() === 'PROFORMA') ? 18 : 7;
    stock.getRange('A4:S2000').getValues().forEach(function(r, i) {
      var av = r[debugInvCol0];
      if (av) Logger.log('  row' + (i + 4) + ' assigned_to="' + av + '" charCodes=' +
        String(av).split('').map(function(c) { return c.charCodeAt(0); }).join(','));
    });
  }
  vehicles.forEach(function(v, i) {
    Logger.log('  [' + i + '] chassis=' + v.chassis_no + ' | model="' + v.model + '" | price=' + v.unit_price_usd + ' | engine=' + v.engine_no + ' | color=' + v.color);
  });

  // ── C17 vehicle selection — generate for only a subset of assigned vehicles ─
  // C17 normally holds either nothing, or a single legacy autofill value
  // ("chassis - engine - model - color", picked just to fill F26 — never
  // filtered anything). The "Select Vehicles for Generation" sidebar writes a
  // comma-separated chassis list instead (always with a trailing comma, even
  // for one vehicle) — that comma is what distinguishes "explicit multi-select
  // in effect" from the old single autofill value, so legacy sheets keep
  // generating for every reserved vehicle exactly as before.
  var c17Raw = String(ctrl.getRange(CFG.stockCell).getValue() || '').trim();
  if (c17Raw.indexOf(',') !== -1) {
    var selectedSet = {};
    c17Raw.split(',').map(function(s) { return s.trim(); }).filter(Boolean)
      .forEach(function(c) { selectedSet[c.toUpperCase()] = true; });
    var filteredVehicles = vehicles.filter(function(v) { return selectedSet[String(v.chassis_no).trim().toUpperCase()]; });
    if (filteredVehicles.length > 0) {
      Logger.log('🎯 C17 selection active: generating for ' + filteredVehicles.length + ' of ' + vehicles.length + ' reserved vehicle(s)');
      vehicles = filteredVehicles;
      generationSelectionActive = true;
    } else {
      Logger.log('⚠ C17 selection set but matched 0 vehicles — ignoring selection, using all ' + vehicles.length + ' reserved vehicle(s)');
    }
  }

  // Build product lookup from Products tab
  // Col A=ID, B=product_name (match key), C=HSN Code (8-digit), D=Commercial desc, E=SCOMET desc,
  // O(14)=Packing List desc, P(15)=Tax Invoice desc, Q(16)=PI Format desc
  var productMap = {};
  var productSheet = ss.getSheetByName('Products');
  if (productSheet) {
    productSheet.getRange('A3:V2000').getValues().forEach(function(row) {  // A3 skips header; V = last desc column
      var modelKey = String(row[1]).toUpperCase().trim();  // col B = product_name (match key)
      if (modelKey) productMap[modelKey] = row;
    });
    Logger.log('📦 PRODUCT TAB keys (' + Object.keys(productMap).length + '): ' + Object.keys(productMap).join(', '));
  } else {
    Logger.log('⚠️ PRODUCT TAB not found — using defaults for all descriptions');
  }

  // Optional 'Descriptions' tab overrides — see DESCRIPTIONS_TAB_NAME_ note
  // near CFG for why this exists. Purely additive: empty map (i.e. exactly
  // today's behavior) if the tab doesn't exist yet or has no matching rows.
  var descriptionMap = {};
  var descriptionsSheet = ss.getSheetByName(DESCRIPTIONS_TAB_NAME_);
  if (descriptionsSheet) {
    descriptionsSheet.getRange('A4:O2000').getValues().forEach(function(row) {
      var modelKey = String(row[1] || '').toUpperCase().trim();  // col B = product_name (match key)
      if (modelKey) descriptionMap[modelKey] = row;
    });
    Logger.log('📝 DESCRIPTIONS TAB overrides (' + Object.keys(descriptionMap).length + '): ' + Object.keys(descriptionMap).join(', '));
  }

  var defaultDesc     = ctrl.getRange(CFG.itemDescriptionCell || 'C51').getValue() || 'Motorcycles';
  var defaultHsn      = ctrl.getRange(CFG.hsnCodeCell || 'C50').getValue() || '8711';
  var defaultDistrict = ctrl.getRange(CFG.districtOriginCell || 'C52').getValue() || '';
  var defaultState    = ctrl.getRange(CFG.stateOriginCell    || 'C53').getValue() || '29';
  var defaultHsnPi    = defaultHsn;  // PI HSN fallback = same as standard default HSN
  Logger.log('📋 DEFAULTS: desc="' + defaultDesc + '" | hsn="' + defaultHsn + '" | district="' + defaultDistrict + '" | state="' + defaultState + '"');

  // Group by model+price so each model gets its own HSN/description row
  var itemsObj = {};
  var controlUnitPrice = Number(ctrl.getRange(CFG.unitPriceCell).getValue()) || 0;
  vehicles.forEach(function(v) {
    var modelKey = String(v.model || '').toUpperCase().trim();

    var prod        = productMap[modelKey] || null;
    var descOverride = descriptionMap[modelKey] || null;
    Logger.log('  🔎 model="' + v.model + '" → key="' + modelKey + '" → productMatch=' + (prod ? 'YES' : 'NO (using defaults)') +
      ' → descriptionsOverride=' + (descOverride ? 'YES' : 'NO'));

    var hsnCode       = prod ? (String(prod[2]).trim()  || defaultHsn)      : defaultHsn;   // col C = HSN code
    var hsnCodePi     = prod ? (String(prod[18]).trim() || defaultHsnPi)    : defaultHsnPi; // col S = PI HSN (destination country)
    var descChaTi     = prod ? (String(prod[19]).trim() || defaultDesc)     : defaultDesc;   // col T = CHA Tax Invoice desc
    var descChaPl     = prod ? (String(prod[20]).trim() || defaultDesc)     : defaultDesc;   // col U = CHA Packing List desc
    var descChaCi     = prod ? (String(prod[21]).trim() || defaultDesc)     : defaultDesc;   // col V = CHA Commercial Invoice desc
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

    // Descriptions-tab override, by column index (see DESCRIPTIONS_HEADERS_):
    // 2=commercial, 3=scomet, 4=packing, 5=tax, 6=PI, 7=annexure1, 8=PI HSN,
    // 9=CHA TI, 10=CHA PL, 11=CHA CI. Used VERBATIM — no withModel() append —
    // since the whole point is exact text with no auto-modification. Falls
    // through to `fallback` (today's withModel()-wrapped value) if this
    // column is blank or the tab/row doesn't exist.
    function pick(colIdx, fallback) {
      var override = descOverride ? String(descOverride[colIdx] || '').trim() : '';
      return override || fallback;
    }

    if (!itemsObj[groupKey]) {
      itemsObj[groupKey] = {
        hsn_code:               hsnCode,
        hsn_code_pi:            pick(8, hsnCodePi),
        description:            pick(6, withModel(descPi)),  // default = PI FORMAT desc; backend overrides per-template
        description_commercial: pick(2, withModel(descComm)),
        description_scomet:     pick(3, withModel(descScomet)),
        description_packing:    pick(4, withModel(descPacking)),
        description_tax:        pick(5, withModel(descTax)),
        description_pi:         pick(6, withModel(descPi)),
        description_annexure1:  pick(7, withModel(descAnnexure1)),
        description_cha_ti:     pick(9, withModel(descChaTi)),
        description_cha_pl:     pick(10, withModel(descChaPl)),
        description_cha_ci:     pick(11, withModel(descChaCi)),
        quantity:               0,
        rate_per_unit:          price,
        amount_usd:             0,
        district_origin_code:   districtCode,
        state_origin_code:      stateCode,
        model_key:              modelKey,      // for the post-loop generator pass below
        model_display:          modelDisplay,
        vehicle_list:           []             // every vehicle of this model, for chassis-level generation
      };
      Logger.log('    ✅ NEW item group: key="' + groupKey + '" | desc_pi="' + withModel(descPi) + '"');
    }
    itemsObj[groupKey].quantity   += 1;
    itemsObj[groupKey].amount_usd += price;
    itemsObj[groupKey].vehicle_list.push(v);
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

  // ── C17 selection active — recompute quantity/FOB/CIF from selected vehicles ─
  // CONTROL!C26 (qty), C27 (FOB), C28 (CIF) are normally typed for the whole
  // invoice. When generating for only a subset via the C17 selector, those
  // static cells would still show the full-invoice numbers — recompute them
  // from the vehicles actually included so the printed totals match the
  // printed line items. Freight/insurance (F27/H27) and the INR/IGST side
  // stay exactly as typed — those aren't reliably derivable per-vehicle, so
  // adjust them manually in CONTROL if a partial shipment needs different values.
  if (generationSelectionActive) {
    var freightUsd   = Number(ctrl.getRange('F27').getValue()) || 0;
    var insuranceUsd = Number(ctrl.getRange('H27').getValue()) || 0;
    quantityForPayload = vehicles.length;
    fobTotalUsd = items.reduce(function(sum, it) { return sum + (it.amount_usd || 0); }, 0);
    cif_usd = fobTotalUsd + freightUsd + insuranceUsd;
    Logger.log('🎯 Recomputed for selection: qty=' + quantityForPayload + ' | fob=' + fobTotalUsd + ' | cif=' + cif_usd +
      ' (freight=' + freightUsd + ', insurance=' + insuranceUsd + ' — unchanged, adjust manually if needed)');
  }

  // ── Per-vehicle CI / CHA CI / CHA TI / CHA PL generation ─────────────────
  // Runs after every item's full vehicle_list is known. Only overwrites a
  // field when the Descriptions tab has no manual text for it — an explicit
  // override there always wins verbatim, exactly as before. Shipment-level
  // trailers (accessories/LC/TIN/dealer-cert for CI; SQC/trade/invoice/origin
  // for CHA CI) are appended only to the LAST item so they print once, not
  // once per model.
  var lcNumberVal = String(ctrl.getRange(CFG.lcCell).getValue() || '').trim();
  var lcDateVal   = String(ctrl.getRange(CFG.lcDateCell).getValue() || '').trim();
  var lcLine      = buildLcReferenceLine_(lcNumberVal, lcDateVal);
  var countryOfOriginVal = String(ctrl.getRange(CFG.countryOriginCell || 'F23').getValue() || 'INDIA').trim();

  items.forEach(function(it, idx) {
    var genProd = productMap[it.model_key] || null;
    var genOverride = descriptionMap[it.model_key] || null;
    var engineCc = genOverride ? String(genOverride[12] || '').trim() : '';
    var make     = genOverride ? String(genOverride[13] || '').trim() : '';
    var accessories = genOverride ? String(genOverride[14] || '').trim() : '';
    var hsCodeDotted = formatHsCodeDotted_(it.hsn_code_pi || it.hsn_code);

    var ciManual = genOverride ? String(genOverride[2] || '').trim() : '';
    if (!ciManual) {
      it.description_commercial = buildVehicleDetailBlock_(it.model_display, engineCc, hsCodeDotted, make, it.vehicle_list);
    }

    // Unlike Commercial Invoice/CHA CI (where a Descriptions-tab manual
    // override replaces the whole generated block outright), Packing List
    // always gets the VIN breakdown appended — the base line (Products tab
    // default, or a Descriptions-tab override if one exists) is kept as-is,
    // but which VINs are listed underneath must always reflect this
    // invoice's actual assigned/selected vehicles, never a static manual list.
    it.description_packing = buildPackingListVehicleBlock_(it.description_packing, it.vehicle_list);

    var tiManual = genOverride ? String(genOverride[9] || '').trim() : '';
    if (!tiManual && lcLine) it.description_cha_ti = lcLine;

    var plManual = genOverride ? String(genOverride[10] || '').trim() : '';
    if (!plManual && lcLine) it.description_cha_pl = lcLine;

    var ciChaManual = genOverride ? String(genOverride[11] || '').trim() : '';
    if (!ciChaManual && lcLine) {
      it.description_cha_ci = buildChaCiVehicleBlock_(lcLine, it.vehicle_list, it.district_origin_code, it.state_origin_code);
    }

    if (idx === items.length - 1) {
      if (!ciManual) {
        var fallbackYear = it.vehicle_list.length ? String(it.vehicle_list[0].year || '') : '';
        var invoiceDateStr = Utilities.formatDate(ctrl.getRange(CFG.dateCell).getValue(), 'GMT+5:30', 'dd.MM.yyyy');
        var trailer = buildCommercialInvoiceTrailer_(ctrl, accessories, inv, invoiceDateStr, countryOfOriginVal, fallbackYear);
        if (trailer) it.description_commercial += '\n' + trailer;
      }
      if (!ciChaManual && lcLine) {
        var prefTradeCode = genProd ? String(genProd[8] || '').trim() : '';
        var invoiceCode   = genProd ? String(genProd[9] || '').trim() : '';
        var sqcCode       = genProd ? String(genProd[11] || '').trim() : '';
        var chaCiTrailer = buildChaCiTrailer_(sqcCode, prefTradeCode, invoiceCode, countryOfOriginVal);
        if (chaCiTrailer) it.description_cha_ci += '\n' + chaCiTrailer;
      }
    }
  });

  // ── Invoice_Descriptions override — highest priority, applied last ───────
  // One row per invoice, exact text typed once per document. Completely
  // bypasses the per-model Descriptions tab AND the auto-generated CI/CHA
  // blocks above when a cell here is filled in. No per-vehicle data, no
  // Stock/CONTROL setup needed — this is the simplest path.
  var invoiceDescSheet = ss.getSheetByName(INVOICE_DESC_TAB_NAME_);
  if (invoiceDescSheet && items.length > 0) {
    // PROFORMA shipments are looked up by pi_invoice_no (col B), everything
    // else by invoice_no (col A) — same row can carry both numbers so it
    // answers to either stage of the same shipment. See header note above.
    var invoiceDescKeyCol = (String(payloadMode).trim() === 'PROFORMA') ? 1 : 0;
    var invoiceDescRow = null;
    invoiceDescSheet.getRange('A4:K2000').getValues().some(function(row) {
      if (row[invoiceDescKeyCol] && sameInvoice_(row[invoiceDescKeyCol], inv)) { invoiceDescRow = row; return true; }
      return false;
    });
    if (invoiceDescRow) {
      Logger.log('📋 INVOICE_DESCRIPTIONS override found for invoice="' + inv + '" (matched via ' +
        (invoiceDescKeyCol === 1 ? 'pi_invoice_no' : 'invoice_no') + ')');
      // Fields where the template loops per model row but the invoice-level
      // text should still print exactly once, on the first row — put the
      // full text on the first item only, blank the rest, so a bulk
      // shipment-wide description never repeats once per model/price group.
      var concatFields = [
        ['description_commercial', 2], ['description_packing', 4],
        ['description_tax', 5], ['description_pi', 6], ['description', 6],
        ['description_cha_ti', 8], ['description_cha_pl', 9], ['description_cha_ci', 10]
      ];
      concatFields.forEach(function(pair) {
        var text = String(invoiceDescRow[pair[1]] || '').trim();
        if (!text) return;
        items[0][pair[0]] = text;
        for (var ri = 1; ri < items.length; ri++) items[ri][pair[0]] = '';
      });
      // Fields where the template does NOT loop per item row (SCOMET
      // Declaration uses a single scomet_product_desc field; Annexure_1
      // loops vin_list, not items) — same text can safely apply to every
      // item since only items[0]'s value ever actually gets read downstream.
      var perRowFields = [
        ['description_scomet', 3], ['description_annexure1', 7]
      ];
      perRowFields.forEach(function(pair) {
        var text = String(invoiceDescRow[pair[1]] || '').trim();
        if (!text) return;
        items.forEach(function(it) { it[pair[0]] = text; });
      });
    }
  }

  // Snapshot the final, fully-resolved description text into the
  // Invoice_Descriptions tab's read-only preview columns — see
  // writeDescriptionPreview_ note above. Runs after every override/auto-
  // generation pass above, so it always matches what actually printed.
  //
  // `inv` (CONTROL!C8) is whichever number is active THIS run — invoice_no
  // for FINAL/DRAFT, pi_invoice_no for PROFORMA. The OTHER stage's number,
  // if this shipment was ever generated under it too, already sits on these
  // same Stock rows (col H / col S) since assignment writes whichever
  // column matched the mode active at assignment time — reading both here
  // is what lets the two stages consolidate into one preview row.
  var previewInvoiceNo   = (String(payloadMode).trim() === 'PROFORMA') ? '' : inv;
  var previewPiInvoiceNo = (String(payloadMode).trim() === 'PROFORMA') ? inv : '';
  for (var pmi = 0; pmi < rawMatchedRows.length; pmi++) {
    if (!previewInvoiceNo)   previewInvoiceNo   = String(rawMatchedRows[pmi][7] || '').trim();
    if (!previewPiInvoiceNo) previewPiInvoiceNo = String(rawMatchedRows[pmi][18] || '').trim();
    if (previewInvoiceNo && previewPiInvoiceNo) break;
  }
  writeDescriptionPreview_(ss, previewInvoiceNo, previewPiInvoiceNo, items);

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
      quantity:          quantityForPayload,
      unit_price_usd:    Number(ctrl.getRange(CFG.unitPriceCell).getValue()) || 0,
      fob_total_usd:     fobTotalUsd,
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
    amount_usd_words:      AMOUNTWORDS(cif_usd, 'USD', true),
    amount_inr_words:      AMOUNTWORDS(total_inr, 'INR', true),

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

// ── Preview descriptions only — no documents generated ──────────────────────
// buildPayload() already resolves every description (Products default →
// Descriptions-tab override → auto-generated VIN/CI blocks → Invoice_
// Descriptions override) and writes the result into the Invoice_Descriptions
// preview columns as a side effect — this just calls buildPayload() and
// stops there, skipping the UrlFetchApp call that actually creates .docx
// files. Respects whatever vehicles are currently selected/assigned: the
// same C17 "Select Vehicles for Generation" subset filter buildPayload()
// always applies, so previewing after narrowing to a partial selection
// shows exactly what THAT subset would print, not the whole invoice.
function generateDescriptionPreview() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName('CONTROL');
  const inv  = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const ui   = SpreadsheetApp.getUi();

  if (!inv) { ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK); return; }

  Logger.log('🔍 DESCRIPTION PREVIEW starting for invoice: ' + inv);

  try {
    const payload = buildPayload(ss, ctrl, inv);
    logAudit(inv, 'DESCRIPTION_PREVIEW', payload.vehicles.length + ' vehicle(s), ' + payload.items.length + ' model group(s)');
    ui.alert('✅ Description Preview Updated',
      'Resolved descriptions for invoice ' + inv + ' (' + payload.vehicles.length + ' vehicle(s) currently selected/assigned) ' +
      'and wrote them to the Invoice_Descriptions tab, columns M onward.\n\n' +
      'No documents were generated — this only updates the preview so you can check the wording first.',
      ui.ButtonSet.OK);
  } catch (err) {
    Logger.log('❌ Description Preview Error: ' + err);
    ui.alert('❌ Error', String(err), ui.ButtonSet.OK);
  }
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

function generateAnnexureCDocument() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName('CONTROL');
  const inv  = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const ui   = SpreadsheetApp.getUi();

  if (!inv) { ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK); return; }

  Logger.log('📑 ANNEXURE C GENERATION starting for invoice: ' + inv);

  const payload    = buildPayload(ss, ctrl, inv);
  const encodedInv = encodeURIComponent(inv);
  // Pass documents=annexure_c so backend generates Annexure C only
  const fullUrl    = CFG.webhookUrl + 'generate?invoice_no=' + encodedInv + '&documents=annexure_c';

  Logger.log('📤 Annexure C URL: ' + fullUrl);

  try {
    const res  = UrlFetchApp.fetch(fullUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText();

    Logger.log('📥 Annexure C Response HTTP ' + code + ': ' + body.substring(0, 500));

    if (code === 200 || code === 201) {
      logAudit(inv, 'ANNEXURE_C_GENERATED', 'Annexure C only');
      const parsed = JSON.parse(body);
      const files = parsed.generated_files || [];
      if (files.length === 0) {
        ui.alert('⚠ Annexure C Generated', 'Document was processed but no download link returned.\nCheck backend logs.', ui.ButtonSet.OK);
        return;
      }
      const driveUrl = files[0].gcs_url || '';
      const fileName = files[0].download_name || ('Annexure_C_' + inv + '.docx');
      const linkHtml = driveUrl
        ? '<p style="margin:12px 0"><a href="' + driveUrl + '" target="_blank" style="background:#1e3a5f;color:white;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:14px">⬇️ Download Annexure C</a></p>'
          + '<p style="font-size:11px;color:#666;margin-top:8px;word-break:break-all">' + driveUrl + '</p>'
        : '<p style="color:#dc2626">No download link returned — check backend logs.</p>';
      SpreadsheetApp.getUi().showModalDialog(
        HtmlService.createHtmlOutput(
          '<style>body{font-family:Arial,sans-serif;padding:20px;text-align:center}</style>' +
          '<h3 style="color:#1e3a5f">✅ Annexure C Generated</h3>' +
          '<p><strong>Invoice:</strong> ' + inv + '</p>' +
          '<p><strong>File:</strong> ' + fileName + '</p>' +
          linkHtml
        ).setWidth(480).setHeight(200),
        'Annexure C — ' + inv
      );
    } else {
      logAudit(inv, 'ANNEXURE_C_FAILED', 'HTTP ' + code);
      ui.alert('❌ Annexure C Generation Failed (HTTP ' + code + ')', body.substring(0, 800), ui.ButtonSet.OK);
    }
  } catch(err) {
    Logger.log('❌ Annexure C Error: ' + err);
    ui.alert('❌ Connection Failed', 'Could not reach backend.\n\nError: ' + String(err), ui.ButtonSet.OK);
  }
}

function generateCHADocuments() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName('CONTROL');
  const inv  = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const ui   = SpreadsheetApp.getUi();

  if (!inv) { ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK); return; }

  Logger.log('🏦 CHA GENERATION starting for invoice: ' + inv);

  const payload    = buildPayload(ss, ctrl, inv);
  const encodedInv = encodeURIComponent(inv);
  const fullUrl    = CFG.webhookUrl + 'generate?invoice_no=' + encodedInv + '&documents=cha_tax_invoice&documents=cha_packing_list&documents=cha_commercial_invoice';

  Logger.log('📤 CHA URL: ' + fullUrl);

  try {
    const res  = UrlFetchApp.fetch(fullUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText();

    Logger.log('📥 CHA Response HTTP ' + code + ': ' + body.substring(0, 500));

    if (code === 200 || code === 201) {
      logAudit(inv, 'CHA_GENERATED', 'CHA TI + CHA PL + CHA CI');
      const parsed = JSON.parse(body);
      const files  = parsed.generated_files || [];
      if (files.length === 0) {
        ui.alert('⚠ CHA Generated', 'Documents processed but no download links returned.\nCheck backend logs.', ui.ButtonSet.OK);
        return;
      }
      const links = files.map(function(f) {
        return '<p style="margin:8px 0"><a href="' + f.gcs_url + '" target="_blank" style="background:#1e3a5f;color:white;padding:8px 16px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:13px">⬇️ ' + f.download_name + '</a></p>';
      }).join('');
      SpreadsheetApp.getUi().showModalDialog(
        HtmlService.createHtmlOutput(
          '<style>body{font-family:Arial,sans-serif;padding:20px;text-align:center}</style>' +
          '<h3 style="color:#1e3a5f">✅ CHA Documents Generated</h3>' +
          '<p><strong>Invoice:</strong> ' + inv + '</p>' +
          links
        ).setWidth(500).setHeight(220),
        'CHA Documents — ' + inv
      );
    } else {
      logAudit(inv, 'CHA_FAILED', 'HTTP ' + code);
      ui.alert('❌ CHA Generation Failed (HTTP ' + code + ')', body.substring(0, 800), ui.ButtonSet.OK);
    }
  } catch(err) {
    Logger.log('❌ CHA Error: ' + err);
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
  const ui      = SpreadsheetApp.getUi();
  const stock   = resolveStockOrAlert_(ss, ctrl, ui);
  if (!stock) return;
  const invoice = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const qty     = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;
  const sideMode  = ctrl.getRange(CFG.modeCell).getValue() || 'FINAL';
  const sideInvCol = (sideMode === 'PROFORMA') ? 18 : 7;  // col S(18) for PROFORMA, col H(7) for FINAL/DRAFT

  if (!invoice) { SpreadsheetApp.getUi().alert('❌ No invoice number found in C8.'); return; }

  const data    = stock.getRange('A4:S2000').getValues();
  const visible = data.filter(function(r) { return r[0] && (r[6] === 'AVAILABLE' || sameInvoice_(r[sideInvCol], invoice)); });
  const alreadyAssigned = data.filter(function(r) { return r[0] && sameInvoice_(r[sideInvCol], invoice); }).length;

  const rows = visible.map(function(r) {
    const isAssigned  = sameInvoice_(r[sideInvCol], invoice);
    const statusColor = r[6] === 'AVAILABLE' ? '#16a34a' : '#d97706';
    const assignedTo  = r[sideInvCol] && !sameInvoice_(r[sideInvCol], invoice)
      ? '<span style="color:#dc2626;font-size:10px">Taken: ' + r[sideInvCol] + '</span>'
      : sameInvoice_(r[sideInvCol], invoice) ? '<span style="color:#d97706;font-size:10px">This invoice</span>' : '';
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

function getCustomerInfo_(ss) {
  var ctrl    = ss.getSheetByName('CONTROL');
  var dropdown = String(ctrl.getRange(CFG.customerCell).getValue() || '').trim();
  var companyName = '', contactName = '';
  if (!dropdown) return { companyName: companyName, contactName: contactName };
  var custSheet = ss.getSheetByName('Customers');
  if (custSheet) {
    var rows = custSheet.getRange('A4:L2000').getValues();
    for (var ci = 0; ci < rows.length; ci++) {
      var cr = rows[ci];
      if (!cr[0] && !cr[11]) break;
      if (String(cr[11]).trim() === dropdown) {
        companyName = String(cr[1]).trim();        // col B = company_name
        contactName = String(cr[7]  || '').trim(); // col H = contact_name
        if (!contactName) {
          var parts = dropdown.split(' — ');  // em-dash "CONTACT — COMPANY"
          contactName = parts.length > 1 ? parts[0].trim() : '';
        }
        break;
      }
    }
  }
  return { companyName: companyName, contactName: contactName };
}

function assignVehiclesFromSidebar(chassisList, invoiceNo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl  = ss.getSheetByName('CONTROL');
  const stock = getSelectedStockSheet_(ss, ctrl);
  const assignMode   = ctrl.getRange(CFG.modeCell).getValue() || 'FINAL';
  const invCol0      = (assignMode === 'PROFORMA') ? 18 : 7;   // 0-based: col S(18) or col H(7)
  const invColNum    = (assignMode === 'PROFORMA') ? 19 : 8;   // 1-based: col S(19) or col H(8)
  const data  = stock.getRange('A4:S2000').getValues();
  const info  = getCustomerInfo_(ss);
  let assigned = 0, released = 0;
  data.forEach(function(row, i) {
    if (!row[0]) return;
    const chassis = String(row[0]);
    if (chassisList.indexOf(chassis) !== -1) {
      stock.getRange(i+4, 7).setValue('RESERVED');
      stock.getRange(i+4, invColNum).setValue(invoiceNo);
      stock.getRange(i+4, 17).setValue(info.companyName);  // col Q = company_name
      stock.getRange(i+4, 18).setValue(info.contactName);  // col R = contact_name
      assigned++;
    } else if (sameInvoice_(row[invCol0], invoiceNo)) {
      stock.getRange(i+4, 7).setValue('AVAILABLE');
      stock.getRange(i+4, invColNum).setValue('');
      stock.getRange(i+4, 17).setValue('');  // col Q
      stock.getRange(i+4, 18).setValue('');  // col R
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
  const stock = resolveStockOrAlert_(ss, ctrl, ui);
  if (!stock) return;

  const inv = ctrl.getRange(CFG.invoiceNoCell).getValue();
  if (!inv) return ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK);

  const qty = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;
  if (qty <= 0) return ui.alert('❌ Quantity is Zero', 'Update quantity in C26 first.', ui.ButtonSet.OK);

  const bulkMode   = ctrl.getRange(CFG.modeCell).getValue() || 'FINAL';
  const bulkInvCol0   = (bulkMode === 'PROFORMA') ? 18 : 7;   // 0-based: col S(18) or col H(7)
  const bulkInvColNum = (bulkMode === 'PROFORMA') ? 19 : 8;   // 1-based: col S(19) or col H(8)

  const alreadyAssigned = stock.getRange('A4:S2000').getValues()
    .filter(function(r) { return sameInvoice_(r[bulkInvCol0], inv); }).length;
  if (alreadyAssigned >= qty)
    return ui.alert('✅ Already Fully Assigned',
      'Invoice ' + inv + ' already has all ' + qty + ' vehicles assigned.\n\nUse the sidebar to review.', ui.ButtonSet.OK);

  const remaining = qty - alreadyAssigned;
  const resp = ui.prompt('🔗 Bulk Assign by Model — ' + inv,
    'Need ' + remaining + ' more vehicle(s)  (' + alreadyAssigned + ' of ' + qty + ' already assigned).\n\n' +
    'Enter model and/or colour — partial match works, words in any order ' +
    '(e.g. "PULSAR" matches "PULSAR NS200"; "PULSAR BLACK" matches a black Pulsar):',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const filter = resp.getResponseText().toUpperCase().trim();
  if (!filter) return ui.alert('❌ Empty Input', 'Please enter a model name.', ui.ButtonSet.OK);
  // Match every word of the query against "MODEL COLOUR" combined — model and
  // colour are separate columns, so a single indexOf(filter) on model alone
  // could never match a query like "PULSAR BLACK".
  const filterTokens = filter.split(/\s+/).filter(Boolean);

  const allData = stock.getRange('A4:S2000').getValues();
  const matches = [];
  for (var i = 0; i < allData.length; i++) {
    const row = allData[i];
    if (!row[0] || row[6] !== 'AVAILABLE') continue;
    const haystack = (String(row[2]) + ' ' + String(row[3])).toUpperCase();
    if (!filterTokens.every(function(tok) { return haystack.indexOf(tok) !== -1; })) continue;
    matches.push({ rowIndex: i, chassis: row[0], model: row[2], color: row[3] });
  }

  if (matches.length === 0)
    return ui.alert('⚠ No Vehicles Found',
      'No AVAILABLE vehicles matched "' + filter + '".\n\n' +
      'Tips:\n• Try a shorter term (just the model, or just the colour)\n• Check the Stock sheet for the exact model/colour spelling\n• All matching vehicles may already be RESERVED or SHIPPED',
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

  const info = getCustomerInfo_(ss);
  toAssign.forEach(function(v) {
    stock.getRange(v.rowIndex+4, 7).setValue('RESERVED');
    stock.getRange(v.rowIndex+4, bulkInvColNum).setValue(inv);
    stock.getRange(v.rowIndex+4, 17).setValue(info.companyName);
    stock.getRange(v.rowIndex+4, 18).setValue(info.contactName);
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
  const audit = ss.getSheetByName('Audit_Trail');
  let stock = null;
  try { stock = getSelectedStockSheet_(ss, ctrl); } catch (err) { /* non-critical FYI check below is skipped if no stock tab is set */ }

  const currentInv = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const currentQty = Number(ctrl.getRange(CFG.qtyCell).getValue()) || 0;

  if (currentInv && stock) {
    const cloneMode   = ctrl.getRange(CFG.modeCell).getValue() || 'FINAL';
    const cloneInvCol = (cloneMode === 'PROFORMA') ? 18 : 7;
    const reserved = stock.getRange('A4:S2000').getValues()
      .filter(function(r) { return sameInvoice_(r[cloneInvCol], currentInv); }).length;
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
  const stock = getSelectedStockSheet_(ss, ctrl);
  const inv   = ctrl.getRange(CFG.invoiceNoCell).getValue();
  if (!inv) throw new Error('No invoice number found in C8.');
  const qaMode    = ctrl.getRange(CFG.modeCell).getValue() || 'FINAL';
  const qaInvColNum = (qaMode === 'PROFORMA') ? 19 : 8;  // 1-based: col S(19) or col H(8)

  let totalAdded = 0;
  for (var i = 0; i < models.length; i++) {
    if (!models[i] || qtys[i] <= 0) continue;
    const startRow = stock.getLastRow() + 1;
    for (var j = 0; j < qtys[i]; j++) {
      const row = startRow + j;
      stock.getRange(row, 1).setValue('MD2' + Utilities.formatString('%07d', Math.floor(Math.random() * 9999999)));
      stock.getRange(row, 3).setValue(models[i]);
      stock.getRange(row, 7).setValue('RESERVED');
      stock.getRange(row, qaInvColNum).setValue(inv);
      stock.getRange(row, 10).setValue(prices[i]);
      totalAdded++;
    }
  }
  logAudit(inv, 'QUICK_MULTI_PRODUCT', totalAdded + ' vehicles added across ' + models.length + ' model(s)');
  updateChassisDropdown();
  return totalAdded + ' vehicle(s) added and reserved under invoice ' + inv;
}

// ── Select Vehicles for Generation (C17 checkbox sidebar) ───────────────────
// When an invoice has multiple vehicles RESERVED against it but you only
// want to generate documents for some of them right now (e.g. a partial
// shipment), this lets you check exactly which ones — writes a comma-
// separated chassis list into CONTROL!C17 (always with a trailing comma, so
// buildPayload() can tell it apart from C17's older single-value autofill
// use, see the C17 filter note in buildPayload()). Clearing C17 (or picking
// every vehicle) reverts to the old behavior: every reserved vehicle for the
// invoice gets included.
function showGenerationVehicleSelector() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl  = ss.getSheetByName('CONTROL');
  const ui    = SpreadsheetApp.getUi();
  const stock = resolveStockOrAlert_(ss, ctrl, ui);
  if (!stock) return;

  const invoice = ctrl.getRange(CFG.invoiceNoCell).getValue();
  if (!invoice) { ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK); return; }

  const mode   = ctrl.getRange(CFG.modeCell).getValue() || 'FINAL';
  const invCol = (mode === 'PROFORMA') ? 18 : 7;
  const reserved = stock.getRange('A4:S2000').getValues()
    .filter(function(r) { return r[0] && sameInvoice_(r[invCol], invoice); });

  if (reserved.length === 0) {
    ui.alert('📋 No Vehicles Assigned',
      'No vehicles are currently assigned to invoice ' + invoice + ' in Stock tab "' + stock.getName() + '".\n\n' +
      'If the vehicles are actually on a different monthly Stock tab, check CONTROL!' + CFG.stockTabCell +
      ' (the "Stock Tab" dropdown) — this always searches whichever tab is selected there, not whichever tab ' +
      'you have open in the browser.',
      ui.ButtonSet.OK);
    return;
  }

  // Existing C17 selection (comma list) pre-checks those chassis; otherwise
  // every reserved vehicle starts checked — matches today's "include all" default.
  const c17Raw = String(ctrl.getRange(CFG.stockCell).getValue() || '').trim();
  const existingSelection = {};
  if (c17Raw.indexOf(',') !== -1) {
    c17Raw.split(',').map(function(s) { return s.trim().toUpperCase(); }).filter(Boolean)
      .forEach(function(c) { existingSelection[c] = true; });
  }
  const hasExistingSelection = Object.keys(existingSelection).length > 0;

  const rows = reserved.map(function(r) {
    const chassis = String(r[0]);
    const checked = hasExistingSelection ? !!existingSelection[chassis.toUpperCase()] : true;
    return '<tr><td style="text-align:center"><input type="checkbox" class="vchk" data-chassis="' + chassis + '" ' + (checked ? 'checked' : '') + '></td>' +
      '<td style="font-family:monospace;font-size:11px">' + chassis + '</td>' +
      '<td>' + (r[2] || '') + '</td><td>' + (r[3] || '') + '</td></tr>';
  }).join('');

  const html = HtmlService.createHtmlOutput(
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:12px;background:#f8fafc}' +
    '.header{background:#1e3a5f;color:white;padding:10px 14px;border-radius:8px;margin-bottom:12px}' +
    '.header h3{font-size:14px;font-weight:600}.header p{font-size:11px;opacity:.8;margin-top:3px}' +
    'table{width:100%;border-collapse:collapse;background:white;border-radius:8px;border:1px solid #e2e8f0}' +
    'thead th{background:#1e3a5f;color:white;padding:7px 8px;font-size:11px;font-weight:600;text-align:left}' +
    'tbody td{padding:6px 8px;border-bottom:1px solid #f1f5f9;font-size:12px}' +
    'tbody tr:last-child td{border-bottom:none}' +
    'input[type=checkbox]{width:15px;height:15px;cursor:pointer;accent-color:#0d9488}' +
    '.btn{width:100%;padding:11px;background:#0d9488;color:white;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;margin-top:10px}' +
    '.btn:disabled{opacity:.45;cursor:not-allowed}' +
    '.btn-clear{width:100%;padding:8px;background:white;color:#475569;border:1px solid #cbd5e1;border-radius:8px;font-size:12px;cursor:pointer;margin-top:6px}' +
    '#msg{margin-top:10px;padding:8px 12px;border-radius:6px;font-size:12px;font-weight:600;display:none}' +
    '#msg.success{background:#dcfce7;color:#15803d;display:block}#msg.error{background:#fee2e2;color:#dc2626;display:block}' +
    '</style>' +
    '<div class="header"><h3>🎯 Select Vehicles for Generation</h3>' +
    '<p>Invoice: <strong>' + invoice + '</strong> &nbsp;·&nbsp; ' + reserved.length + ' vehicle(s) reserved</p></div>' +
    '<table><thead><tr><th>✓</th><th>Chassis No.</th><th>Model</th><th>Colour</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>' +
    '<button class="btn" id="saveBtn" onclick="save()">Apply Selection to C17</button>' +
    '<button class="btn-clear" onclick="clearSel()">Clear (include all reserved vehicles)</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'function save(){' +
    'var sel=[].slice.call(document.querySelectorAll(".vchk:checked")).map(function(c){return c.dataset.chassis;});' +
    'if(sel.length===0){document.getElementById("msg").textContent="❌ Select at least one vehicle.";document.getElementById("msg").className="error";return;}' +
    'var btn=document.getElementById("saveBtn"),msg=document.getElementById("msg");' +
    'btn.disabled=true;btn.textContent="Saving…";msg.className="";msg.style.display="none";' +
    'google.script.run' +
    '.withSuccessHandler(function(r){msg.textContent="✅ "+r;msg.className="success";btn.textContent="Apply Selection to C17";btn.disabled=false;})' +
    '.withFailureHandler(function(e){msg.textContent="❌ Error: "+e.message;msg.className="error";btn.textContent="Apply Selection to C17";btn.disabled=false;})' +
    '.applyGenerationVehicleSelection(sel,"' + invoice + '");}' +
    'function clearSel(){' +
    'var msg=document.getElementById("msg");' +
    'google.script.run' +
    '.withSuccessHandler(function(r){msg.textContent="✅ "+r;msg.className="success";})' +
    '.withFailureHandler(function(e){msg.textContent="❌ Error: "+e.message;msg.className="error";})' +
    '.clearGenerationVehicleSelection("' + invoice + '");}' +
    '<\/script>'
  ).setTitle('Select Vehicles — ' + invoice).setWidth(480);
  ui.showModalDialog(html, 'Select Vehicles for Generation — ' + invoice);
}

// Writes the chosen chassis list into C17 with a trailing comma — see the
// C17 filter note in buildPayload() for why the trailing comma matters even
// when only one chassis is selected.
function applyGenerationVehicleSelection(chassisList, invoiceNo) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName('CONTROL');
  if (!chassisList || chassisList.length === 0) throw new Error('No vehicles selected.');
  ctrl.getRange(CFG.stockCell).setValue(chassisList.join(',') + ',');
  refreshMultiProductItemsTable_(ss, ctrl);
  logAudit(invoiceNo, 'GENERATION_SELECTION_SET', chassisList.length + ' of reserved vehicle(s) selected: ' + chassisList.join(', '));
  return chassisList.length + ' vehicle(s) selected — generation will use only these.';
}

// Clears C17 back to blank — buildPayload() then falls back to including
// every vehicle RESERVED for this invoice, exactly like before this feature existed.
function clearGenerationVehicleSelection(invoiceNo) {
  const ctrl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL');
  ctrl.getRange(CFG.stockCell).clearContent();
  updateChassisDropdown();
  logAudit(invoiceNo, 'GENERATION_SELECTION_CLEARED', 'Reverted to including all reserved vehicles');
  return 'Selection cleared — generation will include all reserved vehicles.';
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

function showAssignedVehiclesPanel() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl     = ss.getSheetByName('CONTROL');
  const ui       = SpreadsheetApp.getUi();
  const stock    = resolveStockOrAlert_(ss, ctrl, ui);
  if (!stock) return;
  const invoice  = ctrl.getRange(CFG.invoiceNoCell).getValue();
  const customer = ctrl.getRange(CFG.customerCell).getValue();

  if (!invoice) { ui.alert('❌ No Invoice', 'Set invoice number in C8 first.', ui.ButtonSet.OK); return; }

  const panelMode   = ctrl.getRange(CFG.modeCell).getValue() || 'FINAL';
  const panelInvCol = (panelMode === 'PROFORMA') ? 18 : 7;
  const allData  = stock.getRange('A4:S2000').getValues();
  const assigned = allData.filter(function(r) { return r[0] && sameInvoice_(r[panelInvCol], invoice); });

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

// LEGACY — read-only reporting only. Groups the (pre-migration) 'Stock' sheet
// by Ship Date into SR_<Month> snapshot tabs. This is NOT the editable monthly
// Stock_YYYY_MM system (see createMonthlyStockTab / migrateStockToMonthlyTabs
// above) — most stock has no Ship Date yet, so this usually collapses
// everything into a single "SR_Unshipped" tab. Kept only for old ship-date
// reports; once 'Stock' has been migrated/archived this will just say so.
function buildMonthlyStockTabs() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stock = ss.getSheetByName('Stock');
  const ui    = SpreadsheetApp.getUi();

  if (!stock) {
    ui.alert('❌ No Stock Sheet',
      'Could not find the "Stock" sheet — it has likely already been migrated to monthly Stock_YYYY_MM tabs.\n\n' +
      'This menu item (Legacy Ship-Date Report) only works against the old single Stock sheet. ' +
      'For the current monthly system, use "Create Monthly Stock Tab" or pick a tab from CONTROL!' + CFG.stockTabCell + '.',
      ui.ButtonSet.OK);
    return;
  }

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

// Builds the customer-picker data validation rule from the Customers sheet
// (col L = smart_dropdown), or null if there's nothing to build it from.
// Shared by setupCustomerDropdowns() (applies to all tabs, on demand) and
// ensureMonthlyStockTab_() (applies automatically to a freshly created tab).
function buildCustomerDropdownRule_(custSheet) {
  if (!custSheet) return null;
  const custData   = custSheet.getRange('A2:L2000').getValues();
  const dropValues = [];
  for (var i = 0; i < custData.length; i++) {
    var val = String(custData[i][11]).trim();  // col L = smart_dropdown
    if (val) dropValues.push(val);
  }
  if (dropValues.length === 0) return null;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(dropValues, true)
    .setAllowInvalid(true)
    .setHelpText('Pick customer — name and company auto-fill into Q and R')
    .build();
  return { rule: rule, values: dropValues };
}

function applyCustomerDropdownToSheet_(stockSheet, rule) {
  if (!stockSheet.getRange(1, 17).getValue()) stockSheet.getRange(1, 17).setValue('customer_name');
  if (!stockSheet.getRange(1, 18).getValue()) stockSheet.getRange(1, 18).setValue('company_name');
  const lastRow = Math.max(stockSheet.getLastRow(), 4);
  stockSheet.getRange(4, 17, lastRow - 3, 1).setDataValidation(rule);
}

function setupCustomerDropdowns() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const custSheet  = ss.getSheetByName('Customers');
  const ui         = SpreadsheetApp.getUi();

  // Applies to every Stock tab (legacy 'Stock' plus all Stock_YYYY_MM tabs) —
  // this is a one-time/occasional setup action, not per-invoice, so it isn't
  // scoped to whichever tab is currently selected in CONTROL.
  const stockSheets = [];
  const legacy = ss.getSheetByName('Stock');
  if (legacy) stockSheets.push(legacy);
  listStockTabNames_(ss).forEach(function(n) { stockSheets.push(ss.getSheetByName(n)); });

  if (stockSheets.length === 0 || !custSheet) {
    ui.alert('❌ Sheet not found', 'Could not find any Stock tab or the Customers sheet.', ui.ButtonSet.OK);
    return;
  }

  const built = buildCustomerDropdownRule_(custSheet);
  if (!built) {
    ui.alert('⚠ No Customers', 'No smart_dropdown values found in Customers sheet (col L).', ui.ButtonSet.OK);
    return;
  }

  stockSheets.forEach(function(stockSheet) { applyCustomerDropdownToSheet_(stockSheet, built.rule); });

  ui.alert(
    '✅ Customer Dropdowns Ready',
    built.values.length + ' customer(s) loaded into column Q across ' + stockSheets.length + ' Stock tab(s): ' +
    stockSheets.map(function(s) { return s.getName(); }).join(', ') + '.\n\n' +
    'Picking from the dropdown auto-fills:\n• Q = Contact Name\n• R = Company Name',
    ui.ButtonSet.OK
  );
}

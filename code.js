/**
 * INVOICE MASTER TRACKER AUTOMATION
 * * Features:
 * - Auto-creates custom menu, system config, GL config, RDC aliases, Email template, and Header config tabs.
 * - Dynamically loads any carrier ending in _ROOT_FOLDER from config.
 * - Extracts Shift Data and Additional Costs dynamically based on custom Header Config aliases.
 * - Enforces GL Cost Tolerances and Ignore Rules to prevent $0.00 or irrelevant items from being coded.
 * - Cross-references with Haulier Reports AND writes the Invoice # & Amount back to the Haulier sheet.
 * - Intelligently reallocates TUs marked as 'O' (Others) into specific GL codes.
 * - Aggregates invoice totals and applies customizable GL coding.
 * - Auto-generates stamped PDFs for the accounting team.
 * - Processes from root folders and automatically sorts files into RDC sub-folders.
 * - Automatically archives old data and clears sheets before new imports.
 * - Generates Discrepancy Email drafts and suspends GL coding on problematic invoices.
 */

// --- MENU & UI ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Invoice Automation')
    .addItem('1. Initialize/Setup Configuration', 'setupConfigTab')
    .addSeparator()
    .addItem('2. Process All Invoices', 'runMainProcess')
    .addItem('3. Clear Tracker Data', 'clearTrackerData')
    .addToUi();
}

// --- CONFIGURATION MANAGEMENT ---
function setupConfigTab(isWebApp) {
  isWebApp = isWebApp === true;  // normalize to boolean; default is false when omitted
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. System Config
  let configSheet = ss.getSheetByName('System Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('System Config');
    const headers = [['Setting Name', 'Value / ID', 'Description']];
    configSheet.getRange(1, 1, 1, 3).setValues(headers).setFontWeight('bold').setBackground('#d9ead3');
    
    const defaultSettings = [
      ['FRG_HAULIER_ID', '1VyvVZ3BB2-tArldhuNKKBzfjU7FATVMJbtzhZP3P8pA', 'FRG Weekly Haulier Sheet ID'],
      ['GRM_HAULIER_ID', '13lvErboGKp-cd06XV1PvhB7zSrCEXXVm0POZlWPBZV4', 'GRM Weekly Haulier Sheet ID'],
      ['PYE_HAULIER_ID', '1AP_qyToLfhyuGihvfV3sUbLpj4WYx-wgAUVjtluHCoY', 'PYE Weekly Haulier Sheet ID'],
      ['FRG_ENABLED', 'YES', 'Enable or disable FRG lane (YES/NO)'],
      ['GRM_ENABLED', 'YES', 'Enable or disable GRM lane (YES/NO)'],
      ['PYE_ENABLED', 'YES', 'Enable or disable PYE lane (YES/NO)'],
      ['CRE_ROOT_FOLDER', '1BCbMyWUu0npRGIQvN1lRtrL5oBh1PdAi', 'Root Folder ID for CRE Invoices'],
      ['HB_ROOT_FOLDER', '1awkIvApl6iyExTnrNTmD1Pt1oM1d-2GO', 'Root Folder ID for HB Invoices'],
      ['SCH_ROOT_FOLDER', '', 'Root Folder ID for SCH Invoices (Optional)'],
      ['WERNER_ROOT_FOLDER', '', 'Root Folder ID for Werner Invoices (Optional)'],
      ['CRE_ENABLED', 'YES', 'Enable or disable CRE lane (YES/NO)'],
      ['HB_ENABLED', 'YES', 'Enable or disable HB lane (YES/NO)'],
      ['SCH_ENABLED', 'YES', 'Enable or disable SCH lane (YES/NO)'],
      ['WERNER_ENABLED', 'YES', 'Enable or disable WERNER lane (YES/NO)'],
      ['INVOICE_NUMBER_REGEX', 'Invoice\\s*(?:#|ID|Number)\\s*:?\\s*([A-Za-z0-9_-]+)', 'Regex used to extract invoice number from invoice sheets'],
      ['INVOICE_NUMBER_CELL', '', 'Optional fixed cell to check first (e.g. B3)'],
      ['INVOICE_NUMBER_FALLBACK', 'USE_FILENAME', 'USE_FILENAME or REQUIRE_EXTRACTED'],
      ['IGNORE_INVOICE_SHIFT2_WHEN_HAULIER_SHIFT1', 'YES', 'When YES, do not block processing if invoice shows Shift 2 and haulier shows Shift 1; log informational discrepancy only'],
      ['ARCHIVE_SHEET_PREFIX', 'Archive - ', 'Prefix used for archive snapshot sheet names'],
      ['ARCHIVE_DATE_FORMAT', 'yyyy-MM-dd', 'Date format appended to archive sheet names'],
      ['ARCHIVE_RETENTION_DAYS', '90', 'Delete archive snapshots older than this many days (0 disables pruning)'],
      ['TEMP_PROCESSING_FOLDER', '', 'Optional: Folder ID to store temp converted sheets (leave blank to use root)']
    ];
    configSheet.getRange(2, 1, defaultSettings.length, 3).setValues(defaultSettings);
    configSheet.setColumnWidth(1, 250); configSheet.setColumnWidth(2, 350); configSheet.setColumnWidth(3, 400);
  }

  ensureSystemConfigSettingRow_(
    configSheet,
    'IGNORE_INVOICE_SHIFT2_WHEN_HAULIER_SHIFT1',
    'YES',
    'When YES, do not block processing if invoice shows Shift 2 and haulier shows Shift 1; log informational discrepancy only'
  );

  // 2. GL Config (Now with Ignore Rules & Tolerance)
  let glSheet = ss.getSheetByName('GL Config');
  if (!glSheet) {
    glSheet = ss.insertSheet('GL Config');
    const glHeaders = [['Rule Type', 'Keyword / RDC', 'GL Account / Center Code / Value', 'Description']];
    glSheet.getRange(1, 1, 1, 4).setValues(glHeaders).setFontWeight('bold').setBackground('#fff2cc');
    
    const defaultGL = [
      ['Cost Center', 'FRG', '50001', 'Fredericksburg Cost Center Code'],
      ['Cost Center', 'GRM', '60001', 'Graham Cost Center Code'],
      ['Cost Center', 'PYE', '70001', 'Perryville Cost Center Code'],
      ['Category Mapping', 'toll', '471000, 47100002', 'Catches any line item containing "toll"'],
      ['Category Mapping', 'fuel', '471000, 47100099', 'Catches any line item containing "fuel"'],
      ['Category Mapping', 'others', '471000, 47100099', 'Used for TUs marked as Type O in Haulier Report'],
      ['Default Category', 'BASE', '471000, 47100001', 'Base freight applied to all other costs'],
      ['Ignore Rule', 'discount', '', 'Completely ignores any line item containing this word'],
      ['Ignore Rule', 'rebate', '', 'Completely ignores any line item containing this word'],
      ['Tolerance', 'MIN_COST', '0.01', 'Ignores any line item with an amount smaller than this value'],
      ['Tolerance', 'AMOUNT_MATCH_PCT', '0.02', 'Percent tolerance for invoice-vs-haulier amount match (2% = 0.02)'],
      ['Tolerance', 'AMOUNT_MATCH_HARD', '5.00', 'Hard-dollar tolerance for invoice-vs-haulier amount match'],
      ['Tolerance', 'ESTIMATE_ALERT_HARD', '200.00', 'Alert threshold for Expected Estimate vs Actual Amount variance']
    ];
    glSheet.getRange(2, 1, defaultGL.length, 4).setValues(defaultGL);
    glSheet.setColumnWidth(1, 150); glSheet.setColumnWidth(2, 150); glSheet.setColumnWidth(3, 220); glSheet.setColumnWidth(4, 350);
  }
  ensureGlToleranceRow_(glSheet, 'ESTIMATE_ALERT_HARD', '200.00', 'Alert threshold for Expected Estimate vs Actual Amount variance');
  
  // 3. RDC Aliases Config
  let rdcSheet = ss.getSheetByName('RDC Aliases');
  if (!rdcSheet) {
    rdcSheet = ss.insertSheet('RDC Aliases');
    rdcSheet.getRange(1, 1, 1, 2).setValues([['RDC Code', 'Aliases (Comma Separated)']]).setFontWeight('bold').setBackground('#cfe2f3');
    const defaultAliases = [
      ['FRG', 'FREDERICKSBURG, FRG, VA, LIDL VA, _VA'],
      ['GRM', 'GRAHAM, GRM, NC, LIDL NC, _NC, MEBANE'],
      ['PYE', 'PERRYVILLE, PYE, MD, LIDL MD, _MD, PER']
    ];
    rdcSheet.getRange(2, 1, defaultAliases.length, 2).setValues(defaultAliases);
    rdcSheet.setColumnWidth(1, 150); rdcSheet.setColumnWidth(2, 500);
  }
  
  // 4. Email Template Config
  let emailSheet = ss.getSheetByName('Email Template');
  if (!emailSheet) {
    emailSheet = ss.insertSheet('Email Template');
    emailSheet.getRange(1, 1, 1, 3).setValues([['Setting', 'Template Text', 'Available Variables']]).setFontWeight('bold').setBackground('#fce5cd');
    const defaultEmail = [
      ['Subject', 'Discrepancy Notice: Invoice {FileName}', '{FileName}, {CarrierName}'],
      ['Greeting', 'Hello {CarrierName} Team,\n\nThe following items are showing as discrepancies in our system. Please advise:', '{CarrierName}'],
      ['Outro', 'Thank you.', '']
    ];
    emailSheet.getRange(2, 1, defaultEmail.length, 3).setValues(defaultEmail);
    emailSheet.setColumnWidth(1, 150); emailSheet.setColumnWidth(2, 500); emailSheet.setColumnWidth(3, 200);
  }

  // 5. Header Aliases Config (Dynamic Column Mapping)
  let headerSheet = ss.getSheetByName('Header Config');
  if (!headerSheet) {
    headerSheet = ss.insertSheet('Header Config');
    headerSheet.getRange(1, 1, 1, 3).setValues([['Target Field', 'Column Name Aliases (Comma Separated)', 'Carrier (Blank = Global)']]).setFontWeight('bold').setBackground('#e6b8af');
    const defaultHeaders = [
      ['Date', 'date, pickup dt, delivery date', ''],
      ['TU', 'tu, mb number', ''],
      ['Store', 'store, dest location, dest city, destination name', ''],
      ['Tour', 'tour, route', ''],
      ['Miles', 'miles, total miles', ''],
      ['NY Pay', 'ny pay, ny, new york, borough fee, dhu $ amt, stp $ amt', ''],
      ['Tolls', 'toll, tol $ amt', ''],
      ['Total Cost', 'total $ amt, total cost, cost, total', ''],
      ['Shift', 'shift, shifts', ''],
      ['Type', 'delivery type, tour type, shift type', ''],
      ['Haulier Invoice', 'invoice, invoice #, invoice number', ''],
      ['Haulier Amount', 'amount, actual amount, invoice amount, total cost', ''],
      ['Haulier Estimate', 'dedicated estimation, deticated estimation, dedicated estimate, deticated estimate, estimated amount, estimate, est amount', '']
    ];
    headerSheet.getRange(2, 1, defaultHeaders.length, 3).setValues(defaultHeaders);
    headerSheet.setColumnWidth(1, 150); headerSheet.setColumnWidth(2, 500); headerSheet.setColumnWidth(3, 200);
  }

  // 6. Carrier Processor Config
  let carrierCfgSheet = ss.getSheetByName('Carrier Config');
  if (!carrierCfgSheet) {
    carrierCfgSheet = ss.insertSheet('Carrier Config');
    carrierCfgSheet.getRange(1, 1, 1, 4).setValues([['Carrier Name', 'Processor Type', 'Invoice # Pattern (Optional)', 'Notes']]).setFontWeight('bold').setBackground('#c9daf8');
    const defaultCarrierConfig = [
      ['CRE', 'CRE', '', 'Use CRE parser'],
      ['HB', 'HB', '', 'Use HB parser'],
      ['SCH', 'SCH', '', 'Use SCH parser'],
      ['WERNER', 'HB', '', 'Werner uses HB-style structure']
    ];
    carrierCfgSheet.getRange(2, 1, defaultCarrierConfig.length, 4).setValues(defaultCarrierConfig);
    carrierCfgSheet.setColumnWidth(1, 180); carrierCfgSheet.setColumnWidth(2, 150); carrierCfgSheet.setColumnWidth(3, 260); carrierCfgSheet.setColumnWidth(4, 360);
  }
  
  // 7. Output Routing Config
  let routingSheet = ss.getSheetByName('Output Routing');
  if (!routingSheet) {
    routingSheet = ss.insertSheet('Output Routing');
    routingSheet.getRange(1, 1, 1, 4).setValues([['Data Type', 'Target Sheet Name', 'Target Spreadsheet ID (blank = same workbook)', 'Enabled']]).setFontWeight('bold').setBackground('#d5a6bd');
    const defaultRouting = [
      ['Additional Costs', 'Additonal Costs', '', 'YES'],
      ['Discrepancy Tracker', 'Discrepancy Tracker', '', 'YES'],
      ['Backhaul Credits', 'Backhaul Credits', '', 'NO']
    ];
    routingSheet.getRange(2, 1, defaultRouting.length, 4).setValues(defaultRouting);
    routingSheet.setColumnWidth(1, 200); routingSheet.setColumnWidth(2, 220); routingSheet.setColumnWidth(3, 380); routingSheet.setColumnWidth(4, 80);
  }

  // 8. Rule Config
  let ruleSheet = ss.getSheetByName('Rule Config');
  if (!ruleSheet) {
    ruleSheet = ss.insertSheet('Rule Config');
  }
  ensureRuleConfigSchema_(ruleSheet);

  // Ensure operations sheet used by estimate-vs-actual alerts exists even before first alert row.
  ensureOperationSheet_(ss, 'Estimate Variance', [
    'RDC',
    'Date',
    'TU',
    'Carrier',
    'Invoice Number',
    'Expected Estimate',
    'Actual Amount',
    'Difference (Actual-Expected)',
    'Alert Threshold',
    'Abs Difference',
    'Status'
  ], '#fce5cd');

  // 9. GL Code Templates (for Invoice Finalization feature)
  let glTemplatesSheet = ss.getSheetByName('GL Code Templates');
  if (!glTemplatesSheet) {
    glTemplatesSheet = ss.insertSheet('GL Code Templates');
    glTemplatesSheet.getRange(1, 1, 1, 4).setValues([['Template Name', 'GL Code / Cost Center String', 'Description', 'Quick Select (YES/NO)']]).setFontWeight('bold').setBackground('#d9d9d9');
    const defaultTemplates = [
      ['Base Freight', '471000, 47100001', 'Default store delivery cost', 'YES'],
      ['Tolls', '471000, 47100002', 'Toll charges', 'YES'],
      ['Fuel Surcharge', '471000, 47100099', 'Fuel and miscellaneous charges', 'YES'],
      ['Transfer', '471000, 47100004', 'Backhaul/transfer charges', 'NO'],
      ['Import Savings', '471003', 'CRE import carve-out', 'YES'],
      ['Self Pickup', '360100', 'CRE self-pickup carve-out', 'YES']
    ];
    glTemplatesSheet.getRange(2, 1, defaultTemplates.length, 4).setValues(defaultTemplates);
    glTemplatesSheet.setColumnWidth(1, 150); glTemplatesSheet.setColumnWidth(2, 280); glTemplatesSheet.setColumnWidth(3, 300); glTemplatesSheet.setColumnWidth(4, 120);
  }

  // Ensure operations sheet for invoice finalization state
  ensureOperationSheet_(ss, 'Invoice Finalization', [
    'Invoice File Name',
    'RDC',
    'Invoice Number',
    'Total Amount',
    'Saved At',
    'GL Code Config (JSON)',
    'Status',
    'Sheet ID',
    'Target Folder ID',
    'Cost Summary JSON',
    'TMST JSON'
  ], '#b6d7a8');
  
  // 10. Additional Cost Writeback Config
  let addlCostCfgSheet = ss.getSheetByName('Addl Cost Config');
  if (!addlCostCfgSheet) {
    addlCostCfgSheet = ss.insertSheet('Addl Cost Config');
    addlCostCfgSheet.getRange(1, 1, 1, 6).setValues([[
      'Type', 'Carrier', 'RDC', 'Invoice Keyword / Sheet Name', 'Haulier Column / Date Column', 'Enabled'
    ]]).setFontWeight('bold').setBackground('#d0e4ff');
    const defaultAddlCostRows = [
      ['Sheet',   'CRE', 'PYE', 'Additional Cost',               'Week Ending',        'YES'],
      ['Mapping', 'CRE', 'PYE', 'additonal single-temp',         'Additional Trailers', 'YES'],
      ['Mapping', 'CRE', 'PYE', 'additonal multi-temp',          'Additional Trailers', 'YES'],
      ['Mapping', 'CRE', 'PYE', "36' reefer",                    "36' Trailers",        'YES'],
      ['Mapping', 'CRE', 'PYE', "48' reefer multi-temp",         "48' Trailers",        'YES'],
      ['Mapping', 'CRE', 'PYE', "48' reefer single temp",        "48' Trailers",        'YES'],
      ['Mapping', 'CRE', 'PYE', 'trailer washout',               'Trailer Washouts',    'YES'],
      ['Mapping', 'CRE', 'PYE', 'third structure tax',           'Third Structure Tax', 'YES'],
      ['Mapping', 'CRE', 'PYE', 'dry van',                       'Dry Vans',            'YES'],
      ['Mapping', 'CRE', 'PYE', 'scale',                         'Scales',              'YES'],
      ['Mapping', 'CRE', 'PYE', 'shared risk',                   'Shared Risk Credit',  'YES'],
    ];
    addlCostCfgSheet.getRange(2, 1, defaultAddlCostRows.length, 6).setValues(defaultAddlCostRows);
    addlCostCfgSheet.setColumnWidth(1, 90); addlCostCfgSheet.setColumnWidth(2, 90);
    addlCostCfgSheet.setColumnWidth(3, 70); addlCostCfgSheet.setColumnWidth(4, 260);
    addlCostCfgSheet.setColumnWidth(5, 220); addlCostCfgSheet.setColumnWidth(6, 80);
  }

  if (!isWebApp) {
    SpreadsheetApp.getUi().alert('Configuration setup complete. Please review the 10 Configuration tabs at the bottom of your sheet.');
  }
}

function isCarrierEnabled_(config, carrierName) {
  const key = String(carrierName || '').trim().toUpperCase() + '_ENABLED';
  const raw = String(config[key] === undefined ? 'YES' : config[key]).trim().toUpperCase();
  return !(raw === 'NO' || raw === 'FALSE' || raw === '0' || raw === 'OFF' || raw === 'DISABLED');
}

function ensureGlToleranceRow_(glSheet, toleranceKey, defaultValue, description) {
  if (!glSheet) return;
  const data = glSheet.getDataRange().getValues();
  const target = String(toleranceKey || '').trim().toUpperCase();
  if (!target) return;

  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][0] || '').trim();
    const key = String(data[i][1] || '').trim().toUpperCase();
    if (type === 'Tolerance' && key === target) return;
  }

  glSheet.appendRow(['Tolerance', target, defaultValue, description || '']);
}

function ensureSystemConfigSettingRow_(configSheet, settingName, defaultValue, description) {
  if (!configSheet) return;
  const data = configSheet.getDataRange().getValues();
  const target = String(settingName || '').trim().toUpperCase();
  if (!target) return;

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim().toUpperCase();
    if (key === target) return;
  }

  configSheet.appendRow([settingName, defaultValue, description || '']);
}

function ensureOperationSheet_(ss, sheetName, headers, headerColor) {
  if (!ss || !sheetName) return;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0 && headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (headerColor) {
      sheet.getRange(1, 1, 1, headers.length).setBackground(headerColor);
    }
  }
}

function getRuleCatalog_() {
  return [
    {
      key: 'HB_BLANK_SHIFT_TOUR2_EXPECT_FIRST_SECOND',
      type: 'SYSTEM',
      enabled: 'YES',
      carrierScope: 'HB, WERNER',
      condition1: '',
      condition2: '',
      issueLabel: 'HB Blank Shift / Tour Mapping Mismatch',
      expectedValue: 'Expected Shift FIRST + Tour SECOND',
      actualSource: 'haulier.shift|haulier.tour',
      description: 'For HB-style carriers, blank invoice shift plus Tour 2 suffix expects haulier Shift FIRST and Tour SECOND.'
    },
    {
      key: 'DELIVERY_TYPE_CNC_WITH_ASSIGNED_SHIFT',
      type: 'SYSTEM',
      enabled: 'YES',
      carrierScope: '',
      condition1: '',
      condition2: '',
      issueLabel: 'Route recorded as Carrier Cancel',
      expectedValue: 'No assigned carrier shift when delivery type is CNC',
      actualSource: 'invoice.shift',
      description: 'Flag a discrepancy when haulier Delivery Type is CNC but the carrier invoice still has a shift assigned.'
    }
  ];
}

function getDefaultRuleConfigRows_() {
  return getRuleCatalog_().map(rule => [
    rule.key,
    rule.type,
    rule.enabled,
    rule.carrierScope,
    rule.condition1,
    rule.condition2,
    rule.issueLabel,
    rule.expectedValue,
    rule.actualSource,
    rule.description
  ]);
}

function normalizeRuleRow_(row) {
  const raw = Array.isArray(row) ? row : [];
  const catalog = {};
  getRuleCatalog_().forEach(rule => catalog[rule.key] = rule);

  const first = String(raw[0] || '').trim().toUpperCase();
  const second = String(raw[1] || '').trim().toUpperCase();
  const isExpandedSchema = second === 'SYSTEM' || second === 'CUSTOM' || raw.length >= 10;

  if (isExpandedSchema) {
    return [
      first,
      second || 'CUSTOM',
      String(raw[2] || 'YES').trim().toUpperCase() || 'YES',
      String(raw[3] || '').trim(),
      String(raw[4] || '').trim(),
      String(raw[5] || '').trim(),
      String(raw[6] || '').trim(),
      String(raw[7] || '').trim(),
      String(raw[8] || '').trim(),
      String(raw[9] || '').trim()
    ];
  }

  const meta = catalog[first] || {};
  return [
    first,
    meta.type || 'SYSTEM',
    String(raw[1] || meta.enabled || 'YES').trim().toUpperCase() || 'YES',
    String(raw[2] || meta.carrierScope || '').trim(),
    meta.condition1 || '',
    meta.condition2 || '',
    meta.issueLabel || first,
    meta.expectedValue || '',
    meta.actualSource || '',
    String(raw[3] || meta.description || '').trim()
  ];
}

function ensureRuleConfigSchema_(ruleSheet) {
  if (!ruleSheet) return;

  const headers = [[
    'Rule Key',
    'Rule Type',
    'Enabled',
    'Carrier Scope (Blank = All)',
    'Condition 1',
    'Condition 2',
    'Issue Label',
    'Expected Value / Guidance',
    'Actual Source',
    'Description'
  ]];

  ruleSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold').setBackground('#d9d2e9');
  ruleSheet.setColumnWidth(1, 280);
  ruleSheet.setColumnWidth(2, 110);
  ruleSheet.setColumnWidth(3, 90);
  ruleSheet.setColumnWidth(4, 220);
  ruleSheet.setColumnWidth(5, 260);
  ruleSheet.setColumnWidth(6, 260);
  ruleSheet.setColumnWidth(7, 230);
  ruleSheet.setColumnWidth(8, 250);
  ruleSheet.setColumnWidth(9, 180);
  ruleSheet.setColumnWidth(10, 420);

  const lastRow = ruleSheet.getLastRow();
  if (lastRow <= 1) {
    const defaults = getDefaultRuleConfigRows_();
    if (defaults.length) ruleSheet.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
    return;
  }

  const data = ruleSheet.getDataRange().getValues();
  const normalizedRows = [];
  for (let i = 1; i < data.length; i++) {
    const normalized = normalizeRuleRow_(data[i]);
    if (normalized.some(value => String(value || '').trim())) normalizedRows.push(normalized);
  }

  if (lastRow > 1) {
    ruleSheet.getRange(2, 1, lastRow - 1, Math.max(ruleSheet.getLastColumn(), headers[0].length)).clearContent();
  }
  if (normalizedRows.length) {
    ruleSheet.getRange(2, 1, normalizedRows.length, normalizedRows[0].length).setValues(normalizedRows);
  }
}

// --- DATA FETCHERS ---

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('System Config');
  if (!configSheet) throw new Error('System Config tab not found. Please run Initialization first.');
  const data = configSheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) config[data[i][0]] = data[i][1];
  return config;
}

function getGlConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const glSheet = ss.getSheetByName('GL Config');
  if (!glSheet) return null;
  const data = glSheet.getDataRange().getValues();
  
  const config = {
    costCenters: {},
    categories: [],
    ignoreKeywords: [],
    tolerance: 0.01,
    tolerances: { MIN_COST: 0.01, AMOUNT_MATCH_PCT: 0.02, AMOUNT_MATCH_HARD: 5.00 },
    backhaulCredits: []
  };
  
  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][0]).trim();
    const key = String(data[i][1]).trim().toLowerCase();
    const value = String(data[i][2]).trim();
    
    if (!type || (!key && type !== 'Tolerance')) continue;
    
    if (type === 'Cost Center') config.costCenters[key.toUpperCase()] = value;
    else if (type === 'Category Mapping') config.categories.push({ keyword: key, glPrefix: value });
    else if (type === 'Default Category') config.defaultCategory = value;
    else if (type === 'Ignore Rule') config.ignoreKeywords.push(key);
    else if (type === 'Tolerance') {
      const tolKey = String(data[i][1] || '').trim().toUpperCase();
      const tolVal = parseFloat(value);
      if (tolKey) config.tolerances[tolKey] = isNaN(tolVal) ? value : tolVal;
      if (tolKey === 'MIN_COST') config.tolerance = isNaN(tolVal) ? 0.01 : tolVal;
    }
    else if (type === 'Backhaul Credit') config.backhaulCredits.push({ keyword: key, glPrefix: value || 'BACKHAUL' });
  }
  return config;
}

function getRdcAliases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RDC Aliases');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const aliases = {};
  for (let i = 1; i < data.length; i++) {
    const rdc = String(data[i][0]).trim();
    if (!rdc) continue;
    aliases[rdc] = String(data[i][1]).toUpperCase().split(',').map(s => s.trim());
  }
  return aliases;
}

function getEmailTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Email Template');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const template = {};
  for (let i = 1; i < data.length; i++) {
    template[String(data[i][0]).trim()] = String(data[i][1]);
  }
  return template;
}

function getHeaderAliases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Header Config');
  if (!sheet) return { global: {}, byCarrier: {} };
  const data = sheet.getDataRange().getValues();
  const aliases = { global: {}, byCarrier: {} };
  for (let i = 1; i < data.length; i++) {
    const field = String(data[i][0]).trim();
    if (!field) continue;
    const aliasList = String(data[i][1]).toLowerCase().split(',').map(s => s.trim()).filter(Boolean);
    const carrier = String(data[i][2] || '').trim().toUpperCase();
    if (!carrier) {
      aliases.global[field] = aliasList;
      continue;
    }
    if (!aliases.byCarrier[carrier]) aliases.byCarrier[carrier] = {};
    aliases.byCarrier[carrier][field] = aliasList;
  }
  return aliases;
}

function getCarrierConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Carrier Config');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < data.length; i++) {
    const carrier = String(data[i][0] || '').trim().toUpperCase();
    if (!carrier) continue;
    const processorType = String(data[i][1] || '').trim().toUpperCase() || 'HB';
    const invoicePattern = String(data[i][2] || '').trim();
    config[carrier] = { processorType: processorType, invoicePattern: invoicePattern };
  }

  return config;
}

function getRuleConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Rule Config');
  if (!sheet) return { __customRules: [] };
  const data = sheet.getDataRange().getValues();
  const rules = { __customRules: [] };

  for (let i = 1; i < data.length; i++) {
    const row = normalizeRuleRow_(data[i]);
    const key = String(row[0] || '').trim().toUpperCase();
    if (!key) continue;
    const rule = {
      key: key,
      type: String(row[1] || 'SYSTEM').trim().toUpperCase() || 'SYSTEM',
      enabled: !(/^(NO|FALSE|0|OFF|DISABLED)$/i).test(String(row[2] || 'YES').trim()),
      carriers: String(row[3] || '').toUpperCase().split(',').map(s => s.trim()).filter(Boolean),
      condition1: String(row[4] || '').trim(),
      condition2: String(row[5] || '').trim(),
      issueLabel: String(row[6] || '').trim(),
      expectedValue: String(row[7] || '').trim(),
      actualSource: String(row[8] || '').trim(),
      description: String(row[9] || '').trim()
    };

    if (rule.type === 'CUSTOM') rules.__customRules.push(rule);
    else rules[key] = rule;
  }

  return rules;
}

function isRuleEnabled_(rules, ruleKey, carrierName) {
  const key = String(ruleKey || '').trim().toUpperCase();
  const carrier = String(carrierName || '').trim().toUpperCase();
  if (!key) return true;
  if (!rules || !rules[key]) return true;

  const rule = rules[key];
  if (!rule.enabled) return false;
  if (!rule.carriers || rule.carriers.length === 0) return true;
  return rule.carriers.indexOf(carrier) !== -1;
}

function parseRuleCondition_(conditionText) {
  const raw = String(conditionText || '').trim();
  if (!raw) return null;

  const operators = ['NOT_CONTAINS', 'NOT_EQUALS', 'STARTS_WITH', 'ENDS_WITH', 'NOT_BLANK', 'EQUALS', 'CONTAINS', 'BLANK'];
  for (let i = 0; i < operators.length; i++) {
    const operator = operators[i];
    const match = raw.match(new RegExp('^(.+?)\\s+' + operator + '(?:\\s+(.+))?$', 'i'));
    if (match) {
      return {
        left: String(match[1] || '').trim(),
        operator: operator,
        right: String(match[2] || '').trim()
      };
    }
  }

  return { invalid: true, raw: raw };
}

function resolveRuleToken_(token, context) {
  const key = String(token || '').trim();
  if (!key) return '';
  if (context && Object.prototype.hasOwnProperty.call(context, key)) return context[key];
  return key;
}

function resolveRuleDisplayValue_(source, context) {
  const raw = String(source || '').trim();
  if (!raw) return '';

  return raw.split('|').map(part => {
    const token = String(part || '').trim();
    if (!token) return '';
    const value = resolveRuleToken_(token, context);
    if (token.indexOf('.') !== -1) {
      return token.split('.').pop() + "='" + String(value || '') + "'";
    }
    return String(value || '');
  }).filter(Boolean).join(' | ');
}

function normalizeRuleValue_(value) {
  return String(value === undefined || value === null ? '' : value).trim().toUpperCase();
}

function evaluateRuleCondition_(conditionText, context) {
  const parsed = parseRuleCondition_(conditionText);
  if (!parsed) return { valid: true, passed: true };
  if (parsed.invalid) return { valid: false, passed: false, error: `Invalid condition: ${parsed.raw}` };

  const leftRaw = resolveRuleToken_(parsed.left, context);
  const left = String(leftRaw === undefined || leftRaw === null ? '' : leftRaw);
  const leftNorm = normalizeRuleValue_(left);
  const operator = parsed.operator;

  if (operator === 'BLANK') return { valid: true, passed: leftNorm === '' };
  if (operator === 'NOT_BLANK') return { valid: true, passed: leftNorm !== '' };

  const rightRaw = resolveRuleToken_(parsed.right, context);
  const right = String(rightRaw === undefined || rightRaw === null ? '' : rightRaw);
  const rightNorm = normalizeRuleValue_(right);

  if (operator === 'EQUALS') return { valid: true, passed: leftNorm === rightNorm };
  if (operator === 'NOT_EQUALS') return { valid: true, passed: leftNorm !== rightNorm };
  if (operator === 'CONTAINS') return { valid: true, passed: leftNorm.indexOf(rightNorm) !== -1 };
  if (operator === 'NOT_CONTAINS') return { valid: true, passed: leftNorm.indexOf(rightNorm) === -1 };
  if (operator === 'STARTS_WITH') return { valid: true, passed: leftNorm.indexOf(rightNorm) === 0 };
  if (operator === 'ENDS_WITH') return { valid: true, passed: rightNorm !== '' && leftNorm.slice(-rightNorm.length) === rightNorm };

  return { valid: false, passed: false, error: `Unsupported operator: ${operator}` };
}

function runCustomRules_(rules, context, discrepancyData, rdcName, date, tuNumber) {
  const customRules = (rules && rules.__customRules) ? rules.__customRules : [];
  if (!customRules.length) return;

  const carrierKey = normalizeRuleValue_(context['carrier.name'] || '');

  customRules.forEach(rule => {
    if (!rule || !rule.enabled) return;
    if (rule.carriers && rule.carriers.length && rule.carriers.indexOf(carrierKey) === -1) return;

    const conditionOne = evaluateRuleCondition_(rule.condition1, context);
    const conditionTwo = evaluateRuleCondition_(rule.condition2, context);
    if (!conditionOne.valid || !conditionTwo.valid) return;
    if (!conditionOne.passed || !conditionTwo.passed) return;

    const expectedValue = rule.expectedValue || rule.condition1 || '';
    const parsedCondition = parseRuleCondition_(rule.condition1);
    const defaultActualSource = parsedCondition && !parsedCondition.invalid ? parsedCondition.left : '';
    const actualValue = rule.actualSource
      ? resolveRuleDisplayValue_(rule.actualSource, context)
      : resolveRuleDisplayValue_(defaultActualSource, context);

    discrepancyData.push([
      rdcName,
      date,
      tuNumber,
      expectedValue,
      actualValue,
      rule.issueLabel || rule.key || 'Custom Rule Triggered'
    ]);
  });
}

function getHeaderKeywords_(aliasesMap, fieldName, carrierName) {
  const globalKeywords = (aliasesMap.global && aliasesMap.global[fieldName]) ? aliasesMap.global[fieldName] : [];
  if (!carrierName) return globalKeywords;
  const carrierKey = String(carrierName).trim().toUpperCase();
  const carrierKeywords = (aliasesMap.byCarrier && aliasesMap.byCarrier[carrierKey] && aliasesMap.byCarrier[carrierKey][fieldName])
    ? aliasesMap.byCarrier[carrierKey][fieldName]
    : [];

  // Carrier-specific aliases should have priority, but still fall back to global aliases.
  return carrierKeywords.concat(globalKeywords);
}

function matchHeader(headerStr, fieldName, aliasesMap, carrierName = '', exactOnly = false) {
  // Backward compatibility: old call shape was (headerStr, fieldName, aliasesMap, exactOnly).
  if (typeof carrierName === 'boolean') {
    exactOnly = carrierName;
    carrierName = '';
  }
  const keywords = getHeaderKeywords_(aliasesMap, fieldName, carrierName);
  
  // Pass 1: Prioritize exact matches (prevents partial string hijacking)
  for (let kw of keywords) {
    if (headerStr === kw) return true;
  }
  
  // If we are strictly running an exact match pass, abort here
  if (exactOnly) return false;
  
  // Pass 2: Permissive partial matches
  for (let kw of keywords) {
    // Strict exclusion to prevent partial word matches for very short acronyms
    if (kw === 'tu' && (headerStr.includes('status') || headerStr.includes('return'))) continue;
    if (kw === 'ny' && headerStr.includes('company')) continue;
    
    if (headerStr.includes(kw)) {
       // Enforce word boundaries for short acronyms to avoid substring false positives
       if (kw === 'tu' || kw === 'ny') {
          if (new RegExp('\\b' + kw + '\\b', 'i').test(headerStr)) return true;
       } else {
          return true;
       }
    }
  }
  return false;
}

function rowHasHeader(rowArray, fieldName, aliasesMap, carrierName = '') {
  // Backward compatibility: old call shape was (rowArray, fieldName, aliasesMap).
  if (typeof carrierName === 'boolean') carrierName = '';
  for (let i = 0; i < rowArray.length; i++) {
    // Flatten multiple spaces into a single space for consistent matching
    const h = String(rowArray[i]).toLowerCase().trim().replace(/\s+/g, ' ');
    if (matchHeader(h, fieldName, aliasesMap, carrierName, false)) return true;
  }
  return false;
}

// --- MAIN PROCESS ---
var _extractOnlyMode = false;
var _pendingInvoices = [];

function runMainProcessExtractOnly(isWebApp = false) {
  _extractOnlyMode = true;
  _pendingInvoices = [];
  try {
    return runMainProcess(isWebApp);
  } finally {
    _extractOnlyMode = false;
  }
}

function finalizePendingInvoices(isWebApp = false) {
  let ui = null;
  if (!isWebApp) {
    ui = SpreadsheetApp.getUi();
  }

  try {
    ensureInvoiceFinalizationSchema_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const finSheet = ss.getSheetByName('Invoice Finalization');
    if (!finSheet || finSheet.getLastRow() < 2) {
      if (ui) ui.alert('No pending invoices to finalize. Run extract first.');
      return 'WARNING: No pending invoices to finalize.';
    }

    const data = finSheet.getDataRange().getValues();
    const header = data[0];
    const idxFile = header.indexOf('Invoice File Name');
    const idxRdc = header.indexOf('RDC');
    const idxStatus = header.indexOf('Status');
    const idxGl = header.indexOf('GL Code Config (JSON)');
    const idxSheetId = header.indexOf('Sheet ID');
    const idxFolderId = header.indexOf('Target Folder ID');
    const idxCostSummary = header.indexOf('Cost Summary JSON');
    const idxTmst = header.indexOf('TMST JSON');

    if (idxFile < 0 || idxRdc < 0 || idxStatus < 0 || idxGl < 0 || idxSheetId < 0) {
      throw new Error('Invoice Finalization sheet is missing required columns. Run Initialize Config and retry.');
    }

    const pendingRows = [];
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][idxStatus] || '').toUpperCase();
      if (status === 'PENDING' || status === 'EDITED') pendingRows.push({ rowNum: i + 1, row: data[i] });
    }

    const totalPending = pendingRows.length;
    if (totalPending === 0) {
      if (ui) ui.alert('No pending invoices to finalize.');
      return 'WARNING: No pending invoices to finalize.';
    }

    if (ui) ui.alert('Finalization started. Applying GL codes and generating PDFs...');
    Logger.log(`--- STARTING FINALIZATION FOR ${totalPending} INVOICE(S) ---`);

    const glConfig = getGlConfig();
    let processedCount = 0;

    pendingRows.forEach(function(item) {
      try {
        const row = item.row;
        const fileName = String(row[idxFile] || '');
        const rdcName = String(row[idxRdc] || 'UNKNOWN');
        const sheetId = idxSheetId >= 0 ? String(row[idxSheetId] || '') : '';
        const folderId = idxFolderId >= 0 ? String(row[idxFolderId] || '') : '';
        const glMap = idxGl >= 0 ? JSON.parse(String(row[idxGl] || '{}')) : {};
        const costSummary = idxCostSummary >= 0 ? JSON.parse(String(row[idxCostSummary] || '[]')) : [];
        const tmstData = idxTmst >= 0 ? JSON.parse(String(row[idxTmst] || '[]')) : [];

        if (!sheetId) throw new Error('Missing Sheet ID in pending queue. Re-run extraction.');
        const invoiceSS = SpreadsheetApp.openById(sheetId);
        const targetFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();

        if (glMap && Object.keys(glMap).length > 0) {
          applyManualGLCodesAndExportPDF(invoiceSS, fileName, glMap, targetFolder);
        } else {
          applyGLCodesAndExportPDF(invoiceSS, fileName, rdcName, costSummary, glConfig, targetFolder, tmstData);
        }

        finSheet.getRange(item.rowNum, idxStatus + 1).setValue('FINALIZED');
        processedCount++;
      } catch (e) {
        Logger.log(`[ERROR] Failed to finalize row ${item.rowNum}: ${e.message}`);
      }
    });

    _pendingInvoices = [];
    const msg = `SUCCESS: Finalization complete. Processed ${processedCount}/${totalPending} invoice(s).`;
    if (ui) ui.alert(msg);
    return msg;
  } catch (error) {
    Logger.log(`CRITICAL ERROR DURING FINALIZATION: ${error.message}`);
    if (ui) ui.alert('Error during finalization: ' + error.message);
    return 'ERROR: ' + error.message;
  }
}

function runMainProcess(isWebApp = false) {
  let ui = null;
  if (!isWebApp) {
    ui = SpreadsheetApp.getUi();
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    if (ui) ui.alert(_extractOnlyMode ? 'Extraction started. Please wait...' : 'Processing Started. This may take a few minutes. Please wait...');
    Logger.log(_extractOnlyMode ? '--- STARTING INVOICE EXTRACTION ---' : '--- STARTING INVOICE PROCESSING ---');

    const config = getConfig();
    const glConfig = getGlConfig();
    const rdcAliases = getRdcAliases();
    const emailTemplate = getEmailTemplate();
    const headerAliases = getHeaderAliases();
    const carrierConfig = getCarrierConfig();
    _runtimeRuleConfig = getRuleConfig();
    _runtimeSystemConfig = config || {};
    _estimateDebugLogCount = 0;

    Logger.log('Fetching Haulier Data...');
    const haulierData = {
      'FRG': fetchHaulierData(config.FRG_HAULIER_ID, headerAliases, 'FRG'),
      'GRM': fetchHaulierData(config.GRM_HAULIER_ID, headerAliases, 'GRM'),
      'PYE': fetchHaulierData(config.PYE_HAULIER_ID, headerAliases, 'PYE')
    };

    const masterData = [];
    const additionalCostsData = [];
    const discrepancyData = [];
    const tmstData = [];
    const estimateVarianceData = [];
    const haulierUpdates = { 'FRG': [], 'GRM': [], 'PYE': [] };
    const addlCostWritebackConfig = getAddlCostWritebackConfig_();
    const addlCostWritebacks = {}; // keyed by RDC, array of { carrierName, weekEndDate, buckets }

    // Dynamically load ALL Carriers based on System Config (_ROOT_FOLDER suffix)
    const carriers = [];
    for (let key in config) {
      if (key.endsWith('_ROOT_FOLDER') && config[key]) {
        const carrierName = key.replace('_ROOT_FOLDER', '');
        carriers.push({ name: carrierName, rootId: config[key], enabled: isCarrierEnabled_(config, carrierName) });
      }
    }

    carriers.forEach(function(carrier) {
      if (!carrier.enabled) {
        Logger.log(`--- Skipping disabled carrier: ${carrier.name} ---`);
        return;
      }
      Logger.log(`\n--- Processing Carrier: ${carrier.name} ---`);
      let rootFolder;
      try { rootFolder = DriveApp.getFolderById(carrier.rootId); } catch (e) { return; }

      searchAndSortInvoices(rootFolder, carrier.name, haulierData, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, glConfig, rdcAliases, emailTemplate, headerAliases, carrierConfig, config, haulierUpdates, addlCostWritebackConfig, addlCostWritebacks);
    });

    Logger.log('--- EXECUTING HAULIER REPORT WRITE-BACKS ---');
    for (let rdc in haulierUpdates) {
      if (haulierUpdates[rdc].length > 0 && haulierData[rdc] && haulierData[rdc].spreadsheetId) {
         applyUpdatesToHaulier(haulierData[rdc], haulierUpdates[rdc]);
      }
    }

    Logger.log('--- EXECUTING ADDITIONAL COST WRITE-BACKS ---');
    for (let rdc in addlCostWritebacks) {
      if (addlCostWritebacks[rdc].length > 0 && haulierData[rdc] && haulierData[rdc].spreadsheetId) {
        applyAddlCostWritebacks_(haulierData[rdc].spreadsheetId, rdc, addlCostWritebacks[rdc], addlCostWritebackConfig);
      }
    }

    archiveAndClearSheets(ss);

    writeDataToSheet(ss, 'Master Input', masterData, 9);
    writeDataToSheet(ss, 'TMST', tmstData, 15);
    writeDataToSheet(ss, 'Additonal Costs', additionalCostsData, 5);
    writeDataToSheet(ss, 'Discrepancy Tracker', discrepancyData, 6);
    writeDataToSheet(ss, 'Estimate Variance', estimateVarianceData, 11);

    if (masterData.length === 0 && additionalCostsData.length === 0) {
      if (ui) ui.alert('Finished running, but no data was extracted. Files may have already been processed.');
      return 'WARNING: Finished running, but no data was extracted. Files may have already been processed.';
    }

    if (_extractOnlyMode) {
      const extractMsg = `SUCCESS: Extraction complete. ${_pendingInvoices.length} invoice(s) are ready for review/finalization.`;
      if (ui) ui.alert(extractMsg);
      return extractMsg;
    }

    if (ui) ui.alert('Processing Complete! Data added, GL Codes applied (if valid), drafts created for discrepancies, and files sorted.');
    return 'SUCCESS: Processing Complete! Data added, GL Codes applied, drafts created, and files sorted.';
  } catch (error) {
    Logger.log(`CRITICAL ERROR: ${error.message}`);
    if (ui) ui.alert('Error during processing: ' + error.message);
    return 'ERROR: ' + error.message;
  }
}

// --- ADDITIONAL COST WRITEBACK HELPERS ---

function getAddlCostWritebackConfig_() {
  const config = { mappings: [], sheetConfigs: {} };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Addl Cost Config');
    if (!sheet) return config;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const type    = String(row[0] || '').trim().toLowerCase();
      const carrier = String(row[1] || '').trim().toUpperCase();
      const rdc     = String(row[2] || '').trim().toUpperCase();
      const col4    = String(row[3] || '').trim();
      const col5    = String(row[4] || '').trim();
      const enabled = String(row[5] || 'YES').trim().toUpperCase();
      if (enabled === 'NO' || !col4) continue;
      if (type === 'mapping') {
        config.mappings.push({ carrier, rdc, keyword: col4.toLowerCase(), targetLabel: col5 });
      } else if (type === 'sheet') {
        config.sheetConfigs[`${carrier}|${rdc}`] = { sheetName: col4, dateColumn: col5 };
      }
    }
  } catch(e) {
    Logger.log(`[ADDL COST CONFIG] Error reading: ${e.message}`);
  }
  return config;
}

function parseCreInvoiceWeekEnd_(invoiceSS) {
  try {
    for (let s = 0; s < invoiceSS.getSheets().length; s++) {
      if (!invoiceSS.getSheets()[s].getName().toLowerCase().includes('summary')) continue;
      const data = invoiceSS.getSheets()[s].getDataRange().getValues();
      for (let i = 0; i < Math.min(data.length, 25); i++) {
        for (let j = 0; j < data[i].length; j++) {
          const cell = String(data[i][j] || '');
          const m = cell.match(/invoice\s+period[:\s]*(.*?)\s+through\s+([\d/]+)/i);
          if (m) {
            const parts = m[2].trim().split('/');
            if (parts.length === 3) {
              return new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
            }
          }
        }
      }
    }
  } catch(e) {}
  return null;
}

function collectAddlCostBuckets_(fileCostSummary, carrierName, rdcName, addlCostConfig) {
  const buckets = {};
  if (!addlCostConfig || !addlCostConfig.mappings || !addlCostConfig.mappings.length) return buckets;
  const cStr = String(carrierName || '').toUpperCase().trim();
  const rStr = String(rdcName || '').toUpperCase().trim();
  const relevant = addlCostConfig.mappings.filter(m => (!m.carrier || m.carrier === cStr) && (!m.rdc || m.rdc === rStr));
  if (!relevant.length) return buckets;
  (fileCostSummary || []).forEach(item => {
    const descLower = String(item.desc || '').toLowerCase();
    for (const m of relevant) {
      if (descLower.includes(m.keyword)) {
        const amt = parseFloat(item.amount) || 0;
        buckets[m.targetLabel] = (buckets[m.targetLabel] || 0) + amt;
        break; // first matching mapping wins
      }
    }
  });
  return buckets;
}

function applyAddlCostWritebacks_(haulierSsId, rdcName, addlCostEntries, addlCostConfig) {
  if (!haulierSsId || !addlCostEntries || !addlCostEntries.length) return;
  try {
    const ss = SpreadsheetApp.openById(haulierSsId);
    for (const entry of addlCostEntries) {
      const key = `${entry.carrierName}|${rdcName}`;
      const sheetConf = (addlCostConfig.sheetConfigs || {})[key];
      if (!sheetConf) {
        Logger.log(`[ADDL COST] No sheet config for ${key}. Skipping writeback.`);
        continue;
      }
      let sheet = null;
      for (let s = 0; s < ss.getSheets().length; s++) {
        if (ss.getSheets()[s].getName().toLowerCase().includes(sheetConf.sheetName.toLowerCase())) {
          sheet = ss.getSheets()[s]; break;
        }
      }
      if (!sheet) {
        Logger.log(`[ADDL COST] Sheet '${sheetConf.sheetName}' not found in haulier for ${rdcName}.`);
        continue;
      }
      const data = sheet.getDataRange().getValues();
      if (!data.length) continue;

      // Detect the actual header row (some sheets use a carrier band row above column headers).
      const dateKeyword = (sheetConf.dateColumn || '').toLowerCase().trim();
      let headerRowIdx = 0;
      for (let r = 0; r < Math.min(data.length, 6); r++) {
        const row = data[r] || [];
        let hasWeekHeader = false;
        for (let c = 0; c < row.length; c++) {
          const h = String(row[c] || '').trim().toLowerCase();
          if (!h) continue;
          if (h === 'week' || h.includes('week') || (dateKeyword && (h.includes(dateKeyword) || dateKeyword.includes(h)))) {
            hasWeekHeader = true;
            break;
          }
        }
        if (hasWeekHeader) {
          headerRowIdx = r;
          break;
        }
      }

      const headers = data[headerRowIdx] || [];
      const bandRow = headerRowIdx > 0 ? (data[headerRowIdx - 1] || []) : [];
      const dataStartRow = headerRowIdx + 1;

      // Determine the column section for this carrier from the optional band row.
      // Example: [ , CRE, ..., HB, ..., SCH, ... ]
      const carrierKey = String(entry.carrierName || '').trim().toUpperCase();
      let sectionStart = 0;
      let sectionEnd = headers.length - 1;
      if (bandRow.length) {
        let carrierBandIdx = -1;
        const upperBand = bandRow.map(v => String(v || '').trim().toUpperCase());
        for (let i = 0; i < upperBand.length; i++) {
          if (upperBand[i] === carrierKey) {
            carrierBandIdx = i;
            break;
          }
        }
        if (carrierBandIdx !== -1) {
          let nextBandIdx = upperBand.length;
          for (let i = carrierBandIdx + 1; i < upperBand.length; i++) {
            if (upperBand[i]) {
              nextBandIdx = i;
              break;
            }
          }
          // Include the Week column immediately before the carrier block when present.
          sectionStart = Math.max(0, carrierBandIdx - 1);
          sectionEnd = Math.max(sectionStart, nextBandIdx - 1);
        }
      }

      let dateColIdx = -1;
      const colMap = {};
      for (let j = sectionStart; j <= sectionEnd; j++) {
        const h = String(headers[j] || '').trim();
        if (!h) continue;
        const hLower = h.toLowerCase();
        // Match if either string contains the other (handles "Week" vs "Week Ending" etc.)
        if (dateColIdx === -1 && dateKeyword && (hLower.includes(dateKeyword) || dateKeyword.includes(hLower))) dateColIdx = j;
        colMap[hLower] = j;
      }
      // Always fall back to column A (index 0) — the week date is always in the first column
      if (dateColIdx === -1) {
        Logger.log(`[ADDL COST] Date column '${sheetConf.dateColumn}' not found by header; falling back to column A.`);
        dateColIdx = 0;
      }
      const weekEnd = entry.weekEndDate;
      if (!weekEnd) { Logger.log(`[ADDL COST] No week-end date for ${rdcName}/${entry.carrierName}. Skipping.`); continue; }
      const targetWeeksOrdered = getCandidateCalendarWeeksOrdered_(weekEnd);
      let targetRow = -1;
      let matchedWeek = null;
      // Prefer ISO week first, then fallback schemes only if no row exists for the preferred week.
      for (let w = 0; w < targetWeeksOrdered.length && targetRow === -1; w++) {
        const candidateWeek = targetWeeksOrdered[w];
        for (let i = dataStartRow; i < data.length; i++) {
          const v = data[i][dateColIdx];
          // Column A may contain plain week numbers (10), prefixed text (CW10), or labels (Week 10).
          const cellCw = parseCalendarWeekCell_(v);
          if (cellCw !== null && cellCw === candidateWeek) {
            targetRow = i + 1;
            matchedWeek = candidateWeek;
            break;
          }
        }
      }
      if (targetRow === -1) {
        Logger.log(`[ADDL COST] No row matching week numbers [${targetWeeksOrdered.join(', ')}] in '${sheetConf.sheetName}' for ${rdcName} (week-end ${Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'MM/dd/yyyy')}).`);
        continue;
      }
      let writeCount = 0;
      for (const [label, amount] of Object.entries(entry.buckets)) {
        const lLower = label.toLowerCase();
        let colIdx = -1;
        for (const [hLower, idx] of Object.entries(colMap)) {
          if (hLower === lLower || hLower.includes(lLower) || lLower.includes(hLower)) { colIdx = idx; break; }
        }
        if (colIdx === -1) { Logger.log(`[ADDL COST] Column for '${label}' not found.`); continue; }
        sheet.getRange(targetRow, colIdx + 1).setValue(amount > 0 ? amount : '');
        writeCount++;
      }
      Logger.log(`[ADDL COST] ${rdcName}/${entry.carrierName}: wrote ${writeCount} values to '${sheetConf.sheetName}' row ${targetRow} (matched CW${matchedWeek}, week-end ${Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'MM/dd/yyyy')}).`);
    }
  } catch(e) {
    Logger.log(`[ADDL COST ERROR] ${rdcName}: ${e.message}`);
  }
}

function getIsoWeekNumber_(date) {
  // Returns the ISO 8601 week number for a given date
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7; // make Sunday = 7
  d.setUTCDate(d.getUTCDate() + 4 - dayNum); // set to nearest Thursday
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function getUsWeekNumber_(date, mondayFirst) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const yearStart = new Date(d.getFullYear(), 0, 1);
  const dayOfYear = Math.floor((d - yearStart) / 86400000) + 1;
  const startDay = yearStart.getDay(); // 0=Sun..6=Sat
  const offset = mondayFirst ? ((startDay + 6) % 7) : startDay;
  return Math.floor((dayOfYear + offset - 1) / 7);
}

function getCandidateCalendarWeeksOrdered_(weekEndDate) {
  // Prefer ISO week first; only fallback to alternates if that row does not exist.
  const list = [];
  const seen = {};

  function pushWeek_(n) {
    if (typeof n !== 'number' || isNaN(n) || n < 1 || n > 53) return;
    if (seen[n]) return;
    seen[n] = true;
    list.push(n);
  }

  const d = new Date(weekEndDate.getFullYear(), weekEndDate.getMonth(), weekEndDate.getDate());

  pushWeek_(getIsoWeekNumber_(d));       // primary
  pushWeek_(getUsWeekNumber_(d, false)); // Sunday-based (%U-like)
  pushWeek_(getUsWeekNumber_(d, true));  // Monday-based (%W-like)

  return list;
}

function parseCalendarWeekCell_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && !isNaN(value)) {
    const n = Math.round(value);
    return (n >= 1 && n <= 53) ? n : null;
  }
  const s = String(value).trim();
  if (!s) return null;

  // Handles labels such as "CW10", "Week 10", "10 - Mar", etc.
  const matches = s.match(/\d{1,2}/g);
  if (!matches || !matches.length) return null;
  for (let i = 0; i < matches.length; i++) {
    const n = parseInt(matches[i], 10);
    if (n >= 1 && n <= 53) return n;
  }
  return null;
}

// --- HAULIER WRITE-BACK LOGIC ---

function applyUpdatesToHaulier(haulierInfo, updates) {
  try {
    const ss = SpreadsheetApp.openById(haulierInfo.spreadsheetId);
    const sheet = ss.getSheetByName(haulierInfo.sheetName);
    if (!sheet) return;

    let invCol = haulierInfo.invoiceColIdx;
    let amtCol = haulierInfo.amountColIdx;

    // Aggregate by row and keep a single TMST-derived amount per row.
    // Prefer exact TU matches over base/fuzzy matches when multiple candidates target the same row.
    const rowUpdates = {};
    const seenUpdateSignatures = {};

    function getMatchPriority_(mode) {
      const m = String(mode || '').toLowerCase();
      if (m === 'exact') return 3;
      if (m === 'base') return 2;
      if (m === 'fuzzy') return 1;
      return 0;
    }

    updates.forEach(u => {

      const carrierKey = String(u.carrierName || '').toUpperCase().trim();
      const isSchCarrier = carrierKey === 'SCH' || carrierKey.indexOf('SCH') !== -1;
      const amountNumber = (typeof u.tmstAmount === 'number' && !isNaN(u.tmstAmount))
        ? u.tmstAmount
        : parseCurrency_(u.amount);
      const hasAmount = amountNumber !== null && amountNumber !== undefined;
      const normalizedAmount = hasAmount ? Number(amountNumber).toFixed(2) : '';
      const matchPriority = getMatchPriority_(u.matchMode);
      const updateSig = [
        String(u.rowNumber || ''),
        String(u.invoiceNumber || ''),
        String(u.tuNumber || ''),
        String(u.shift || ''),
        normalizedAmount,
        String(u.matchMode || ''),
        isSchCarrier ? 'SCH' : carrierKey
      ].join('|');
      if (seenUpdateSignatures[updateSig]) return;
      seenUpdateSignatures[updateSig] = true;

      if (!rowUpdates[u.rowNumber]) {
        rowUpdates[u.rowNumber] = { invoiceNumber: u.invoiceNumber, amount: null, hasValidAmount: false, bestPriority: -1 };
      }
      if (!hasAmount) return;

      const candidateAmount = Number(amountNumber);
      rowUpdates[u.rowNumber].hasValidAmount = true;

      if (matchPriority > rowUpdates[u.rowNumber].bestPriority) {
        rowUpdates[u.rowNumber].amount = candidateAmount;
        rowUpdates[u.rowNumber].bestPriority = matchPriority;
      } else if (matchPriority === rowUpdates[u.rowNumber].bestPriority && Math.abs(Number(rowUpdates[u.rowNumber].amount) - candidateAmount) > 0.005) {
        Logger.log(
          `[WRITEBACK WARNING] Row ${u.rowNumber} received competing TMST amounts at same priority (${rowUpdates[u.rowNumber].amount} vs ${candidateAmount}, mode=${u.matchMode || 'unknown'}). Keeping first value.`
        );
      }
    });
    
    // Write the aggregated data back to the sheet
    let updateCount = 0;
    for (let row in rowUpdates) {
       const r = parseInt(row);
       sheet.getRange(r, invCol + 1).setValue(rowUpdates[row].invoiceNumber);
       // ONLY update amount if shift, store, and tour perfectly matched
         if (rowUpdates[row].hasValidAmount && rowUpdates[row].amount !== null) {
           sheet.getRange(r, amtCol + 1).setValue(rowUpdates[row].amount);
       }
       updateCount++;
    }
    
    Logger.log(`      -> Successfully wrote ${updateCount} aggregated records back to Haulier sheet: ${haulierInfo.sheetName}`);
  } catch (e) {
    Logger.log(`[ERROR] Failed to write back to Haulier sheet: ${e.message}`);
  }
}

// --- FOLDER & FILE SEARCHING LOGIC ---

function searchAndSortInvoices(rootFolder, carrierName, haulierData, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, glConfig, rdcAliases, emailTemplate, headerAliases, carrierConfig, systemConfig, haulierUpdates, addlCostConfig, addlCostWritebacks) {
  let count = 0;
  const files = rootFolder.getFiles();
  
  while (files.hasNext()) {
    count++;
    let originalFile = files.next(); 
    let file = originalFile; // Retain original item (e.g. the shortcut) for moving
    let fileName = file.getName();
    let mimeType = file.getMimeType();
    
    // If it's a shortcut, resolve it to read data, but keep 'originalFile' pointing to the shortcut!
    if (mimeType === 'application/vnd.google-apps.shortcut') {
      try {
        file = DriveApp.getFileById(file.getTargetId());
        fileName = file.getName();
        mimeType = file.getMimeType();
      } catch(e) { continue; }
    }
    
    const isExcelOrSheet = fileName.toLowerCase().indexOf('.xlsx') > -1 || mimeType === MimeType.GOOGLE_SHEETS || mimeType === MimeType.MICROSOFT_EXCEL || mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    if (isExcelOrSheet) {
      let finalRdc = parseRDCName(fileName, rdcAliases);
      
      if (finalRdc === 'UNKNOWN') {
        processInvoiceRouter(file, finalRdc, carrierName, {}, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, glConfig, rootFolder, rdcAliases, emailTemplate, headerAliases, carrierConfig, systemConfig, haulierUpdates, addlCostConfig, addlCostWritebacks);
      } else {
         let targetFolder = getOrCreateFolder(rootFolder, finalRdc);
        processInvoiceRouter(file, finalRdc, carrierName, haulierData[finalRdc] || {}, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, glConfig, targetFolder, rdcAliases, emailTemplate, headerAliases, carrierConfig, systemConfig, haulierUpdates, addlCostConfig, addlCostWritebacks);
         
         // Move the ORIGINAL FILE (the shortcut we own), not the Target File (which we don't own).
         try { moveFileToFolder(originalFile, targetFolder); } catch(e) { Logger.log("Failed to move file."); }
      }
    }
  }
  return count;
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function moveFileToFolder(file, newParent) {
  try {
    file.moveTo(newParent);
  } catch (e) {
    Logger.log(`[WARNING] Native moveTo failed: ${e.message}. Attempting Advanced Drive API fallback...`);
    
    try {
      const fileId = file.getId();
      let previousParents = [];
      const parents = file.getParents();
      while (parents.hasNext()) {
        previousParents.push(parents.next().getId());
      }
      
      Drive.Files.update({}, fileId, null, {
        addParents: newParent.getId(),
        removeParents: previousParents.join(','),
        supportsAllDrives: true
      });
    } catch (err) {
      Logger.log(`[ERROR] Advanced Drive API Move failed: ${err.message}`);
    }
  }
}

// --- INVOICE PROCESSING ROUTER ---

function processInvoiceRouter(file, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, glConfig, targetFolder, rdcAliases, emailTemplate, headerAliases, carrierConfig, systemConfig, haulierUpdates, addlCostConfig, addlCostWritebacks) {
  let sheetId = file.getId();
  let isTemp = false;
  let fileCostSummary = []; 
  const initialDiscrepancyCount = discrepancyData.length; 
  const initialTmstCount = tmstData.length;
  
  // Extract clean invoice name to act as the Invoice Number for writing back
  const fallbackInvoiceName = file.getName().replace(/\.xlsx$/i, '');
  
  if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    try {
      sheetId = convertExcelToGoogleSheet(file.getId(), targetFolder.getId());
      isTemp = true;
    } catch(e) { return; }
  }
  
  try {
    const invoiceSS = SpreadsheetApp.openById(sheetId);
    const invoiceCfg = getInvoiceExtractionConfig_(systemConfig, carrierConfig, carrierName);
    const invoiceNumber = extractInvoiceNumber(invoiceSS, fallbackInvoiceName, invoiceCfg);
    Logger.log(`      -> Extracted Invoice Number: '${invoiceNumber}'`);
    
    let isSchByStructure = false;
    if (!carrierName.toUpperCase().includes('SCH')) {
       for(let s=0; s<invoiceSS.getSheets().length; s++) {
           let sheetNameLower = invoiceSS.getSheets()[s].getName().toLowerCase();
           if(sheetNameLower.includes('order detail') || sheetNameLower.includes('customer detail') || sheetNameLower.includes('schneider')) {
               isSchByStructure = true; carrierName = 'SCH'; break;
           }
       }
    }

     let processorType = String((carrierConfig[carrierName.toUpperCase()] || {}).processorType || '').toUpperCase();
     if (!processorType) {
      if (carrierName.toUpperCase().includes('CRE')) processorType = 'CRE';
      else if (carrierName.toUpperCase().includes('SCH') || isSchByStructure) processorType = 'SCH';
      else processorType = 'HB';
     }

     if (processorType === 'CRE') {
       processCreInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates);
     } else if (processorType === 'SCH') {
       processSchInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates);
    } else {
       // Default parser for HB and HB-style dynamic carriers.
       processHbInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates);
    }

    // For CRE invoices, parse the invoice week-ending date now (fileCostSummary is populated)
    const invoiceWeekEndDate = (processorType === 'CRE') ? parseCreInvoiceWeekEnd_(invoiceSS) : null;

    const fileDiscrepancies = discrepancyData.slice(initialDiscrepancyCount);
    const fileTmstData = tmstData.slice(initialTmstCount);

    Logger.log(`    -> Discrepancy Check for ${file.getName()}: Found ${fileDiscrepancies.length} issues.`);

    // Amount Mismatch is informational only (used to verify estimation formulas).
    // It must not block GL coding or PDF generation — only structural issues (TU not found,
    // shift/store/tour mismatch) should prevent the invoice from being completed.
    const nonBlockingIssues = {
      'Amount Mismatch': true,
      'Ignored Shift 2 vs Shift 1 (Config Enabled)': true,
      'HB Blank Shift / Tour Mapping Mismatch': true
    };
    const blockingDiscrepancies = fileDiscrepancies.filter(d => !nonBlockingIssues[String(d[5] || '')]);
    const hasOnlyNonBlockingWarnings = fileDiscrepancies.length > 0 && blockingDiscrepancies.length === 0;

    if (blockingDiscrepancies.length > 0) {
      Logger.log(`    -> DRAFTING DISCREPANCY EMAIL (Skipping GL PDF) for ${file.getName()}`);
      fileDiscrepancies.forEach((d, i) => {
         Logger.log(`       [Error ${i+1}] Date: ${d[1]}, TU: '${d[2]}', Issue: ${d[5]}`);
      });
      createDiscrepancyEmailDraft(carrierName, file.getName(), fileDiscrepancies, emailTemplate);
    } 
    else if (rdcName !== 'UNKNOWN' && fileCostSummary.length > 0) {
      if (_extractOnlyMode) {
        const suggestedGl = buildSuggestedGLCodes_(invoiceSS, file.getName(), rdcName, fileCostSummary, glConfig, fileTmstData, false);
        const suggestedCount = Object.keys(suggestedGl.glTotals || {}).length;
        Logger.log(`    -> [EXTRACT MODE] Loaded ${file.getName()} for review/finalization. Cost items: ${fileCostSummary.length}.`);
        if (suggestedCount > 0) {
          Logger.log(`    -> [EXTRACT MODE] Preloaded ${suggestedCount} suggested GL code bucket(s) for manual review.`);
        }
        savePendingInvoiceGLCodes_(
          file.getName(),
          rdcName,
          invoiceNumber || fallbackInvoiceName,
          suggestedGl.grandTotal || sumFileCostSummary_(fileCostSummary),
          suggestedGl.glTotals || {},
          {
            sheetId: sheetId,
            targetFolderId: targetFolder.getId(),
            costSummary: fileCostSummary,
            tmstData: fileTmstData
          }
        );
        _pendingInvoices.push({
          fileName: file.getName(),
          sheetId: sheetId,
          rdcName: rdcName,
          carrierName: carrierName,
          costSummary: fileCostSummary,
          tmstData: fileTmstData,
          isTemp: isTemp
        });
        if (hasOnlyNonBlockingWarnings) {
          createDiscrepancyEmailDraft(carrierName, file.getName(), fileDiscrepancies, emailTemplate);
        }
      } else {
        Logger.log(`    -> Applying GL Codes & generating PDF for ${file.getName()}. Total cost items found: ${fileCostSummary.length}${hasOnlyNonBlockingWarnings ? ' (Non-blocking warnings present — coding proceeding)' : ''}`);
        applyGLCodesAndExportPDF(invoiceSS, file.getName(), rdcName, fileCostSummary, glConfig, targetFolder, fileTmstData);
        // If there were only non-blocking warnings, still draft a warning email so they can be reviewed.
        if (hasOnlyNonBlockingWarnings) {
          createDiscrepancyEmailDraft(carrierName, file.getName(), fileDiscrepancies, emailTemplate);
        }
      }

      // Queue additional cost writeback for CRE invoices after successful processing
      if (processorType === 'CRE' && invoiceWeekEndDate && addlCostConfig && addlCostWritebacks && rdcName !== 'UNKNOWN') {
        const buckets = collectAddlCostBuckets_(fileCostSummary, carrierName, rdcName, addlCostConfig);
        if (Object.keys(buckets).length > 0) {
          if (!addlCostWritebacks[rdcName]) addlCostWritebacks[rdcName] = [];
          addlCostWritebacks[rdcName].push({ carrierName: carrierName, weekEndDate: invoiceWeekEndDate, buckets: buckets });
          Logger.log(`    -> [ADDL COST] Queued ${Object.keys(buckets).length} bucket(s) for ${rdcName}/CRE, week-end ${Utilities.formatDate(invoiceWeekEndDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')}`);
        }
      }
    } 
    else {
      Logger.log(`    -> SKIPPING GL PDF for ${file.getName()}: RDC is '${rdcName}' (Must not be UNKNOWN) AND Cost Summary length is ${fileCostSummary.length} (Must be > 0).`);
    }

  } catch (e) {
    Logger.log(`Error processing sheet: ${e.message}`);
    
    // Move unprocessable files to 'Unprocessed Temporary' folder
    try {
      const unprocessedFolder = getOrCreateFolder(targetFolder, 'Unprocessed Temporary');
      moveFileToFolder(file, unprocessedFolder);
      Logger.log(`    -> File moved to 'Unprocessed Temporary' folder due to processing error.`);
    } catch (moveErr) {
      Logger.log(`    -> Could not move file to Unprocessed folder: ${moveErr.message}`);
    }
  }
  
  if (isTemp && !_extractOnlyMode) {
    try { DriveApp.getFileById(sheetId).setTrashed(true); } catch(e) { }
  }
}

// --- EMAIL DRAFTING LOGIC ---

function createDiscrepancyEmailDraft(carrierName, fileName, discrepancies, emailTemplate) {
  const subjTemplate = (emailTemplate && emailTemplate['Subject']) ? emailTemplate['Subject'] : `Discrepancy Notice: Invoice {FileName}`;
  const greetTemplate = (emailTemplate && emailTemplate['Greeting']) ? emailTemplate['Greeting'] : `Hello {CarrierName} Team,\n\nThe following items are showing as discrepancies in our system. Please advise:`;
  const outroTemplate = (emailTemplate && emailTemplate['Outro']) ? emailTemplate['Outro'] : `Thank you.`;

  const subject = subjTemplate.replace(/{FileName}/g, fileName).replace(/{CarrierName}/g, carrierName);
  const greetingText = greetTemplate.replace(/{CarrierName}/g, carrierName).replace(/\n/g, '<br>');
  const outroText = outroTemplate.replace(/\n/g, '<br>');
  
  let plainBody = `${greetingText.replace(/<br>/g, '\n')}\n\n`;
  
  let htmlBody = `<div style="font-family: Arial, sans-serif; color: #333;">`;
  htmlBody += `<p>${greetingText}</p>`;
  htmlBody += `<table style="border-collapse: collapse; width: 100%; max-width: 800px; margin-top: 15px;">`;
  htmlBody += `<thead>`;
  htmlBody += `<tr style="background-color: #f2f2f2;">`;
  htmlBody += `<th style="border: 1px solid #000; padding: 10px; text-align: left;">Date</th>`;
  htmlBody += `<th style="border: 1px solid #000; padding: 10px; text-align: left;">TU</th>`;
  htmlBody += `<th style="border: 1px solid #000; padding: 10px; text-align: left;">Issue</th>`;
  htmlBody += `<th style="border: 1px solid #000; padding: 10px; text-align: left;">Billed Store/Tour</th>`;
  htmlBody += `</tr>`;
  htmlBody += `</thead>`;
  htmlBody += `<tbody>`;

  discrepancies.forEach(d => {
    let dateStr = d[1];
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    }
    
    plainBody += `• Date: ${dateStr} | TU: ${d[2]} | Issue: ${d[5]} (Billed: ${d[4]})\n`;
    
    htmlBody += `<tr>`;
    htmlBody += `<td style="border: 1px solid #000; padding: 8px;">${dateStr}</td>`;
    htmlBody += `<td style="border: 1px solid #000; padding: 8px;">${d[2]}</td>`;
    htmlBody += `<td style="border: 1px solid #000; padding: 8px;">${d[5]}</td>`;
    htmlBody += `<td style="border: 1px solid #000; padding: 8px;">${d[4]}</td>`;
    htmlBody += `</tr>`;
  });

  plainBody += `\n${outroText.replace(/<br>/g, '\n')}`;
  
  htmlBody += `</tbody>`;
  htmlBody += `</table>`;
  htmlBody += `<p style="margin-top: 20px;">${outroText}</p>`;
  htmlBody += `</div>`;

  GmailApp.createDraft("", subject, plainBody, { htmlBody: htmlBody });
}

// --- GL CODING & PDF EXPORT LOGIC ---

function buildSuggestedGLCodes_(invoiceSS, originalFileName, rdcName, fileCostSummary, glConfig, fileTmstData, enableLogs) {
  const result = { glTotals: {}, grandTotal: 0 };
  if (!glConfig) return result;

  const costCenter = glConfig.costCenters[rdcName] || 'XXXXX';
  const defaultPrefix = glConfig.defaultCategory || 'XXXXXX, XXXXXXXX';
  const defaultGlString = `${defaultPrefix}, ${costCenter}`;
  const othersPrefix = getCategoryPrefix_(glConfig, 'others', '471000, 47100099');
  const othersGlString = `${othersPrefix}, ${costCenter}`;
  const tollPrefix = getCategoryPrefix_(glConfig, 'toll', '471000, 47100002');
  const tollGlString = `${tollPrefix}, ${costCenter}`;
  const transferPrefix = getCategoryPrefix_(glConfig, 'transfer', '471000, 47100004');
  const import471003Prefix = getCategoryPrefix_(glConfig, 'import_471003', '471003');
  const selfPickup360100Prefix = getCategoryPrefix_(glConfig, 'self_pickup_360100', '360100');

  const glTotals = {};
  let grandTotal = 0;

  const invoiceTotal = findInvoiceTotal_(invoiceSS);
  const tmstAllocation = buildTmstAllocation_(fileTmstData, glConfig);
  const summaryTotal = sumFileCostSummary_(fileCostSummary);
  const tmstTollTotal = sumTmstColumn_(fileTmstData, 7);
  const invoiceCarrier = (fileTmstData && fileTmstData.length > 0)
    ? String(fileTmstData[0][14] || '').trim().toUpperCase()
    : '';
  const isCreInvoice = invoiceCarrier === 'CRE';

  if (tmstAllocation.total > 0) {
    const othersTotal = tmstAllocation.othersTotal;
    const transferTotal = sumObjectValues_(tmstAllocation.transferByCostCenter);
    const authoritativeTotal = (invoiceTotal !== null && invoiceTotal > 0)
      ? invoiceTotal
      : Math.max(tmstAllocation.total, summaryTotal);

    let baseTotal = authoritativeTotal - othersTotal - transferTotal;
    if (baseTotal < 0) {
      if (enableLogs) {
        Logger.log(
          `[WARNING] Base allocation negative for ${originalFileName}. Invoice: $${authoritativeTotal.toFixed(2)} | Others: $${othersTotal.toFixed(2)} | Transfer: $${transferTotal.toFixed(2)}.`
        );
      }
      baseTotal = 0;
    }

    const tollToAllocate = Math.max(0, Math.min(baseTotal, tmstTollTotal));
    if (tollToAllocate > 0) {
      baseTotal -= tollToAllocate;
      glTotals[tollGlString] = (glTotals[tollGlString] || 0) + tollToAllocate;
    }

    if (isCreInvoice) {
      const savings = getCreImportSavingsTotals_(rdcName, fileTmstData);
      const importAlloc = Math.max(0, Math.min(baseTotal, Number(savings.import471003 || 0)));
      baseTotal -= importAlloc;

      const selfPickAlloc = Math.max(0, Math.min(baseTotal, Number(savings.selfPickup360100 || 0)));
      baseTotal -= selfPickAlloc;

      if (importAlloc > 0) {
        const importGlString = `${import471003Prefix}, ${costCenter}`;
        glTotals[importGlString] = (glTotals[importGlString] || 0) + importAlloc;
      }
      if (selfPickAlloc > 0) {
        const selfPickupGlString = `${selfPickup360100Prefix}, ${costCenter}`;
        glTotals[selfPickupGlString] = (glTotals[selfPickupGlString] || 0) + selfPickAlloc;
      }

      if (enableLogs) {
        Logger.log(
          `[CRE SAVINGS] ${originalFileName} | RDC=${rdcName} | Import471003=$${importAlloc.toFixed(2)} | SelfPickup360100=$${selfPickAlloc.toFixed(2)} | Source=${savings.sourceNote}`
        );
      }
    }

    if (baseTotal > 0) glTotals[defaultGlString] = baseTotal;
    if (othersTotal > 0) glTotals[othersGlString] = othersTotal;

    Object.keys(tmstAllocation.transferByCostCenter).forEach(cc => {
      const amount = tmstAllocation.transferByCostCenter[cc];
      if (amount <= 0) return;
      const transferGlString = `${transferPrefix}, ${cc}`;
      glTotals[transferGlString] = (glTotals[transferGlString] || 0) + amount;
    });

    grandTotal = sumObjectValues_(glTotals);
    if (authoritativeTotal > 0 && Math.abs(grandTotal - authoritativeTotal) > 0.01) {
      const delta = authoritativeTotal - grandTotal;
      glTotals[defaultGlString] = Math.max(0, (glTotals[defaultGlString] || 0) + delta);
      grandTotal = sumObjectValues_(glTotals);
      if (enableLogs) {
        Logger.log(
          `[WARNING] Allocation normalized for ${originalFileName}. Delta applied to base bucket: $${delta.toFixed(2)}.`
        );
      }
    }

    if (enableLogs) {
      Logger.log(
        `[GL ALLOCATION] ${originalFileName} | Invoice: $${authoritativeTotal.toFixed(2)} | Base(Store): $${(glTotals[defaultGlString] || 0).toFixed(2)} | Tolls: $${(glTotals[tollGlString] || 0).toFixed(2)} | Others(O): $${(glTotals[othersGlString] || 0).toFixed(2)} | Transfer: $${transferTotal.toFixed(2)} | Stamped Total: $${grandTotal.toFixed(2)}`
      );
    }
  } else {
    if (fileCostSummary.length === 0) return result;

    fileCostSummary.forEach(item => {
      const descLower = String(item.desc || '').toLowerCase();
      let glPrefix = defaultPrefix;

      for (let i = 0; i < (glConfig.categories || []).length; i++) {
        const cat = glConfig.categories[i];
        if (descLower.includes(cat.keyword)) {
          glPrefix = cat.glPrefix;
          break;
        }
      }

      const amount = parseCurrency_(item.amount) || 0;
      const fullGlString = `${glPrefix}, ${costCenter}`;
      if (!glTotals[fullGlString]) glTotals[fullGlString] = 0;
      glTotals[fullGlString] += amount;
      grandTotal += amount;
    });

    if (invoiceTotal !== null && invoiceTotal > 0 && grandTotal > invoiceTotal + 0.01) {
      let overflow = grandTotal - invoiceTotal;

      if ((glTotals[defaultGlString] || 0) > 0) {
        const cut = Math.min(glTotals[defaultGlString], overflow);
        glTotals[defaultGlString] -= cut;
        overflow -= cut;
      }

      if (overflow > 0) {
        const keysBySize = Object.keys(glTotals)
          .filter(k => glTotals[k] > 0)
          .sort((a, b) => glTotals[b] - glTotals[a]);

        for (let i = 0; i < keysBySize.length && overflow > 0; i++) {
          const key = keysBySize[i];
          const cut = Math.min(glTotals[key], overflow);
          glTotals[key] -= cut;
          overflow -= cut;
        }
      }

      grandTotal = invoiceTotal;
    }

    if (enableLogs) {
      Logger.log(`[GL ALLOCATION] ${originalFileName} | Fallback summary allocation used | Stamped Total: $${grandTotal.toFixed(2)}`);
    }
  }

  result.glTotals = glTotals;
  result.grandTotal = grandTotal;
  return result;
}

function applyGLCodesAndExportPDF(invoiceSS, originalFileName, rdcName, fileCostSummary, glConfig, targetFolder, fileTmstData) {
  const suggestion = buildSuggestedGLCodes_(invoiceSS, originalFileName, rdcName, fileCostSummary, glConfig, fileTmstData, true);
  const glTotals = suggestion.glTotals || {};
  const grandTotal = Number(suggestion.grandTotal || 0);
  if (!Object.keys(glTotals).length) return;
  
  let stampRows = [];
  stampRows.push(['GL CODING SUMMARY', 'GL ACCOUNT / COST CENTER']);
  for (let gl in glTotals) {
     if (glTotals[gl] > 0) {
        stampRows.push([`$${glTotals[gl].toFixed(2)}`, gl]);
     }
  }
  stampRows.push([`$${grandTotal.toFixed(2)}`, `TOTAL INVOICE AMOUNT`]);
  
  let summarySheet = null;
  for (let sheet of invoiceSS.getSheets()) {
    const name = sheet.getName().toLowerCase();
    if (name.includes('summary') || name.includes('invoice')) { summarySheet = sheet; break; }
  }
  if (!summarySheet) summarySheet = invoiceSS.getSheets()[0];
  
  summarySheet.insertRowsBefore(1, stampRows.length + 1);
  const range = summarySheet.getRange(1, 1, stampRows.length, 2);
  range.setValues(stampRows);
  
  // Styling the stamp
  range.setFontWeight('bold').setFontColor('#cc0000').setWrap(false);
  range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  summarySheet.getRange(1, 1, 1, 2).setBackground('#ffff00').setFontColor('#000000');
  
  // Auto-resize and enforce minimum widths so SCH/narrow tabs don't squish the PDF
  summarySheet.autoResizeColumn(1);
  summarySheet.autoResizeColumn(2);
  if (summarySheet.getColumnWidth(1) < 160) summarySheet.setColumnWidth(1, 160);
  if (summarySheet.getColumnWidth(2) < 280) summarySheet.setColumnWidth(2, 280);
  
  const lastCol = summarySheet.getLastColumn();
  const maxCols = summarySheet.getMaxColumns();
  if (maxCols > lastCol) {
    summarySheet.hideColumns(lastCol + 1, maxCols - lastCol);
  }
  
  SpreadsheetApp.flush(); 
  
  const pdfName = originalFileName.replace(/\.xlsx$/i, '') + ' - CODED.pdf';
  exportSheetToPDF(invoiceSS.getId(), summarySheet.getSheetId(), pdfName, targetFolder);
}

function applyManualGLCodesAndExportPDF(invoiceSS, originalFileName, glCodeMap, targetFolder) {
  const glTotals = glCodeMap || {};
  const grandTotal = Object.keys(glTotals).reduce(function(sum, key) {
    return sum + (parseFloat(glTotals[key]) || 0);
  }, 0);

  let stampRows = [];
  stampRows.push(['GL CODING SUMMARY', 'GL ACCOUNT / COST CENTER']);
  Object.keys(glTotals).forEach(function(gl) {
    const amount = parseFloat(glTotals[gl]) || 0;
    if (amount > 0) stampRows.push([`$${amount.toFixed(2)}`, gl]);
  });
  stampRows.push([`$${grandTotal.toFixed(2)}`, 'TOTAL INVOICE AMOUNT']);

  let summarySheet = null;
  const sheets = invoiceSS.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName().toLowerCase();
    if (name.includes('summary') || name.includes('invoice')) { summarySheet = sheets[i]; break; }
  }
  if (!summarySheet) summarySheet = sheets[0];

  summarySheet.insertRowsBefore(1, stampRows.length + 1);
  const range = summarySheet.getRange(1, 1, stampRows.length, 2);
  range.setValues(stampRows);
  range.setFontWeight('bold').setFontColor('#cc0000').setWrap(false);
  range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  summarySheet.getRange(1, 1, 1, 2).setBackground('#ffff00').setFontColor('#000000');

  summarySheet.autoResizeColumn(1);
  summarySheet.autoResizeColumn(2);
  if (summarySheet.getColumnWidth(1) < 160) summarySheet.setColumnWidth(1, 160);
  if (summarySheet.getColumnWidth(2) < 280) summarySheet.setColumnWidth(2, 280);

  SpreadsheetApp.flush();

  const pdfName = originalFileName.replace(/\.xlsx$/i, '') + ' - CODED.pdf';
  exportSheetToPDF(invoiceSS.getId(), summarySheet.getSheetId(), pdfName, targetFolder);
}

// ═══════════════════════════════════════════════════════════════════════
// INVOICE FINALIZATION FUNCTIONS (Edit GL codes before saving)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Save current GL codes for an invoice into the finalization sheet
 * for manual editing before final PDF export
 */
function savePendingInvoiceGLCodes_(invoiceFileName, rdcName, invoiceNumber, totalAmount, glCodeMap, meta) {
  ensureInvoiceFinalizationSchema_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finSheet = ss.getSheetByName('Invoice Finalization');
  if (!finSheet) return { success: false, message: 'Invoice Finalization sheet not found' };

  try {
    meta = meta || {};
    const glCodeJson = JSON.stringify(glCodeMap);
    const costSummaryJson = JSON.stringify(meta.costSummary || []);
    const tmstJson = JSON.stringify(meta.tmstData || []);
    const timestamp = new Date().toISOString();
    
    // Add a new row to track this pending finalization
    finSheet.appendRow([
      invoiceFileName,
      rdcName,
      invoiceNumber,
      totalAmount,
      timestamp,
      glCodeJson,
      'PENDING',
      String(meta.sheetId || ''),
      String(meta.targetFolderId || ''),
      costSummaryJson,
      tmstJson
    ]);
    
    Logger.log(`[FINALIZATION] Saved GL codes for ${invoiceFileName} - awaiting manual review`);
    return { success: true, invoiceFileName: invoiceFileName };
  } catch (e) {
    Logger.log(`[ERROR] Failed to save pending GL codes: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Get current pending invoice GL codes for editing in the UI
 */
function getPendingInvoiceGLCodes(invoiceFileName) {
  ensureInvoiceFinalizationSchema_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finSheet = ss.getSheetByName('Invoice Finalization');
  if (!finSheet) return { success: false, message: 'Invoice Finalization sheet not found', codes: [] };

  try {
    const data = finSheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return { success: false, message: 'No pending invoices found', codes: [] };
    }

    const headerRow = data[0];
    const fileNameIdx = headerRow.indexOf('Invoice File Name');
    const glCodeIdx = headerRow.indexOf('GL Code Config (JSON)');
    const rdcIdx = headerRow.indexOf('RDC');
    const amountIdx = headerRow.indexOf('Total Amount');
    const invoiceNumIdx = headerRow.indexOf('Invoice Number');
    const statusIdx = headerRow.indexOf('Status');

    if (fileNameIdx === -1 || glCodeIdx === -1) {
      return { success: false, message: 'Column headers not found', codes: [] };
    }

    let targetRow = null;
    for (let i = data.length - 1; i > 0; i--) {
      const status = statusIdx >= 0 ? String(data[i][statusIdx] || '').toUpperCase() : '';
      if (status !== 'PENDING' && status !== 'EDITED') continue;
      if (!invoiceFileName || String(data[i][fileNameIdx]).indexOf(invoiceFileName) !== -1) {
        targetRow = data[i];
        break;
      }
    }

    if (!targetRow) {
      return { success: false, message: 'No matching invoice found', codes: [] };
    }

    const rdcName = rdcIdx >= 0 ? String(targetRow[rdcIdx] || '') : '';
    const allocationMeta = getFinalizationGlMeta_(rdcName);

    const glCodeJson = String(targetRow[glCodeIdx] || '{}');
    const glCodeMap = JSON.parse(glCodeJson || '{}');
    const codesArray = Object.entries(glCodeMap).map(function(entry) {
      const normalizedCode = normalizeFinalizationGLCode_(String(entry[0] || ''), allocationMeta);
      return {
        glCode: normalizedCode,
        amount: Number(entry[1] || 0),
        editable: true
      };
    });

    return {
      success: true,
      invoiceFileName: String(targetRow[fileNameIdx] || ''),
      rdc: rdcName,
      invoiceNumber: invoiceNumIdx >= 0 ? String(targetRow[invoiceNumIdx] || '') : '',
      totalAmount: amountIdx >= 0 ? Number(targetRow[amountIdx] || 0) : 0,
      codes: codesArray,
      allocationMeta: allocationMeta
    };
  } catch (e) {
    Logger.log(`[ERROR] Failed to retrieve pending GL codes: ${e.message}`);
    return { success: false, message: e.message, codes: [] };
  }
}

/**
 * Update GL codes for a pending invoice
 */
function updatePendingInvoiceGLCodes(invoiceFileName, updatedGlCodeArray) {
  ensureInvoiceFinalizationSchema_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finSheet = ss.getSheetByName('Invoice Finalization');
  if (!finSheet) return { success: false, message: 'Invoice Finalization sheet not found' };

  try {
    const data = finSheet.getDataRange().getValues();
    const headerRow = data[0];
    const fileNameIdx = headerRow.indexOf('Invoice File Name');
    const glCodeIdx = headerRow.indexOf('GL Code Config (JSON)');
    const statusIdx = headerRow.indexOf('Status');

    if (fileNameIdx === -1 || glCodeIdx === -1) {
      return { success: false, message: 'Column headers not found' };
    }

    // Find the row with this invoice
    let updateRowNum = null;
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][fileNameIdx]).includes(invoiceFileName)) {
        updateRowNum = i + 1; // Convert to 1-based row number
        break;
      }
    }

    if (!updateRowNum) {
      return { success: false, message: 'Invoice not found in finalization sheet' };
    }

    const rdcIdx = headerRow.indexOf('RDC');
    const currentData = finSheet.getRange(updateRowNum, 1, 1, headerRow.length).getValues()[0];
    const allocationMeta = getFinalizationGlMeta_(rdcIdx >= 0 ? String(currentData[rdcIdx] || '') : '');

    // Convert array back to object
    const glCodeMap = {};
    let totalAmount = 0;
    updatedGlCodeArray.forEach(item => {
      const amount = Number(item.amount || 0);
      if (amount <= 0) return;

      const normalizedCode = normalizeFinalizationGLCode_(String(item.glCode || ''), allocationMeta);
      if (!normalizedCode) return;

      glCodeMap[normalizedCode] = (glCodeMap[normalizedCode] || 0) + amount;
      totalAmount += amount;
    });

    // Validate total hasn't changed significantly
    const originalAmount = Number(currentData[headerRow.indexOf('Total Amount')] || 0);
    
    if (Math.abs(totalAmount - originalAmount) > 0.01) {
      Logger.log(`[WARNING] GL code total changed from $${originalAmount.toFixed(2)} to $${totalAmount.toFixed(2)}`);
    }

    // Update the row
    const glCodeJson = JSON.stringify(glCodeMap);
    finSheet.getRange(updateRowNum, glCodeIdx + 1).setValue(glCodeJson);
    finSheet.getRange(updateRowNum, statusIdx + 1).setValue('EDITED');
    
    Logger.log(`[FINALIZATION] Updated GL codes for ${invoiceFileName}`);
    return { success: true, message: 'GL codes updated successfully' };
  } catch (e) {
    Logger.log(`[ERROR] Failed to update GL codes: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function getFinalizationGlMeta_(rdcName) {
  const glConfig = getGlConfig() || {};
  const rdcKey = String(rdcName || '').trim().toUpperCase();

  const costCenter = ((glConfig.costCenters || {})[rdcKey] || '').trim();
  const basePrefix = String(glConfig.defaultCategory || '471000, 47100001').trim();
  const importPrefix = getCategoryPrefix_(glConfig, 'import_471003', '471003');
  const selfPickupPrefix = getCategoryPrefix_(glConfig, 'self_pickup_360100', '360100');

  return {
    rdc: rdcKey,
    costCenter: costCenter,
    basePrefix: basePrefix,
    importPrefix: String(importPrefix || '471003').trim(),
    selfPickupPrefix: String(selfPickupPrefix || '360100').trim()
  };
}

function normalizeFinalizationGLCode_(glCode, allocationMeta) {
  const raw = String(glCode || '').trim().replace(/\s+/g, ' ');
  if (!raw) return '';

  const meta = allocationMeta || {};
  const costCenter = String(meta.costCenter || '').trim();
  const importPrefix = String(meta.importPrefix || '471003').trim();
  const selfPickupPrefix = String(meta.selfPickupPrefix || '360100').trim();

  function parsePrefix_(text) {
    const code = String(text || '').split(',')[0].trim();
    return code;
  }

  const firstCode = parsePrefix_(raw);
  const importCode = parsePrefix_(importPrefix);
  const selfCode = parsePrefix_(selfPickupPrefix);

  const needsCostCenter = (firstCode === importCode || firstCode === selfCode);
  if (!needsCostCenter || !costCenter) return raw;

  const parts = raw.split(',').map(function(p) { return String(p || '').trim(); }).filter(Boolean);
  if (parts.length >= 2 && parts[parts.length - 1] === costCenter) return raw;

  return `${firstCode}, ${costCenter}`;
}

/**
 * Get all GL code templates for the dropdown in finalization UI
 */
function getGLCodeTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templates = ss.getSheetByName('GL Code Templates');
  if (!templates || templates.getLastRow() < 2) {
    return { success: false, templates: [] };
  }

  try {
    const data = templates.getDataRange().getValues();
    const headerRow = data[0];
    const nameIdx = headerRow.indexOf('Template Name');
    const codeIdx = headerRow.indexOf('GL Code / Cost Center String');
    const descIdx = headerRow.indexOf('Description');
    const quickSelectIdx = headerRow.indexOf('Quick Select (YES/NO)');

    const templateArray = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[nameIdx]) continue; // Skip empty rows
      
      templateArray.push({
        name: String(row[nameIdx] || ''),
        code: String(row[codeIdx] || ''),
        description: String(row[descIdx] || ''),
        quickSelect: String(row[quickSelectIdx] || 'NO').toUpperCase() === 'YES'
      });
    }

    return { success: true, templates: templateArray };
  } catch (e) {
    Logger.log(`[ERROR] Failed to retrieve GL code templates: ${e.message}`);
    return { success: false, templates: [] };
  }
}

/**
 * Add a new GL code to a pending invoice
 */
function addGLCodeToInvoice(invoiceFileName, newGlCode, amount) {
  ensureInvoiceFinalizationSchema_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finSheet = ss.getSheetByName('Invoice Finalization');
  if (!finSheet) return { success: false, message: 'Invoice Finalization sheet not found' };

  try {
    const current = getPendingInvoiceGLCodes(invoiceFileName);
    if (!current.success) {
      return current;
    }

    // Add the new code to the array
    current.codes.push({
      glCode: newGlCode,
      amount: Number(amount || 0),
      editable: true
    });

    // Update the sheet
    return updatePendingInvoiceGLCodes(invoiceFileName, current.codes);
  } catch (e) {
    Logger.log(`[ERROR] Failed to add GL code: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function ensureInvoiceFinalizationSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  const requiredHeaders = [
    'Invoice File Name',
    'RDC',
    'Invoice Number',
    'Total Amount',
    'Saved At',
    'GL Code Config (JSON)',
    'Status',
    'Sheet ID',
    'Target Folder ID',
    'Cost Summary JSON',
    'TMST JSON'
  ];

  let finSheet = ss.getSheetByName('Invoice Finalization');
  if (!finSheet) {
    finSheet = ss.insertSheet('Invoice Finalization');
  }

  if (finSheet.getLastRow() === 0) {
    finSheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]).setFontWeight('bold').setBackground('#b6d7a8');
    return;
  }

  const headerWidth = Math.max(finSheet.getLastColumn(), requiredHeaders.length);
  const headerRow = finSheet.getRange(1, 1, 1, headerWidth).getValues()[0];

  requiredHeaders.forEach(function(colName) {
    if (headerRow.indexOf(colName) === -1) {
      const newCol = finSheet.getLastColumn() + 1;
      finSheet.getRange(1, newCol).setValue(colName).setFontWeight('bold').setBackground('#b6d7a8');
      headerRow.push(colName);
    }
  });
}

function sumFileCostSummary_(fileCostSummary) {
  if (!fileCostSummary || !fileCostSummary.length) return 0;
  return fileCostSummary.reduce((sum, item) => {
    const amount = parseCurrency_(item && item.amount);
    return sum + (amount && amount > 0 ? amount : 0);
  }, 0);
}

function sumTmstColumn_(rows, index) {
  if (!rows || !rows.length) return 0;
  return rows.reduce((sum, row) => {
    const amount = parseCurrency_(row && row[index]);
    return sum + (amount && amount > 0 ? amount : 0);
  }, 0);
}

function findInvoiceTotal_(invoiceSS) {
  const labelRegex = /(invoice total|amount due|total due)/i;
  const sheets = invoiceSS.getSheets();
  const candidates = [];

  for (let s = 0; s < sheets.length; s++) {
    const data = sheets[s].getDataRange().getDisplayValues();

    for (let r = 0; r < Math.min(data.length, 350); r++) {
      const row = data[r];
      if (isStampedCodingRow_(row)) continue;

      for (let c = 0; c < Math.min(row.length, 20); c++) {
        const cellText = String(row[c] || '').trim();
        if (!cellText || !labelRegex.test(cellText)) continue;

        const inlineAmount = parseCurrency_(cellText);
        if (inlineAmount !== null && inlineAmount > 0) candidates.push(inlineAmount);

        // Prefer amount in the same row to the right of the label.
        for (let k = c + 1; k < Math.min(row.length, c + 8); k++) {
          const parsed = parseCurrency_(row[k]);
          if (parsed !== null && parsed > 0) candidates.push(parsed);
        }

        // Fallback: sometimes amount is placed directly under the label.
        for (let rr = r + 1; rr <= Math.min(r + 2, data.length - 1); rr++) {
          const parsedNext = parseCurrency_(data[rr][c]);
          if (parsedNext !== null && parsedNext > 0) candidates.push(parsedNext);
        }
      }
    }
  }

  if (candidates.length === 0) return null;
  return Math.max.apply(null, candidates);
}

function parseCurrency_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return isNaN(value) ? null : value;

  const text = String(value).trim();
  if (!text) return null;

  // Prefer explicit currency/decimal formats and ignore date-like integers.
  const matches = text.match(/-?\$?\s*\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\$?\s*\d+\.\d+/g);
  if (matches && matches.length) {
    let best = null;
    matches.forEach(token => {
      const n = parseFloat(String(token).replace(/[^0-9.-]/g, ''));
      if (!isNaN(n) && (best === null || Math.abs(n) > Math.abs(best))) best = n;
    });
    return best;
  }

  // Fallback for plain numeric cells (no commas/decimals), e.g. 5000.
  if (!/^\s*[$-]?\s*\d+\s*$/.test(text)) return null;
  const cleaned = text.replace(/[^0-9.-]/g, '');
  if (!cleaned || cleaned === '-' || cleaned === '.' || cleaned === '-.') return null;

  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? null : parsed;
}

function getCategoryPrefix_(glConfig, keyword, fallbackPrefix) {
  const target = String(keyword || '').toLowerCase();
  for (let i = 0; i < (glConfig.categories || []).length; i++) {
    const cat = glConfig.categories[i];
    if (String(cat.keyword || '').toLowerCase() === target) {
      return cat.glPrefix || fallbackPrefix;
    }
  }
  return fallbackPrefix;
}

function getCreImportSavingsTotals_(rdcName, fileTmstData) {
  const rdcKey = String(rdcName || '').trim().toUpperCase();
  const result = {
    import471003: 0,
    selfPickup360100: 0,
    sourceNote: `SC_Import Savings not found in ${rdcKey || 'UNKNOWN'} haulier workbook`
  };

  function toDateOnly_(value) {
    if (value instanceof Date && !isNaN(value.getTime())) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }
    if (value === null || value === undefined) return null;
    const text = String(value).trim();
    if (!text) return null;

    const direct = new Date(text);
    if (!isNaN(direct.getTime())) {
      return new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
    }

    const m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (!m) return null;
    const mm = parseInt(m[1], 10) - 1;
    const dd = parseInt(m[2], 10);
    let yy = parseInt(m[3], 10);
    if (yy < 100) yy += 2000;
    const d = new Date(yy, mm, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  function startOfIsoWeek_(year, isoWeek) {
    const jan4 = new Date(year, 0, 4);
    const jan4Dow = (jan4.getDay() + 6) % 7; // Monday=0
    const week1Mon = new Date(year, 0, 4 - jan4Dow);
    return new Date(week1Mon.getFullYear(), week1Mon.getMonth(), week1Mon.getDate() + ((isoWeek - 1) * 7));
  }

  function overlapDaysInclusive_(aStart, aEnd, bStart, bEnd) {
    const start = Math.max(aStart.getTime(), bStart.getTime());
    const end = Math.min(aEnd.getTime(), bEnd.getTime());
    if (end < start) return 0;
    return Math.floor((end - start) / 86400000) + 1;
  }

  try {
    const cfg = getConfig();
    const haulierSheetId = cfg[`${rdcKey}_HAULIER_ID`];
    if (!haulierSheetId) {
      result.sourceNote = `Missing ${rdcKey}_HAULIER_ID in System Config`;
      return result;
    }

    const ss = SpreadsheetApp.openById(haulierSheetId);
    let sheet = ss.getSheetByName('SC_Import Savings');
    if (!sheet) {
      // Tolerate minor naming differences while still preferring exact tab name.
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const n = String(sheets[i].getName() || '').toLowerCase().replace(/\s+/g, ' ').trim();
        if (n === 'sc_import savings' || n === 'sc import savings' || n.indexOf('import savings') !== -1) {
          sheet = sheets[i];
          break;
        }
      }
    }
    if (!sheet) return result;

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      result.sourceNote = 'SC_Import Savings has no data rows';
      return result;
    }

    const invoiceDates = [];
    (fileTmstData || []).forEach(row => {
      const d = toDateOnly_(row && row[1]);
      if (d) invoiceDates.push(d);
    });
    if (!invoiceDates.length) {
      result.sourceNote = 'No invoice dates found for CRE proration';
      return result;
    }

    let invStart = invoiceDates[0];
    let invEnd = invoiceDates[0];
    invoiceDates.forEach(d => {
      if (d.getTime() < invStart.getTime()) invStart = d;
      if (d.getTime() > invEnd.getTime()) invEnd = d;
    });

    let headerRow = -1;
    let importIdx = -1;
    let importBookToCreIdx = -1;
    let selfPickIdx = -1;
    let selfPickThirdIdx = -1;

    for (let r = 0; r < Math.min(data.length, 12); r++) {
      const row = data[r] || [];
      for (let c = 0; c < row.length; c++) {
        const h = String(row[c] || '').toLowerCase().trim().replace(/\s+/g, ' ');
        if (!h) continue;
        if (importIdx === -1 && h.indexOf('import (471003)') !== -1) importIdx = c;
        if (importBookToCreIdx === -1 && h.indexOf('import book to cre') !== -1) importBookToCreIdx = c;
        if (selfPickIdx === -1 && h.indexOf('self pick up (360100)') !== -1) selfPickIdx = c;
        if (selfPickThirdIdx === -1 && h.indexOf('self pick up third party (360100)') !== -1) selfPickThirdIdx = c;
      }
      if (importIdx !== -1 || importBookToCreIdx !== -1 || selfPickIdx !== -1 || selfPickThirdIdx !== -1) {
        headerRow = r;
        break;
      }
    }

    if (headerRow === -1) {
      result.sourceNote = 'SC_Import Savings header row not detected';
      return result;
    }

    const importSourceIdx = (rdcKey === 'GRM' && importBookToCreIdx !== -1) ? importBookToCreIdx : importIdx;
    if (importSourceIdx === -1 && selfPickIdx === -1 && selfPickThirdIdx === -1) {
      result.sourceNote = 'SC_Import Savings required columns missing';
      return result;
    }

    let importTotal = 0;
    let selfPickTotal = 0;
    const appliedRows = [];
    const years = {};
    years[invStart.getFullYear()] = true;
    years[invEnd.getFullYear()] = true;
    const yearList = Object.keys(years).map(v => parseInt(v, 10));

    for (let r = headerRow + 1; r < data.length; r++) {
      const row = data[r] || [];
      const cwNum = parseInt(String(row[0] || '').trim(), 10);
      if (isNaN(cwNum) || cwNum <= 0 || cwNum > 53) continue;

      let bestOverlap = 0;
      for (let y = 0; y < yearList.length; y++) {
        const year = yearList[y];
        const cwStart = startOfIsoWeek_(year, cwNum);
        const cwEnd = new Date(cwStart.getFullYear(), cwStart.getMonth(), cwStart.getDate() + 6);
        const overlap = overlapDaysInclusive_(invStart, invEnd, cwStart, cwEnd);
        if (overlap > bestOverlap) bestOverlap = overlap;
      }

      if (bestOverlap <= 0) continue;
      const weight = bestOverlap / 7;

      const importVal = importSourceIdx !== -1 ? parseCurrency_(row[importSourceIdx]) : null;
      const selfPickVal = selfPickIdx !== -1 ? parseCurrency_(row[selfPickIdx]) : null;
      const selfPickThirdVal = selfPickThirdIdx !== -1 ? parseCurrency_(row[selfPickThirdIdx]) : null;

      const importAbs = importVal !== null ? Math.abs(importVal) : 0;
      const selfAbs = (selfPickVal !== null ? Math.abs(selfPickVal) : 0) + (selfPickThirdVal !== null ? Math.abs(selfPickThirdVal) : 0);

      importTotal += importAbs * weight;
      selfPickTotal += selfAbs * weight;

      if (importAbs > 0 || selfAbs > 0) {
        appliedRows.push(`CW${cwNum} x ${(weight * 100).toFixed(0)}%`);
      }
    }

    result.import471003 = importTotal;
    result.selfPickup360100 = selfPickTotal;
    result.sourceNote = appliedRows.length
      ? `Prorated from ${sheet.getName()} (${rdcKey}) ${appliedRows.join(', ')}`
      : 'No overlapping CW savings rows found';
    return result;
  } catch (e) {
    result.sourceNote = `SC_Import Savings read failed: ${e.message}`;
    return result;
  }
}

function buildTmstAllocation_(fileTmstData, glConfig) {
  const allocation = { total: 0, othersTotal: 0, transferByCostCenter: {} };
  if (!fileTmstData || fileTmstData.length === 0) return allocation;

  fileTmstData.forEach(row => {
    const amount = parseCurrency_(row[8]); // totalCost
    if (amount === null || amount <= 0) return;

    allocation.total += amount;

    const deliveryType = String(row[9] || '').trim().toUpperCase();
    const storeCode = String(row[4] || '').trim().toUpperCase();

    if (deliveryType === 'O') {
      allocation.othersTotal += amount;
      return;
    }

    // Transfers are keyed by transfer code (PT/GT/FT) and charged to destination warehouse.
    const transferCostCenter = getTransferCostCenterForRow_(deliveryType, storeCode, glConfig);
    if (!transferCostCenter) return;

    if (!allocation.transferByCostCenter[transferCostCenter]) {
      allocation.transferByCostCenter[transferCostCenter] = 0;
    }
    allocation.transferByCostCenter[transferCostCenter] += amount;
  });

  return allocation;
}

function getTransferCostCenterForRow_(deliveryType, storeCode, glConfig) {
  const typeText = String(deliveryType || '').toUpperCase();
  const storeText = String(storeCode || '').toUpperCase();
  const compactType = typeText.replace(/\s+/g, '');
  const compactStore = storeText.replace(/\s+/g, '');
  const costCenters = (glConfig && glConfig.costCenters) ? glConfig.costCenters : {};

  // First priority: explicit transfer code in Type column.
  if (/\bPT\b/.test(typeText) || compactType.indexOf('PT') === 0) {
    return costCenters.PYE || '70001';
  }
  if (/\bFT\b/.test(typeText) || compactType.indexOf('FT') === 0) {
    return costCenters.FRG || '50001';
  }
  if (/\bGT\b/.test(typeText) || compactType.indexOf('GT') === 0) {
    return costCenters.GRM || '60001';
  }

  // Fallback: infer destination warehouse from store/location text when transfer code is not explicit.
  if (/\bPT\b/.test(storeText) || compactStore.indexOf('PT') === 0 || storeText.indexOf('PERRYVILLE') !== -1) {
    return costCenters.PYE || '70001';
  }
  if (/\bFT\b/.test(storeText) || compactStore.indexOf('FT') === 0 || storeText.indexOf('FREDERICKSBURG') !== -1) {
    return costCenters.FRG || '50001';
  }
  if (/\bGT\b/.test(storeText) || compactStore.indexOf('GT') === 0 || storeText.indexOf('GRAHAM') !== -1) {
    return costCenters.GRM || '60001';
  }

  return '';
}

function sumObjectValues_(obj) {
  return Object.keys(obj || {}).reduce((sum, key) => sum + (parseFloat(obj[key]) || 0), 0);
}

function exportSheetToPDF(spreadsheetId, sheetId, pdfName, folder) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?exportFormat=pdf&format=pdf&size=letter&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false&gid=${sheetId}`;
  const token = ScriptApp.getOAuthToken();
  const options = { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true };
  
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    const blob = response.getBlob().setName(pdfName);
    folder.createFile(blob);
  } else {
    Logger.log(`[ERROR] PDF Export failed: ${response.getContentText()}`);
  }
}

// --- UNIVERSAL TMST & DISCREPANCY LOGIC ---

function getInvoiceExtractionConfig_(systemConfig, carrierConfig, carrierName) {
  const carrierKey = String(carrierName || '').toUpperCase();
  const carrierCfg = (carrierConfig && carrierConfig[carrierKey]) ? carrierConfig[carrierKey] : {};
  return {
    regexText: String(carrierCfg.invoicePattern || systemConfig.INVOICE_NUMBER_REGEX || '').trim(),
    preferredCell: String(systemConfig.INVOICE_NUMBER_CELL || '').trim(),
    fallbackMode: String(systemConfig.INVOICE_NUMBER_FALLBACK || 'USE_FILENAME').trim().toUpperCase()
  };
}

function extractInvoiceNumber(invoiceSS, fallbackName, extractionConfig) {
  extractionConfig = extractionConfig || {};
  const regexText = extractionConfig.regexText || 'Invoice\\s*(?:#|ID|Number)\\s*:?\\s*([A-Za-z0-9_-]+)';
  const preferredCell = extractionConfig.preferredCell || '';
  const fallbackMode = extractionConfig.fallbackMode || 'USE_FILENAME';

  let compiledRegex = null;
  try {
    compiledRegex = new RegExp(regexText, 'i');
  } catch (e) {
    Logger.log(`[WARNING] Invalid INVOICE_NUMBER_REGEX: ${regexText}. Falling back to default.`);
    compiledRegex = /Invoice\s*(?:#|ID|Number)\s*:?\s*([A-Za-z0-9_-]+)/i;
  }

  try {
    const sheets = invoiceSS.getSheets();

    if (preferredCell) {
      for (let s = 0; s < sheets.length; s++) {
        try {
          const v = String(sheets[s].getRange(preferredCell).getDisplayValue() || '').trim();
          if (!v) continue;
          const m = v.match(compiledRegex);
          if (m && m[1]) return m[1].trim();
          if (/^[A-Za-z0-9_-]{4,}$/.test(v)) return v;
        } catch (e) {
          // Ignore invalid range errors and continue.
        }
      }
    }
    
    for (let s = 0; s < sheets.length; s++) {
      const sheetName = sheets[s].getName().toLowerCase();
      if (!sheetName.includes('summary') && !sheetName.includes('invoice')) continue;
      
      const data = sheets[s].getDataRange().getValues();
      for (let r = 0; r < Math.min(data.length, 40); r++) {
        for (let c = 0; c < Math.min(data[r].length, 10); c++) {
          let cellVal = String(data[r][c]).trim();
          if (!cellVal) continue;
          
          let match = cellVal.match(compiledRegex);
          if (match && match[1]) return match[1].trim();
          
          if (cellVal.match(/^Invoice\s*(?:#|ID|Number)\s*:?$/i)) {
            for (let nextC = c + 1; nextC < Math.min(data[r].length, c + 4); nextC++) {
              let nextVal = String(data[r][nextC]).trim();
              if (nextVal) return nextVal;
            }
          }
        }
      }
    }
    
    const data = sheets[0].getDataRange().getValues();
    for (let r = 0; r < Math.min(data.length, 30); r++) {
      for (let c = 0; c < Math.min(data[r].length, 10); c++) {
          let cellVal = String(data[r][c]).trim();
          if (!cellVal) continue;
          
          let match = cellVal.match(compiledRegex);
          if (match && match[1]) return match[1].trim();
          
          if (cellVal.match(/^Invoice\s*(?:#|ID|Number)\s*:?$/i)) {
            for (let nextC = c + 1; nextC < Math.min(data[r].length, c + 4); nextC++) {
              let nextVal = String(data[r][nextC]).trim();
              if (nextVal) return nextVal;
            }
          }
      }
    }
  } catch (e) {
    Logger.log(`[WARNING] Failed to extract invoice number from sheet: ${e.message}`);
  }
  if (fallbackMode === 'REQUIRE_EXTRACTED') return '';
  return fallbackName;
}

function cleanTuNumber(tu) {
  if (!tu) return "";
  let cleanTu = String(tu).trim().toUpperCase();
  // Strip out "US" followed by exactly 4 digits (e.g., US0005) which represents the warehouse code
  cleanTu = cleanTu.replace(/^US\s*\d{4}/i, '');
  return cleanTu;
}

// Prevent estimate debug logs from flooding execution logs on large files.
let _estimateDebugLogCount = 0;
let _runtimeRuleConfig = {};
let _runtimeSystemConfig = {};

function configYesNo_(value, defaultYes) {
  if (value === undefined || value === null || String(value).trim() === '') return !!defaultYes;
  const raw = String(value).trim().toUpperCase();
  return !(raw === 'NO' || raw === 'FALSE' || raw === '0' || raw === 'OFF' || raw === 'DISABLED');
}

function logEstimateDebug_(message) {
  if (_estimateDebugLogCount >= 40) return;
  Logger.log(message);
  _estimateDebugLogCount++;
}

function extractInvoiceTourAndBaseTu_(safeTuNumber) {
  let invoiceTour = '1';
  let baseTu = String(safeTuNumber || '');
  if (baseTu.length >= 8) {
    const lastChar = baseTu.slice(-1);
    if (!isNaN(lastChar)) {
      if (lastChar === '0') { invoiceTour = '1'; baseTu = baseTu.slice(0, -1); }
      else if (lastChar === '1') { invoiceTour = '2'; baseTu = baseTu.slice(0, -1); }
      else if (lastChar === '2') { invoiceTour = '3'; baseTu = baseTu.slice(0, -1); }
      else if (lastChar === '3') { invoiceTour = '4'; baseTu = baseTu.slice(0, -1); }
    }
  }
  return { invoiceTour: invoiceTour, baseTu: baseTu };
}

function getShiftFlags_(shiftValue) {
  const text = String(shiftValue || '').toUpperCase().trim();
  return {
    raw: text,
    isShift1: text.includes('1') || text.includes('FIRST'),
    isShift2: text.includes('2') || text.includes('SECOND')
  };
}

function getTourFlags_(tourValue) {
  const text = String(tourValue || '').toUpperCase().trim();
  return {
    raw: text,
    isFirst:
      text === '1' ||
      text === '01' ||
      text.indexOf('FIRST') !== -1 ||
      text.indexOf('1ST') !== -1 ||
      /(^|[^0-9])0?1([^0-9]|$)/.test(text),
    isSecond:
      text === '2' ||
      text === '02' ||
      text.indexOf('SECOND') !== -1 ||
      text.indexOf('2ND') !== -1 ||
      /(^|[^0-9])0?2([^0-9]|$)/.test(text)
  };
}

function getHaulierRecordCandidates_(haulierInfo, safeTuNumber, baseTu, disableFuzzyTuMatch) {
  const seen = {};
  const candidates = [];

  function pushCandidate(record, mode) {
    if (!record) return;
    const key = String(record.rowNumber || '') + '|' + String(mode || '');
    if (seen[key]) return;
    seen[key] = true;
    candidates.push({ record: record, mode: mode });
  }

  const exactGroup = (haulierInfo.recordGroups && haulierInfo.recordGroups[safeTuNumber]) || [];
  if (exactGroup.length) {
    exactGroup.forEach(function(record) { pushCandidate(record, 'exact'); });
  } else if (haulierInfo.records && haulierInfo.records[safeTuNumber]) {
    pushCandidate(haulierInfo.records[safeTuNumber], 'exact');
  }

  const baseGroup = (haulierInfo.baseGroups && haulierInfo.baseGroups[baseTu]) || [];
  if (baseGroup.length) {
    baseGroup.forEach(function(record) { pushCandidate(record, 'base'); });
  } else if (haulierInfo.records && haulierInfo.records[baseTu]) {
    pushCandidate(haulierInfo.records[baseTu], 'base');
  }

  if (!disableFuzzyTuMatch && haulierInfo.recordGroups) {
    for (let haulierKey in haulierInfo.recordGroups) {
      const safeKey = String(haulierKey || '').trim().toUpperCase();
      if (safeKey.length <= 4) continue;
      if (!(safeTuNumber.includes(safeKey) || safeKey.includes(safeTuNumber) || baseTu.includes(safeKey) || safeKey.includes(baseTu))) continue;
      haulierInfo.recordGroups[haulierKey].forEach(function(record) { pushCandidate(record, 'fuzzy'); });
    }
  }

  return candidates;
}

function chooseBestHaulierCandidate_(candidates, invoiceShiftValue, invoiceTour, carrierKey, rules, ignoreShift2Vs1) {
  if (!candidates || !candidates.length) return null;

  const isHbStyleCarrier = (carrierKey === 'HB' || carrierKey === 'WERNER');
  const isSchCarrier = carrierKey === 'SCH' || carrierKey.indexOf('SCH') !== -1;
  let invoiceShiftStr = String(invoiceShiftValue || '').toUpperCase().trim();
  const hbBlankShiftTour2 =
    isRuleEnabled_(rules, 'HB_BLANK_SHIFT_TOUR2_EXPECT_FIRST_SECOND', carrierKey) &&
    isHbStyleCarrier && !invoiceShiftStr && invoiceTour === '2';
  if (hbBlankShiftTour2) invoiceShiftStr = 'FIRST';
  const invShift = getShiftFlags_(invoiceShiftStr);

  function modeWeight(mode) {
    if (mode === 'exact') return 12;
    if (mode === 'base') return 8;
    if (mode === 'fuzzy') return 3;
    return 0;
  }

  function scoreCandidate(candidate) {
    const record = candidate.record || {};
    const tour = getTourFlags_(record.tour);
    const shift = getShiftFlags_(record.shift);
    let score = modeWeight(candidate.mode);

    if (!tour.raw) score += 10;
    else if ((invoiceTour === '1' && tour.isFirst) || (invoiceTour === '2' && tour.isSecond)) score += 100;
    else if (tour.raw === invoiceTour || tour.raw.indexOf(invoiceTour) !== -1 || tour.raw.indexOf('0' + invoiceTour) !== -1) score += 70;

    if (isSchCarrier) {
      score += 20;
    } else if (!shift.raw) {
      score += 10;
    } else if (hbBlankShiftTour2 && shift.isShift1 && tour.isSecond) {
      score += 40;
    } else if ((invShift.isShift1 && shift.isShift1) || (invShift.isShift2 && shift.isShift2)) {
      score += 30;
    } else if (ignoreShift2Vs1 && invShift.isShift2 && shift.isShift1) {
      score += 18;
    }

    return score;
  }

  let best = null;
  let bestScore = -1;
  candidates.forEach(function(candidate) {
    const score = scoreCandidate(candidate);
    if (score > bestScore) {
      best = candidate;
      bestScore = score;
    }
  });

  return best;
}

function evaluateTmstRow(rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, carrierName, haulierInfo, discrepancyData, invoiceNumber, haulierUpdates, glConfig, estimateVarianceData) {
  let type = 'INVOICE'; let match = 'NO'; let shiftMatch = 'UNMATCHED'; let storeMatch = 'UNMATCHED'; let tourMatch = 'UNMATCHED';
  const safeTuNumber = cleanTuNumber(tuNumber);
  const carrierKey = String(carrierName || '').trim().toUpperCase();
  const isHbStyleCarrier = (carrierKey === 'HB' || carrierKey === 'WERNER');
  const disableFuzzyTuMatch = carrierKey === 'HB';
  const ignoreShift2Vs1 = configYesNo_((_runtimeSystemConfig || {}).IGNORE_INVOICE_SHIFT2_WHEN_HAULIER_SHIFT1, true);
  const rules = _runtimeRuleConfig || {};
  
  // Extract intelligent Tour number based on TU suffix (0/1/2/3 => Tour 1/2/3/4).
  // Supports both 8-digit and 9-digit BOL/TU formats.
  const tourInfo = extractInvoiceTourAndBaseTu_(safeTuNumber);
  const invoiceTour = tourInfo.invoiceTour;
  const baseTu = tourInfo.baseTu;

  let haulierRecord = null;
  let matchMode = 'none';
  const haulierMap = haulierInfo.records || {};
  
  Logger.log(`      -> Evaluating TU: '${tuNumber}' (Cleaned: '${safeTuNumber}', Base: '${baseTu}', Tour: ${invoiceTour})`);
  
  const candidateMatch = chooseBestHaulierCandidate_(
    getHaulierRecordCandidates_(haulierInfo, safeTuNumber, baseTu, disableFuzzyTuMatch),
    shift,
    invoiceTour,
    carrierKey,
    rules,
    ignoreShift2Vs1
  );
  if (candidateMatch) {
    haulierRecord = candidateMatch.record;
    matchMode = candidateMatch.mode;
  } else if (haulierMap[safeTuNumber]) {
    haulierRecord = haulierMap[safeTuNumber];
    matchMode = 'exact';
  }
  
  if (haulierRecord) {
    match = 'YES'; type = haulierRecord.deliveryType || 'MATCHED'; 
    
    // Strict Store matching (TEMPORARILY DISABLED PER USER REQUEST)
    const invoiceStoreStr = String(store).trim();
    const storeRegexMatch = invoiceStoreStr.match(/(?:US)?(\d{4})/i);
    const invoiceStoreNum = storeRegexMatch ? storeRegexMatch[1] : invoiceStoreStr;
    
    // TEMPORARILY DISABLED: Force to MATCHED so it doesn't block write-backs or trigger discrepancies
    storeMatch = 'MATCHED'; 
    /*
    if (String(haulierRecord.store).includes(invoiceStoreNum) || invoiceStoreNum === "") {
       storeMatch = 'MATCHED';
    } else {
       storeMatch = 'UNMATCHED';
    }
    */
    
    // Strict Tour matching (supports both numeric and FIRST/SECOND wording).
    let haulierTourStr = String(haulierRecord.tour).toUpperCase().trim();
    const hauTourIsFirst =
      haulierTourStr === '1' ||
      haulierTourStr === '01' ||
      haulierTourStr.indexOf('FIRST') !== -1 ||
      haulierTourStr.indexOf('1ST') !== -1 ||
      /(^|[^0-9])0?1([^0-9]|$)/.test(haulierTourStr);
    const hauTourIsSecond =
      haulierTourStr === '2' ||
      haulierTourStr === '02' ||
      haulierTourStr.indexOf('SECOND') !== -1 ||
      haulierTourStr.indexOf('2ND') !== -1 ||
      /(^|[^0-9])0?2([^0-9]|$)/.test(haulierTourStr);
    if (
      haulierTourStr === invoiceTour ||
      haulierTourStr.includes(invoiceTour) ||
      haulierTourStr.includes('0' + invoiceTour) ||
      (invoiceTour === '1' && hauTourIsFirst) ||
      (invoiceTour === '2' && hauTourIsSecond)
    ) {
      tourMatch = 'MATCHED';
    } else if (!haulierTourStr) {
        tourMatch = 'MATCHED'; // Default to matched if Haulier report has no Tour column configured
    } else {
        tourMatch = 'UNMATCHED';
    }
    
    // Strict Shift matching
    const originalInvoiceShiftStr = String(shift).toUpperCase().trim();
    let invoiceShiftStr = originalInvoiceShiftStr;
    let haulierShiftStr = String(haulierRecord.shift).toUpperCase();

    // HB/Werner specific rule:
    // if invoice shift is blank and TU indicates Tour 2 (e.g. ...90 vs ...91),
    // treat expected shift as FIRST (this is a second TOUR, not second SHIFT).
    const hbMissingShiftExpectFirstSecond =
      isRuleEnabled_(rules, 'HB_BLANK_SHIFT_TOUR2_EXPECT_FIRST_SECOND', carrierName) &&
      isHbStyleCarrier && !originalInvoiceShiftStr;
    if (hbMissingShiftExpectFirstSecond) invoiceShiftStr = 'FIRST';
    
    let invIsShift1 = invoiceShiftStr.includes('1') || invoiceShiftStr.includes('FIRST');
    let invIsShift2 = invoiceShiftStr.includes('2') || invoiceShiftStr.includes('SECOND');
    let hauIsShift1 = haulierShiftStr.includes('1') || haulierShiftStr.includes('FIRST');
    let hauIsShift2 = haulierShiftStr.includes('2') || haulierShiftStr.includes('SECOND');
    const isSchCarrier = carrierKey === 'SCH' || carrierKey.indexOf('SCH') !== -1;
    const ignoredShift2Vs1Mismatch = ignoreShift2Vs1 && invIsShift2 && hauIsShift1;
    const hbTour2FirstSecondMatch =
      (carrierKey === 'HB') &&
      hauIsShift1 &&
      hauTourIsSecond &&
      ((invoiceTour === '2' && invIsShift1) || !originalInvoiceShiftStr);
    
    if (isSchCarrier) {
      // SCH haulier extracts can be single-row per TU and not reliably split by shift.
      // Keep shift as informational for TMST, but do not block matching/write-back on shift.
      shiftMatch = 'MATCHED';
    } else if (hbTour2FirstSecondMatch) {
      // HB-specific behavior: Tour 2 can validly be represented as blank or Shift 1 on the invoice
      // while the haulier row shows First Shift / Second Tour.
      shiftMatch = 'MATCHED';
    } else if (ignoredShift2Vs1Mismatch) {
      // Optional rule: allow invoice Shift 2 against haulier Shift 1 when configured.
      shiftMatch = 'MATCHED';
    } else if ((invIsShift1 && hauIsShift1) || (invIsShift2 && hauIsShift2) || (invoiceShiftStr === haulierShiftStr && invoiceShiftStr !== "")) {
        shiftMatch = 'MATCHED';
    } else if (!haulierShiftStr) {
        shiftMatch = 'MATCHED'; // Default if haulier report has no Shift column
    } else {
        shiftMatch = 'UNMATCHED';
    }

    if (carrierKey === 'HB' && shiftMatch === 'UNMATCHED') {
      Logger.log(`[HB SHIFT OVERRIDE] RDC ${rdcName} TU ${tuNumber}: ignoring HB shift mismatch. Invoice Shift='${shift}' Haulier Shift='${haulierRecord.shift}'.`);
      shiftMatch = 'MATCHED';
    }

    if (hbTour2FirstSecondMatch) {
      tourMatch = 'MATCHED';
    }
    
    // Discrepancy Logging for valid TU but mismatched data
    if (storeMatch === 'UNMATCHED') discrepancyData.push([rdcName, date, tuNumber, haulierRecord.store, store, 'Store Mismatch']);
    if (shiftMatch === 'UNMATCHED') discrepancyData.push([rdcName, date, tuNumber, haulierRecord.shift, shift, 'Shift Mismatch']);
    if (tourMatch === 'UNMATCHED') discrepancyData.push([rdcName, date, tuNumber, haulierRecord.tour, invoiceTour, 'Tour Mismatch']);
    // Keep this mismatch non-blocking, but do not add a discrepancy row to reduce noise.

    const haulierDeliveryTypeStr = String(haulierRecord.deliveryType || '').toUpperCase().trim();
    const hasAssignedCarrierShift = invoiceShiftStr !== '' || String(shift || '').trim() !== '';
    const isCarrierCancelRoute = haulierDeliveryTypeStr === 'CNC' || haulierDeliveryTypeStr.indexOf('CNC') !== -1;
    if (
      isRuleEnabled_(rules, 'DELIVERY_TYPE_CNC_WITH_ASSIGNED_SHIFT', carrierName) &&
      isCarrierCancelRoute && hasAssignedCarrierShift
    ) {
      discrepancyData.push([
        rdcName,
        date,
        tuNumber,
        haulierRecord.deliveryType,
        String(shift || ''),
        'Route recorded as Carrier Cancel'
      ]);
      Logger.log(`[CNC ALERT] RDC ${rdcName} TU ${tuNumber}: haulier delivery type marked CNC but carrier invoice has shift='${shift}'.`);
    }

    // HB/Werner missing-shift validation:
    // when invoice shift is missing, haulier should show Shift FIRST + Tour SECOND.
    if (hbMissingShiftExpectFirstSecond && !(hauIsShift1 && hauTourIsSecond)) {
      discrepancyData.push([
        rdcName, date, tuNumber,
        `Haulier Shift='${haulierRecord.shift}' Tour='${haulierRecord.tour}'`,
        'Expected Shift FIRST + Tour SECOND',
        'HB Blank Shift / Tour Mapping Mismatch'
      ]);
      Logger.log(`[HB TOUR CHECK] RDC ${rdcName} TU ${tuNumber}: missing invoice shift expected FIRST/SECOND on haulier, found Shift='${haulierRecord.shift}' Tour='${haulierRecord.tour}'.`);
    }

    runCustomRules_(rules, {
      'carrier.name': carrierName,
      'invoice.tu': safeTuNumber,
      'invoice.baseTu': baseTu,
      'invoice.shift': String(shift || ''),
      'invoice.store': String(store || ''),
      'invoice.totalCost': String(totalCost || ''),
      'invoice.number': String(invoiceNumber || ''),
      'haulier.shift': String(haulierRecord.shift || ''),
      'haulier.store': String(haulierRecord.store || ''),
      'haulier.tour': String(haulierRecord.tour || ''),
      'haulier.deliveryType': String(haulierRecord.deliveryType || ''),
      'haulier.amount': String(haulierRecord.amount || ''),
      'calc.invoiceTour': invoiceTour,
      'calc.storeMatch': storeMatch,
      'calc.shiftMatch': shiftMatch,
      'calc.tourMatch': tourMatch,
      'calc.match': match,
      'calc.type': type
    }, discrepancyData, rdcName, date, tuNumber);

    // Amount tolerance check against existing haulier amount when present.
    let amountMatch = true;
    const invoiceAmount = parseCurrency_(totalCost);
    const haulierAmount = parseCurrency_(haulierRecord.amount);
    const pctTol = Number((glConfig && glConfig.tolerances && glConfig.tolerances.AMOUNT_MATCH_PCT) || 0.02);
    const hardTol = Number((glConfig && glConfig.tolerances && glConfig.tolerances.AMOUNT_MATCH_HARD) || 5.0);
    if (invoiceAmount !== null && haulierAmount !== null && haulierAmount > 0) {
      const allowedDelta = Math.max(hardTol, Math.abs(haulierAmount) * pctTol);
      const delta = Math.abs(invoiceAmount - haulierAmount);
      if (delta > allowedDelta) {
        amountMatch = false;
        discrepancyData.push([
          rdcName,
          date,
          tuNumber,
          `$${haulierAmount.toFixed(2)} ± $${allowedDelta.toFixed(2)}`,
          `$${invoiceAmount.toFixed(2)}`,
          'Amount Mismatch'
        ]);
      }
    }

    // Dedicated estimate vs actual variance alert (non-blocking).
    const expectedEstimateRaw = haulierRecord.estimate;
    const expectedEstimateDisplay = haulierRecord.estimateDisplay;
    const expectedEstimate = (parseCurrency_(expectedEstimateRaw) !== null)
      ? parseCurrency_(expectedEstimateRaw)
      : parseCurrency_(expectedEstimateDisplay);
    const estimateAlertThreshold = Number((glConfig && glConfig.tolerances && glConfig.tolerances.ESTIMATE_ALERT_HARD) || 200);

    if (expectedEstimate === null) {
      if (haulierRecord.hasEstimateFormula) {
        logEstimateDebug_(
          `[ESTIMATE DEBUG] RDC ${rdcName} TU ${tuNumber}: estimate formula present on haulier row ${haulierRecord.rowNumber} but parsed value is blank/non-numeric. Raw='${expectedEstimateRaw}' Display='${expectedEstimateDisplay}'`
        );
      } else {
        logEstimateDebug_(
          `[ESTIMATE DEBUG] RDC ${rdcName} TU ${tuNumber}: no parsable estimate value found on haulier row ${haulierRecord.rowNumber}. Raw='${expectedEstimateRaw}' Display='${expectedEstimateDisplay}'`
        );
      }
    }

    if (invoiceAmount !== null && expectedEstimate !== null && estimateAlertThreshold >= 0) {
      const variance = invoiceAmount - expectedEstimate;
      const absVariance = Math.abs(variance);
      if (absVariance > estimateAlertThreshold) {
        estimateVarianceData.push([
          rdcName,
          date,
          tuNumber,
          carrierName,
          invoiceNumber,
          expectedEstimate,
          invoiceAmount,
          variance,
          estimateAlertThreshold,
          absVariance,
          absVariance > estimateAlertThreshold ? 'ALERT' : 'OK'
        ]);
        logEstimateDebug_(
          `[ESTIMATE ALERT] RDC ${rdcName} TU ${tuNumber}: expected=$${expectedEstimate.toFixed(2)} actual=$${invoiceAmount.toFixed(2)} diff=$${variance.toFixed(2)} threshold=$${estimateAlertThreshold.toFixed(2)}`
        );
      }
    }

    // Push updates to be safely written to the Haulier Report later
    if (haulierUpdates && haulierUpdates[rdcName]) {
      const tmstWriteAmount = parseCurrency_(totalCost);
      let writeAmount = tmstWriteAmount;
      
      // Core rule: Block the amount write-back if NOT a perfect match.
      // Amount mismatch is deliberately excluded — it is informational only (used to verify
      // estimation formulas) and should not prevent the invoice amount being recorded.
      if (storeMatch === 'UNMATCHED' || tourMatch === 'UNMATCHED' || shiftMatch === 'UNMATCHED') {
         writeAmount = ""; 
      }
      
      haulierUpdates[rdcName].push({
        rowNumber: haulierRecord.rowNumber,
        invoiceNumber: invoiceNumber,
        amount: writeAmount,
        tmstAmount: writeAmount,
        tuNumber: safeTuNumber,
        shift: String(shift || ''),
        carrierName: carrierName,
        matchMode: matchMode
      });
    }

  } else {
    Logger.log(`      -> [DISCREPANCY ALERT] TU: '${safeTuNumber}' not found in Haulier report. Date: ${date}, Store: ${store}`);
    discrepancyData.push([rdcName, date, tuNumber, 'N/A', store, 'TU Not Found in Haulier Report']);
  }
  
  return [rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, type, match, shiftMatch, storeMatch, tourMatch, carrierName];
}

// --- CARRIER SPECIFIC MODULES ---

function processHbInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates) {
  invoiceSS.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return;
    
    let isShiftTab = false; let headerRowIndex = -1; let headers = [];
    for (let r = 0; r < Math.min(data.length, 10); r++) {
      if (rowHasHeader(data[r], 'TU', headerAliases, carrierName) && (rowHasHeader(data[r], 'Date', headerAliases, carrierName) || rowHasHeader(data[r], 'Shift', headerAliases, carrierName) || rowHasHeader(data[r], 'Store', headerAliases, carrierName))) {
        isShiftTab = true; headerRowIndex = r; headers = data[r]; break;
      }
    }
    if (!isShiftTab && sheetName.toLowerCase().includes('shift')) { isShiftTab = true; headerRowIndex = 0; headers = data[0]; }

    if (isShiftTab && headers.length > 0) {
      let dateIdx = -1, shiftIdx = -1, tuIdx = -1, storeIdx = -1, milesIdx = -1, nyPayIdx = -1, tollsIdx = -1, totalCostIdx = -1, shift1TotalCostIdx = -1, shift2TotalCostIdx = -1;
      
      // Pass 1: Exact matches first
      for (let j = 0; j < headers.length; j++) {
        const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
        if (shift1TotalCostIdx === -1 && (h === 'shift 1 total cost' || h === 'first shift total cost' || h === '1st shift total cost')) shift1TotalCostIdx = j;
        else if (shift2TotalCostIdx === -1 && (h === 'shift 2 total cost' || h === 'second shift total cost' || h === '2nd shift total cost')) shift2TotalCostIdx = j;
        else if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, carrierName, true)) totalCostIdx = j;
        else if (shiftIdx === -1 && matchHeader(h, 'Shift', headerAliases, carrierName, true) && !h.includes('total')) shiftIdx = j;
        else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, carrierName, true)) dateIdx = j;
        else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, carrierName, true)) tuIdx = j;
        else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, carrierName, true)) storeIdx = j;
        else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, carrierName, true)) milesIdx = j;
        else if (nyPayIdx === -1 && matchHeader(h, 'NY Pay', headerAliases, carrierName, true)) nyPayIdx = j;
        else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, carrierName, true)) tollsIdx = j;
      }

      // Pass 2: Partial Matches
      for (let j = 0; j < headers.length; j++) {
        const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
        if (shift1TotalCostIdx === -1 && ((h.includes('shift 1') || h.includes('first shift') || h.includes('1st shift')) && h.includes('total') && h.includes('cost'))) shift1TotalCostIdx = j;
        else if (shift2TotalCostIdx === -1 && ((h.includes('shift 2') || h.includes('second shift') || h.includes('2nd shift')) && h.includes('total') && h.includes('cost'))) shift2TotalCostIdx = j;
        else if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, carrierName, false)) totalCostIdx = j;
        else if (shiftIdx === -1 && matchHeader(h, 'Shift', headerAliases, carrierName, false) && !h.includes('total')) shiftIdx = j;
        else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, carrierName, false)) dateIdx = j;
        else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, carrierName, false)) tuIdx = j;
        else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, carrierName, false)) storeIdx = j;
        else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, carrierName, false)) milesIdx = j;
        else if (nyPayIdx === -1 && matchHeader(h, 'NY Pay', headerAliases, carrierName, false)) nyPayIdx = j;
        else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, carrierName, false)) tollsIdx = j;
      }

      for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        const date = dateIdx !== -1 ? row[dateIdx] : "";
        const tuNumber = tuIdx !== -1 ? cleanTuNumber(row[tuIdx]) : "";
        
        // Skip junk data, empty rows, or generic subtotal labels
        if (!tuNumber || String(tuNumber).includes('TOTAL') || String(tuNumber).length < 4) continue;
        
        const shift = shiftIdx !== -1 && row[shiftIdx] !== undefined ? row[shiftIdx] : "";
        const shiftUpper = String(shift || '').toUpperCase();
        const isShift2 = shiftUpper.includes('2') || shiftUpper.includes('SECOND');
        const isShift1OrBlank = !shiftUpper || shiftUpper.includes('1') || shiftUpper.includes('FIRST');
        const store = storeIdx !== -1 && row[storeIdx] !== undefined ? row[storeIdx] : "";
        const miles = milesIdx !== -1 && row[milesIdx] !== undefined ? row[milesIdx] : "";
        const nyPay = nyPayIdx !== -1 && row[nyPayIdx] !== undefined ? row[nyPayIdx] : "";
        const tolls = tollsIdx !== -1 && row[tollsIdx] !== undefined ? row[tollsIdx] : "";
        const shift1CostRaw = shift1TotalCostIdx !== -1 ? row[shift1TotalCostIdx] : "";
        const shift2CostRaw = shift2TotalCostIdx !== -1 ? row[shift2TotalCostIdx] : "";
        const totalCostRaw = totalCostIdx !== -1 ? row[totalCostIdx] : "";
        const shift1CostNum = parseCurrency_(shift1CostRaw);
        const shift2CostNum = parseCurrency_(shift2CostRaw);
        const totalCostNum = parseCurrency_(totalCostRaw);

        let totalCost = "";
        if (isShift2) {
          if (shift2CostNum !== null && shift2CostNum > 0) totalCost = shift2CostRaw;
          else if (totalCostNum !== null && totalCostNum > 0) totalCost = totalCostRaw;
          else if (shift1CostNum !== null && shift1CostNum > 0) totalCost = shift1CostRaw;
          else totalCost = shift2CostRaw || totalCostRaw || shift1CostRaw || "";
        } else {
          if (shift1CostNum !== null && shift1CostNum > 0) totalCost = shift1CostRaw;
          else if (totalCostNum !== null && totalCostNum > 0) totalCost = totalCostRaw;
          else if (shift2CostNum !== null && shift2CostNum > 0) totalCost = shift2CostRaw;
          else totalCost = shift1CostRaw || totalCostRaw || shift2CostRaw || "";
        }
        
        masterData.push([rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost]);
        tmstData.push(evaluateTmstRow(rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, carrierName, haulierInfo, discrepancyData, invoiceNumber, haulierUpdates, glConfig, estimateVarianceData));
      }
    }
    
    if (sheetName.toLowerCase().includes('invoice')) {
      Logger.log(`        -> Processing HB Invoice tab for additional costs...`);
      let costsFound = 0;
      let inChargesSection = false;
      
      for (let i = 0; i < data.length; i++) {
        let rowStr = data[i].join(" ").toLowerCase();

        if (isStampedCodingRow_(data[i])) continue;
        
        // Auto-detect start of charges to avoid reading invoice # or dates as costs
        if (!inChargesSection) {
           if (rowStr.includes('dedicated charges') || rowStr.includes('qty') || rowStr.includes('rate')) {
              inChargesSection = true;
           }
        }
        
        let desc = "";
        // Safely scan the first 3 columns for the description to handle empty column formatting issues
        for (let col = 0; col < Math.min(data[i].length, 3); col++) {
          if (String(data[i][col]).trim() !== "") {
            desc = String(data[i][col]).trim();
            break;
          }
        }
        
        let descLower = desc.toLowerCase();
        
        // Highly strict metadata skipping
        if (!desc || descLower.includes('total') || descLower.includes('invoice') || descLower.includes('week ending') || descLower.includes('diesel') || descLower.includes('truck count') || descLower.includes('@lidl') || descLower.includes('lidl us') || descLower.includes('arlington') || descLower.includes('http') || descLower === 'qty' || descLower.includes('dedicated charges')) continue;
        
        // Failsafe: if we haven't hit the charges section yet, only accept obvious charge lines
        if (!inChargesSection && !descLower.includes('charge') && !descLower.includes('shift') && !descLower.includes('mileage') && !descLower.includes('toll')) continue;
        
        // Apply dynamic Ignore Rules
        let shouldIgnore = false;
        if (glConfig && glConfig.ignoreKeywords) {
           for (let kw of glConfig.ignoreKeywords) {
              if (kw && descLower.includes(kw)) { shouldIgnore = true; break; }
           }
        }
        if (shouldIgnore) continue;
        
        let amount = 0;
        // Parse from right to left to grab the final numerical value
        for (let col = data[i].length - 1; col >= 0; col--) {
          const rawVal = String(data[i][col]).replace(/[^0-9.-]+/g, "");
          if (rawVal !== "" && rawVal !== "-") {
            const val = parseFloat(rawVal);
            if (!isNaN(val)) { amount = val; break; }
          }
        }
        
        // Apply dynamic Tolerance logic
        const tolerance = (glConfig && glConfig.tolerance) ? glConfig.tolerance : 0.01;
        if (amount >= tolerance) {
          fileCostSummary.push({ desc: desc, amount: amount });
          costsFound++;
          const isBaseCharge = /mileage|fuel|tolls|shift/i.test(desc);
          if (!isBaseCharge) additionalCostsData.push([rdcName, carrierName, desc, amount, invoiceSS.getName()]);
        }
      }
      Logger.log(`        -> Found ${costsFound} valid cost line items on HB Invoice tab.`);
    }
  });
}

function processCreInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates) {
  invoiceSS.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return;
    
    if (sheetName.toLowerCase().includes('detail')) {
      let headerRowIndex = -1; let headers = [];
      for (let r = 0; r < Math.min(data.length, 10); r++) {
        if (rowHasHeader(data[r], 'TU', headerAliases, carrierName) && rowHasHeader(data[r], 'Date', headerAliases, carrierName)) { headerRowIndex = r; headers = data[r]; break; }
      }
      
      if (headerRowIndex !== -1) {
        let dateIdx = -1, tuIdx = -1, storeIdx = -1, milesIdx = -1, nyPayIdx = -1, tollsIdx = -1, totalCostIdx = -1, firstShiftIdx = -1, secondShiftIdx = -1;
        
        // Pass 1: Exact matches
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
          if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, carrierName, true)) totalCostIdx = j;
          else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, carrierName, true)) dateIdx = j;
          else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, carrierName, true)) tuIdx = j;
          else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, carrierName, true)) storeIdx = j;
          else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, carrierName, true)) milesIdx = j;
          else if (nyPayIdx === -1 && matchHeader(h, 'NY Pay', headerAliases, carrierName, true)) nyPayIdx = j;
          else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, carrierName, true)) tollsIdx = j;
          else if (firstShiftIdx === -1 && h === 'first shift') firstShiftIdx = j;
          else if (secondShiftIdx === -1 && h === 'second shift') secondShiftIdx = j;
        }

        // Pass 2: Partial matches
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
          if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, carrierName, false)) totalCostIdx = j;
          else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, carrierName, false)) dateIdx = j;
          else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, carrierName, false)) tuIdx = j;
          else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, carrierName, false)) storeIdx = j;
          else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, carrierName, false)) milesIdx = j;
          else if (nyPayIdx === -1 && matchHeader(h, 'NY Pay', headerAliases, carrierName, false)) nyPayIdx = j;
          else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, carrierName, false)) tollsIdx = j;
          else if (firstShiftIdx === -1 && h.includes('first shift')) firstShiftIdx = j;
          else if (secondShiftIdx === -1 && h.includes('second shift')) secondShiftIdx = j;
        }

        for (let i = headerRowIndex + 1; i < data.length; i++) {
          const row = data[i];
          const date = dateIdx !== -1 ? row[dateIdx] : "";
          const tuNumber = tuIdx !== -1 ? cleanTuNumber(row[tuIdx]) : "";
          
          if (!tuNumber || String(tuNumber).includes('TOTAL') || String(tuNumber).length < 4) continue; 
          
          let shift = "Shift 1";
          if (firstShiftIdx !== -1 && parseFloat(row[firstShiftIdx]) > 0) shift = "Shift 1";
          else if (secondShiftIdx !== -1 && parseFloat(row[secondShiftIdx]) > 0) shift = "Shift 2";
          
          const store = storeIdx !== -1 && row[storeIdx] !== undefined ? row[storeIdx] : "";
          const miles = milesIdx !== -1 && row[milesIdx] !== undefined ? row[milesIdx] : "";
          const nyPay = nyPayIdx !== -1 && row[nyPayIdx] !== undefined ? row[nyPayIdx] : "";
          const tolls = tollsIdx !== -1 && row[tollsIdx] !== undefined ? row[tollsIdx] : "";
          const totalCost = totalCostIdx !== -1 && row[totalCostIdx] !== undefined ? row[totalCostIdx] : "";
          
          masterData.push([rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost]);
          tmstData.push(evaluateTmstRow(rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, carrierName, haulierInfo, discrepancyData, invoiceNumber, haulierUpdates, glConfig, estimateVarianceData));
        }
      }
    }
    
    if (sheetName.toLowerCase().includes('summary')) {
      Logger.log(`        -> Processing CRE Summary tab for additional costs...`);
      let headerRowIndex = -1; let headers = [];
      for (let r = 0; r < Math.min(data.length, 25); r++) {
        const rowStr = data[r].join(" ").toLowerCase();
        if (rowStr.includes('description') && rowStr.includes('total')) { headerRowIndex = r; headers = data[r]; break; }
      }
      
      if (headerRowIndex !== -1) {
        let descIdx = -1, totalIdx = -1;
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim();
          if (descIdx === -1 && h.includes('description')) descIdx = j;
          else if (totalIdx === -1 && (h === 'total' || h.includes('total'))) totalIdx = j;
        }
        
        let costsFound = 0;
        for (let i = headerRowIndex + 1; i < data.length; i++) {
          if (isStampedCodingRow_(data[i])) continue;

          let desc = "";
          for (let col = 0; col < Math.min(data[i].length, 3); col++) {
            if (String(data[i][col]).trim() !== "") { desc = String(data[i][col]).trim(); break; }
          }
          let descLower = desc.toLowerCase();
          
          if (!desc || descLower.includes('fixed charges') || descLower.includes('variable charges') || descLower === 'total') continue;
          
          // Apply dynamic Ignore Rules
          let shouldIgnore = false;
          if (glConfig && glConfig.ignoreKeywords) {
             for (let kw of glConfig.ignoreKeywords) {
                if (kw && descLower.includes(kw)) { shouldIgnore = true; break; }
             }
          }
          if (shouldIgnore) continue;

          // Use the correct Total column only (most reliable source)
          let amount = 0;
          if (totalIdx !== -1 && totalIdx < data[i].length) {
            // Use the Total column that was found in headers
            const rawVal = String(data[i][totalIdx]).replace(/[^0-9.-]+/g, "");
            if (rawVal !== "" && rawVal !== "-") {
              const val = parseFloat(rawVal);
              if (!isNaN(val)) { 
                // For CRE, reject values over 100k (these are daily charges, should be under 100k)
                if (val <= 100000) {
                  amount = val;
                }
              }
            }
          }
          
          // Fallback: if Total column not found, search only first 5 columns (avoid corrupted right-side data)
          if (amount === 0 && totalIdx === -1) {
            for (let col = 0; col < Math.min(data[i].length, 5); col++) {
              const rawVal = String(data[i][col]).replace(/[^0-9.-]+/g, "");
              if (rawVal !== "" && rawVal !== "-") {
                const val = parseFloat(rawVal);
                if (!isNaN(val) && val > 0 && val <= 100000) { amount = val; break; }
              }
            }
          }
          
          const tolerance = (glConfig && glConfig.tolerance) ? glConfig.tolerance : 0.01;
          if (!isNaN(amount) && amount >= tolerance) {
            fileCostSummary.push({ desc: desc, amount: amount });
            costsFound++;
            const isBaseCharge = /first shift|second shift|linehaul|tolls/i.test(desc);
            if (!isBaseCharge) additionalCostsData.push([rdcName, carrierName, desc, amount, invoiceSS.getName()]);
          }
        }
        Logger.log(`        -> Found ${costsFound} valid cost line items on CRE Summary tab.`);
      }
    }
  });
}

function processSchInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, estimateVarianceData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates) {
  // SCH invoices can contain duplicate route rows across multiple detail tabs
  // (e.g. Order Detail + Customer Detail). Track seen rows to avoid double counting.
  const seenSchDetailRows = {};
  invoiceSS.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return;
    
    if (sheetName.toLowerCase().includes('order detail') || sheetName.toLowerCase().includes('customer detail')) {
      let headerRowIndex = -1; let headers = [];
      for (let r = 0; r < Math.min(data.length, 10); r++) {
        if (rowHasHeader(data[r], 'TU', headerAliases, carrierName) && rowHasHeader(data[r], 'Miles', headerAliases, carrierName)) { headerRowIndex = r; headers = data[r]; break; }
      }
      
      if (headerRowIndex !== -1) {
        let dateIdx = -1, tuIdx = -1, storeIdx = -1, milesIdx = -1, nyPayIdx1 = -1, nyPayIdx2 = -1, tollsIdx = -1, totalAmtIdx = -1, drcAmtIdx = -1;
        
        // Pass 1: Exact matches
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
          if (totalAmtIdx === -1 && h === 'total $ amt') totalAmtIdx = j;
          else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, carrierName, true)) dateIdx = j;
          else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, carrierName, true)) tuIdx = j;
          else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, carrierName, true)) storeIdx = j;
          else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, carrierName, true)) milesIdx = j;
          else if (nyPayIdx1 === -1 && h === 'dhu $ amt') nyPayIdx1 = j;
          else if (nyPayIdx2 === -1 && h === 'stp $ amt') nyPayIdx2 = j;
          else if (drcAmtIdx === -1 && h === 'drc $ amt') drcAmtIdx = j;
          else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, carrierName, true)) tollsIdx = j;
        }

        // Pass 2: Partial matches
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
          if (totalAmtIdx === -1 && h.includes('total $ amt')) totalAmtIdx = j;
          else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, carrierName, false)) dateIdx = j;
          else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, carrierName, false)) tuIdx = j;
          else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, carrierName, false)) storeIdx = j;
          else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, carrierName, false)) milesIdx = j;
          else if (nyPayIdx1 === -1 && h.includes('dhu $ amt')) nyPayIdx1 = j;
          else if (nyPayIdx2 === -1 && h.includes('stp $ amt')) nyPayIdx2 = j;
          else if (drcAmtIdx === -1 && h.includes('drc $ amt')) drcAmtIdx = j;
          else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, carrierName, false)) tollsIdx = j;
        }

        for (let i = headerRowIndex + 1; i < data.length; i++) {
          const row = data[i];
          const tuNumber = tuIdx !== -1 ? cleanTuNumber(row[tuIdx]) : "";
          
          if (!tuNumber || String(tuNumber).includes('TOTAL') || String(tuNumber).length < 4) continue; 
          
          let date = dateIdx !== -1 ? row[dateIdx] : "";
          if (date && typeof date === 'string') date = date.split(' ')[0]; 
          
          // SCH shift logic: DRC $ Amt present with a non-zero value → Shift 1 (first shift).
          // If DRC $ Amt is blank or zero → Shift 2 (second shift).
          // If the DRC $ Amt column wasn't found at all, default to Shift 1.
          let shift = "Shift 1";
          if (drcAmtIdx !== -1) {
            const drcRaw = row[drcAmtIdx];
            const drcNum = parseFloat(String(drcRaw).replace(/[^0-9.-]/g, ''));
            shift = (!drcRaw || String(drcRaw).trim() === '' || isNaN(drcNum) || drcNum === 0) ? "Shift 2" : "Shift 1";
          }
          const store = storeIdx !== -1 && row[storeIdx] !== undefined ? row[storeIdx] : "";
          const miles = milesIdx !== -1 && row[milesIdx] !== undefined ? row[milesIdx] : "";
          const tolls = tollsIdx !== -1 && row[tollsIdx] !== undefined ? row[tollsIdx] : "";
          // SCH write-back source of truth: only use "Total $ Amt".
          const totalCost = totalAmtIdx !== -1 && row[totalAmtIdx] !== undefined ? row[totalAmtIdx] : "";
          
          let nyPay = 0;
          if (nyPayIdx1 !== -1 && row[nyPayIdx1]) nyPay += parseFloat(row[nyPayIdx1]) || 0;
          if (nyPayIdx2 !== -1 && row[nyPayIdx2]) nyPay += parseFloat(row[nyPayIdx2]) || 0;
          nyPay = nyPay > 0 ? nyPay : "";

          // Deduplicate identical SCH detail records across tabs to prevent doubled write-backs.
          const schDedupKey = [
            String(date || '').trim(),
            String(tuNumber || '').trim(),
            String(shift || '').trim(),
            String(store || '').trim(),
            String(totalCost || '').trim()
          ].join('|');
          if (seenSchDetailRows[schDedupKey]) continue;
          seenSchDetailRows[schDedupKey] = true;

          masterData.push([rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost]);
          tmstData.push(evaluateTmstRow(rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, carrierName, haulierInfo, discrepancyData, invoiceNumber, haulierUpdates, glConfig, estimateVarianceData));
        }
      }
    }
    
    if (sheetName.toLowerCase().includes('summary')) {
      Logger.log(`        -> Processing SCH Summary tab for additional costs...`);
      let headerRowIndex = -1; let headers = [];
      for (let r = 0; r < Math.min(data.length, 35); r++) {
        const rowStr = data[r].join(" ").toLowerCase();
        if (rowStr.includes('description') && rowStr.includes('total')) { headerRowIndex = r; headers = data[r]; break; }
      }
      
      if (headerRowIndex !== -1) {
        let descIdx = -1, totalIdx = -1;
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim();
          if (descIdx === -1 && h.includes('description')) descIdx = j;
          else if (totalIdx === -1 && h === 'total') totalIdx = j;
        }
        
        let costsFound = 0;
        for (let i = headerRowIndex + 1; i < data.length; i++) {
          if (isStampedCodingRow_(data[i])) continue;

          let desc = "";
          for (let col = 0; col < Math.min(data[i].length, 3); col++) {
            if (String(data[i][col]).trim() !== "") { desc = String(data[i][col]).trim(); break; }
          }
          let descLower = desc.toLowerCase();
          
          if (!desc || descLower === 'total') continue;
          
          // Apply dynamic Ignore Rules
          let shouldIgnore = false;
          if (glConfig && glConfig.ignoreKeywords) {
             for (let kw of glConfig.ignoreKeywords) {
                if (kw && descLower.includes(kw)) { shouldIgnore = true; break; }
             }
          }
          if (shouldIgnore) continue;
          
          let amount = 0;
          for (let col = data[i].length - 1; col >= 0; col--) {
            const rawVal = String(data[i][col]).replace(/[^0-9.-]+/g, "");
            if (rawVal !== "" && rawVal !== "-") {
              const val = parseFloat(rawVal);
              if (!isNaN(val)) { amount = val; break; }
            }
          }
          
          const tolerance = (glConfig && glConfig.tolerance) ? glConfig.tolerance : 0.01;
          if (!isNaN(amount) && amount >= tolerance) {
            fileCostSummary.push({ desc: desc, amount: amount });
            costsFound++;
            const isBaseCharge = /linehaul|driver charge|slip seat|fuel surcharge|tolls/i.test(desc);
            if (!isBaseCharge) additionalCostsData.push([rdcName, carrierName, desc, amount, invoiceSS.getName()]);
          }
        }
        Logger.log(`        -> Found ${costsFound} valid cost line items on SCH Summary tab.`);
      }
    }
  });
}

// --- HELPER FUNCTIONS ---

function isStampedCodingRow_(rowValues) {
  const rowText = rowValues.map(v => String(v || '')).join(' ').toLowerCase();

  if (rowText.includes('gl coding summary') || rowText.includes('gl account / cost center') || rowText.includes('total invoice amount')) {
    return true;
  }

  // Rows written by this script look like: "$123.45", "471000, 47100001, 60001"
  return /\b\d{6}\s*,\s*\d{6,8}\s*,\s*\d{4,6}\b/.test(rowText);
}

function fetchHaulierData(spreadsheetId, headerAliases, rdcLogName = "Unknown RDC") {
  const info = {
    spreadsheetId: spreadsheetId,
    sheetName: null,
    records: {},
    recordGroups: {},
    baseGroups: {},
    invoiceColIdx: -1,
    amountColIdx: -1,
    estimateColIdx: -1,
    maxCol: 0
  };
  if (!spreadsheetId) return info;
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = null;
    for (let s = 0; s < ss.getSheets().length; s++) {
      const name = ss.getSheets()[s].getName().toLowerCase();
      if (name.includes('weekly') && (name.includes('haulier') || name.includes('hauler'))) { sheet = ss.getSheets()[s]; break; }
    }
    if (!sheet) {
      Logger.log(`[WARNING] Haulier sheet for ${rdcLogName} is missing 'Weekly Haulier' tab.`);
      return info;
    }
    
    info.sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    const displayData = sheet.getDataRange().getDisplayValues();
    const formulaData = sheet.getDataRange().getFormulas();
    let tuIndex = -1, storeIndex = -1, typeIndex = -1, shiftIndex = -1, tourIndex = -1;
    let invCol = -1, amtCol = -1, estCol = -1;
    const headers = data[0];
    
    // Pass 1: Exact matches
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i]).toLowerCase().trim().replace(/\s+/g, ' ');
      if (typeIndex === -1 && matchHeader(h, 'Type', headerAliases, rdcLogName, true)) typeIndex = i;
      else if (storeIndex === -1 && matchHeader(h, 'Store', headerAliases, rdcLogName, true)) storeIndex = i;
      else if (tourIndex === -1 && matchHeader(h, 'Tour', headerAliases, rdcLogName, true)) tourIndex = i;
      else if (shiftIndex === -1 && matchHeader(h, 'Shift', headerAliases, rdcLogName, true)) shiftIndex = i;
      else if (tuIndex === -1 && matchHeader(h, 'TU', headerAliases, rdcLogName, true) && !h.includes('status') && !h.includes('return')) tuIndex = i;
      else if (invCol === -1 && matchHeader(h, 'Haulier Invoice', headerAliases, rdcLogName, true)) invCol = i;
      else if (amtCol === -1 && matchHeader(h, 'Haulier Amount', headerAliases, rdcLogName, true)) amtCol = i;
      else if (estCol === -1 && matchHeader(h, 'Haulier Estimate', headerAliases, rdcLogName, true)) estCol = i;
    }
    
    // Pass 2: Partial matches
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i]).toLowerCase().trim().replace(/\s+/g, ' ');
      if (typeIndex === -1 && matchHeader(h, 'Type', headerAliases, rdcLogName, false)) typeIndex = i;
      else if (storeIndex === -1 && matchHeader(h, 'Store', headerAliases, rdcLogName, false)) storeIndex = i;
      else if (tourIndex === -1 && matchHeader(h, 'Tour', headerAliases, rdcLogName, false)) tourIndex = i;
      else if (shiftIndex === -1 && matchHeader(h, 'Shift', headerAliases, rdcLogName, false)) shiftIndex = i;
      else if (tuIndex === -1 && matchHeader(h, 'TU', headerAliases, rdcLogName, false) && !h.includes('status') && !h.includes('return')) tuIndex = i;
      else if (invCol === -1 && matchHeader(h, 'Haulier Invoice', headerAliases, rdcLogName, false)) invCol = i;
      else if (amtCol === -1 && matchHeader(h, 'Haulier Amount', headerAliases, rdcLogName, false)) amtCol = i;
      else if (estCol === -1 && matchHeader(h, 'Haulier Estimate', headerAliases, rdcLogName, false)) estCol = i;
    }

    // Pass 3: Heuristic estimate-column detection for unconventional headers
    // (including common misspelling "deticated estimation").
    if (estCol === -1) {
      for (let i = 0; i < headers.length; i++) {
        const rawHeader = String(headers[i] || '').toLowerCase().trim();
        const compactHeader = rawHeader.replace(/[^a-z0-9]/g, '');
        const looksLikeEstimate =
          compactHeader.indexOf('estimate') !== -1 ||
          compactHeader.indexOf('estimation') !== -1 ||
          compactHeader.indexOf('dedicatedestimate') !== -1 ||
          compactHeader.indexOf('dedicatedestimation') !== -1 ||
          compactHeader.indexOf('deticatedestimate') !== -1 ||
          compactHeader.indexOf('deticatedestimation') !== -1 ||
          compactHeader === 'est' ||
          compactHeader.indexOf('expectedamount') !== -1;
        if (looksLikeEstimate) {
          estCol = i;
          Logger.log(`[ESTIMATE CONFIG] ${rdcLogName}: heuristic matched estimate column '${rawHeader}' at index ${i + 1}.`);
          break;
        }
      }
    }
    
    if (tuIndex === -1) tuIndex = 3; // Enforced fallback to Column D (Index 3)
    if (storeIndex === -1) storeIndex = 4; 
    
    // Apply RDC specific columns for Invoice write-backs
    if (invCol === -1) {
      invCol = (rdcLogName === 'GRM') ? 20 : 19; // U for GRM, T for FRG/PYE
    }
    if (amtCol === -1) {
      amtCol = (rdcLogName === 'GRM') ? 21 : 20; // V for GRM, U for FRG/PYE
    }
    
    info.invoiceColIdx = invCol;
    info.amountColIdx = amtCol;
    info.estimateColIdx = (estCol !== -1) ? estCol : amtCol;
    info.maxCol = headers.length;

    if (estCol === -1) {
      Logger.log(`[ESTIMATE CONFIG] ${rdcLogName}: no explicit estimate column found; falling back to Amount column index ${amtCol + 1}.`);
    }

    const estimateHeaderName = (info.estimateColIdx >= 0 && headers[info.estimateColIdx] !== undefined)
      ? String(headers[info.estimateColIdx])
      : '';
    Logger.log(
      `[ESTIMATE CONFIG] ${rdcLogName}: invoiceCol=${info.invoiceColIdx + 1}, amountCol=${info.amountColIdx + 1}, estimateCol=${info.estimateColIdx + 1} (${estimateHeaderName || 'n/a'})`
    );
    
    let loadedCount = 0;
    let estimatePresentCount = 0;
    let estimateFormulaCount = 0;
    let estimateBlankCount = 0;
    const map = {};
    for (let i = 1; i < data.length; i++) {
      if (!data[i][tuIndex]) continue;
      const tu = cleanTuNumber(data[i][tuIndex]);
      const estimateRaw = info.estimateColIdx !== -1 ? data[i][info.estimateColIdx] : '';
      const estimateDisplay = info.estimateColIdx !== -1 ? displayData[i][info.estimateColIdx] : '';
      const estimateFormula = info.estimateColIdx !== -1 ? formulaData[i][info.estimateColIdx] : '';
      const hasEstimateFormula = !!String(estimateFormula || '').trim();
      const hasEstimateValue = parseCurrency_(estimateRaw) !== null || parseCurrency_(estimateDisplay) !== null;
      if (hasEstimateFormula) estimateFormulaCount++;
      if (hasEstimateValue) estimatePresentCount++;
      else estimateBlankCount++;

      const record = {
        store: data[i][storeIndex],
        tour: tourIndex !== -1 ? String(data[i][tourIndex]).trim() : '',
        shift: shiftIndex !== -1 ? String(data[i][shiftIndex]).trim() : '',
        amount: amtCol !== -1 ? data[i][amtCol] : '',
        estimate: estimateRaw,
        estimateDisplay: estimateDisplay,
        hasEstimateFormula: hasEstimateFormula,
        deliveryType: typeIndex !== -1 ? String(data[i][typeIndex]).trim() : 'MATCHED',
        rawRow: data[i],
        rowNumber: i + 1 
      };

      const baseInfo = extractInvoiceTourAndBaseTu_(tu);
      if (!map[tu]) map[tu] = record;
      if (!info.recordGroups[tu]) info.recordGroups[tu] = [];
      info.recordGroups[tu].push(record);
      if (!info.baseGroups[baseInfo.baseTu]) info.baseGroups[baseInfo.baseTu] = [];
      info.baseGroups[baseInfo.baseTu].push(record);
      loadedCount++;
    }
    info.records = map;
    Logger.log(`      -> Loaded ${loadedCount} valid TU records from Haulier Report for ${rdcLogName}.`);
    Logger.log(
      `[ESTIMATE SUMMARY] ${rdcLogName}: estimate values parsed=${estimatePresentCount}, blank/non-numeric=${estimateBlankCount}, formula-cells=${estimateFormulaCount}`
    );
  } catch(e) {
    Logger.log(`[ERROR] Could not load haulier sheet ID ${spreadsheetId} for ${rdcLogName}: ${e.message}`);
  }
  return info;
}

function parseRDCName(text, rdcAliases) {
  const upper = text.toUpperCase();
  
  if (rdcAliases) {
     for (let rdc in rdcAliases) {
        for (let alias of rdcAliases[rdc]) {
           if (alias && upper.includes(alias)) return rdc;
        }
     }
  }
  
  // Fallback if aliases are missing or nothing matched
  if (upper.includes('FREDERICKSBURG') || upper.includes('FRG') || upper.includes(' VA') || upper.includes('LIDL VA') || upper.includes('_VA')) return 'FRG';
  if (upper.includes('GRAHAM') || upper.includes('GRM') || upper.includes(' NC') || upper.includes('LIDL NC') || upper.includes('_NC') || upper.includes('MEBANE')) return 'GRM';
  if (upper.includes('PERRYVILLE') || upper.includes('PYE') || upper.includes(' MD') || upper.includes('LIDL MD') || upper.includes('_MD') || upper.includes('PER')) return 'PYE';
  
  return 'UNKNOWN';
}

function convertExcelToGoogleSheet(excelFileId, parentFolderId) {
  const file = DriveApp.getFileById(excelFileId);
  const blob = file.getBlob();
  try {
    const resource = { name: "[TEMP] " + file.getName().replace(/\.xlsx$/i, ""), parents: [parentFolderId], mimeType: MimeType.GOOGLE_SHEETS };
    return Drive.Files.create(resource, blob, { supportsAllDrives: true }).id;
  } catch (e) {
    if (typeof Drive.Files.insert === "function") {
      const config = { title: "[TEMP] " + file.getName().replace(/\.xlsx$/i, ""), parents: [{id: parentFolderId}], mimeType: MimeType.GOOGLE_SHEETS };
      return Drive.Files.insert(config, blob, { supportsAllDrives: true }).id;
    } else { throw e; }
  }
}

function writeDataToSheet(ss, sheetName, dataArray, targetColumnCount) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    let headers = [];
    if (sheetName === 'Estimate Variance') {
      headers = [
        'RDC',
        'Date',
        'TU',
        'Carrier',
        'Invoice Number',
        'Expected Estimate',
        'Actual Amount',
        'Difference (Actual-Expected)',
        'Alert Threshold',
        'Abs Difference',
        'Status'
      ];
    } else {
      for(let i=1; i<=targetColumnCount; i++) headers.push("Column " + i);
    }
    sheet.appendRow(headers);
  }

  if (dataArray.length === 0) return;
  
  const cleanData = dataArray.map(row => {
    const newRow = [...row];
    while (newRow.length < targetColumnCount) newRow.push("");
    return newRow.slice(0, targetColumnCount);
  });
  
  sheet.getRange(sheet.getLastRow() + 1, 1, cleanData.length, targetColumnCount).setValues(cleanData);
}

function getArchiveSettings_() {
  let cfg = {};
  try { cfg = getConfig(); } catch (e) {}

  const prefix = String(cfg.ARCHIVE_SHEET_PREFIX || 'Archive - ').trim() || 'Archive - ';
  const dateFormat = String(cfg.ARCHIVE_DATE_FORMAT || 'yyyy-MM-dd').trim() || 'yyyy-MM-dd';
  const parsedRetention = parseInt(String(cfg.ARCHIVE_RETENTION_DAYS || '90').trim(), 10);
  const retentionDays = isNaN(parsedRetention) ? 90 : Math.max(0, parsedRetention);

  return { prefix: prefix, dateFormat: dateFormat, retentionDays: retentionDays };
}

function buildArchiveSnapshotName_(sheetName, archiveSettings) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), archiveSettings.dateFormat);
  return `${archiveSettings.prefix}${sheetName} ${stamp}`;
}

function pruneOldArchiveSnapshots_(ss, archiveSettings) {
  if (!archiveSettings || archiveSettings.retentionDays <= 0) return;

  const cutoffMs = Date.now() - (archiveSettings.retentionDays * 24 * 60 * 60 * 1000);
  const prefix = archiveSettings.prefix;
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(prefix)) return;

    const note = sheet.getRange(1, 1).getNote() || '';
    const m = note.match(/archiveCreatedEpoch=(\d+)/);
    if (!m) return;
    const createdEpoch = parseInt(m[1], 10);
    if (isNaN(createdEpoch)) return;
    if (createdEpoch < cutoffMs) ss.deleteSheet(sheet);
  });
}

function archiveAndClearSheets(ss) {
  const sheetsToProcess = ['Master Input', 'TMST', 'Additonal Costs', 'Discrepancy Tracker', 'Estimate Variance'];
  const archiveSettings = getArchiveSettings_();

  sheetsToProcess.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow > 1) {
        const dataToArchive = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        const archiveName = buildArchiveSnapshotName_(sheetName, archiveSettings);
        let archiveSheet = ss.getSheetByName(archiveName);
        if (!archiveSheet) {
          archiveSheet = ss.insertSheet(archiveName);
          const header = sheet.getRange(1, 1, 1, lastCol).getValues();
          archiveSheet.appendRow(header[0]);
          archiveSheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#efefef');
          archiveSheet.getRange(1, 1).setNote('archiveCreatedEpoch=' + Date.now());
        }
        archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, dataToArchive.length, lastCol).setValues(dataToArchive);
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    }
  });

  pruneOldArchiveSnapshots_(ss, archiveSettings);
}

function clearTrackerData(isWebApp = false) {
  if (!isWebApp) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Warning', 'Are you sure you want to clear and archive all current data in Master Input, TMST, Additional Costs, Discrepancy Tracker, and Estimate Variance?', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
      archiveAndClearSheets(SpreadsheetApp.getActiveSpreadsheet());
      ui.alert('Data archived and cleared successfully.');
    }
  } else {
    archiveAndClearSheets(SpreadsheetApp.getActiveSpreadsheet());
    return 'Data archived and cleared successfully.';
  }
}

// --- WEB APP LOGIC ---

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Invoice Automation Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getRowCount = (name) => {
    const s = ss.getSheetByName(name);
    return s ? Math.max(0, s.getLastRow() - 1) : 0;
  };

  let config = {};
  let configLoadOk = false;
  try { config = getConfig(); configLoadOk = true; } catch (e) {}

  // --- Build carrier list (primary RDCs first, then dynamic carriers) ---
  const primaryRdcs = ['FRG', 'GRM', 'PYE'];
  const seen = new Set();
  const carriers = [];

  primaryRdcs.forEach(rdc => {
    seen.add(rdc);
    const folderId  = String(config[rdc + '_ROOT_FOLDER']  || '').trim();
    const haulierId = String(config[rdc + '_HAULIER_ID']   || '').trim();
    const enabled = isCarrierEnabled_(config, rdc);
    carriers.push({
      name: rdc,
      classification: 'Primary RDC',
      primaryLane: true,
      enabled: enabled,
      rootFolderConfigured: !!folderId,
      haulierLinked: !!haulierId,
      folderId: folderId,
      haulierId: haulierId,
      driveUrl:   folderId  ? `https://drive.google.com/drive/folders/${folderId}`  : '',
      haulierUrl: haulierId ? `https://docs.google.com/spreadsheets/d/${haulierId}` : ''
    });
  });

  Object.keys(config).forEach(k => {
    if (!k.endsWith('_ROOT_FOLDER')) return;
    const name = k.replace('_ROOT_FOLDER', '');
    if (seen.has(name)) return;
    seen.add(name);
    const folderId  = String(config[k] || '').trim();
    const haulierId = String(config[name + '_HAULIER_ID'] || '').trim();
    const enabled = isCarrierEnabled_(config, name);
    carriers.push({
      name: name,
      classification: 'Dynamic Carrier',
      primaryLane: false,
      enabled: enabled,
      rootFolderConfigured: !!folderId,
      haulierLinked: !!haulierId,
      folderId: folderId,
      haulierId: haulierId,
      driveUrl:   folderId  ? `https://drive.google.com/drive/folders/${folderId}`  : '',
      haulierUrl: haulierId ? `https://docs.google.com/spreadsheets/d/${haulierId}` : ''
    });
  });

  // --- Readiness metrics ---
  const cfgTabs = ['System Config', 'GL Config', 'RDC Aliases', 'Email Template', 'Header Config', 'Carrier Config', 'Output Routing', 'Rule Config', 'GL Code Templates', 'Addl Cost Config'];
  const configReadyCount       = cfgTabs.filter(n => !!ss.getSheetByName(n)).length;
  const enabledCarrierCount    = carriers.filter(c => c.enabled).length;
  const configuredCarrierCount = carriers.filter(c => c.enabled && c.rootFolderConfigured).length;
  const linkedHaulierCount     = carriers.filter(c => c.enabled && c.haulierLinked).length;

  const masterRows    = getRowCount('Master Input');
  const tmstRows      = getRowCount('TMST');
  const discrepancies = getRowCount('Discrepancy Tracker');
  const addlCosts     = getRowCount('Additonal Costs');
  const estimateVarianceRows = getRowCount('Estimate Variance');
  const discRate      = masterRows > 0 ? Math.round((discrepancies / masterRows) * 100) : 0;

  let score = 0;
  score += Math.round((configReadyCount / 10) * 40);
  score += Math.round((Math.min(configuredCarrierCount, 2) / 2) * 30);
  score += Math.round((Math.min(linkedHaulierCount, 2) / 2) * 30);

  // --- Alerts ---
  const alerts = [];
  if (!configLoadOk || configReadyCount < 10) {
    alerts.push({ severity: 'warning', title: 'Configuration Incomplete', detail: `${10 - configReadyCount} config tab(s) missing. Run Initialization.` });
  }
  if (enabledCarrierCount === 0) {
    alerts.push({ severity: 'critical', title: 'No Carriers Enabled', detail: 'Set at least one *_ENABLED key to YES in System Config.' });
  } else if (configuredCarrierCount === 0) {
    alerts.push({ severity: 'critical', title: 'No Enabled Carriers Configured', detail: 'Add at least one enabled carrier root folder in System Config.' });
  } else if (configReadyCount >= 8) {
    alerts.push({ severity: 'success', title: 'System Ready', detail: `${configuredCarrierCount} enabled carrier(s) configured, ${linkedHaulierCount} haulier link(s) active.` });
  }
  if (discRate > 20) {
    alerts.push({ severity: 'warning', title: 'High Discrepancy Rate', detail: `${discRate}% of master rows have unresolved discrepancies.` });
  }
  if (estimateVarianceRows > 0) {
    alerts.push({ severity: 'warning', title: 'Estimate Variance Alerts', detail: `${estimateVarianceRows} TU row(s) exceed estimate-vs-actual threshold.` });
  }

  // --- Sheet summary helper ---
  const sheetObj = (name, label, desc) => {
    const s   = ss.getSheetByName(name);
    const rows = s ? Math.max(0, s.getLastRow() - 1) : 0;
    const gid  = s ? s.getSheetId() : null;
    return {
      name, label, description: desc,
      exists: !!s, rows,
      openUrl:  s ? `${ss.getUrl()}#gid=${gid}` : null,
      embedUrl: s ? `https://docs.google.com/spreadsheets/d/${ss.getId()}/htmlembed?gid=${gid}` : null
    };
  };

  const operationsSheets = [
    sheetObj('Master Input',        'Master Input',        'All extracted invoice line items'),
    sheetObj('TMST',                'TMST',                'TU reconciliation with haulier data'),
    sheetObj('Additonal Costs',     'Additional Costs',    'Non-base cost line items'),
    sheetObj('Discrepancy Tracker', 'Discrepancy Tracker', 'Flagged mismatches requiring review'),
    sheetObj('Estimate Variance',   'Estimate Variance',   'Expected estimate vs actual amount alerts')
  ];

  const configSheets = [
    sheetObj('System Config',  'System Config',  'Carrier folders and haulier sheet IDs'),
    sheetObj('GL Config',      'GL Config',      'Cost categories, ignore rules and tolerance'),
    sheetObj('RDC Aliases',    'RDC Aliases',    'RDC code name aliases for file matching'),
    sheetObj('Email Template', 'Email Template', 'Discrepancy notification email template'),
    sheetObj('Header Config',  'Header Config',  'Column aliases with optional carrier-specific overrides'),
    sheetObj('Carrier Config', 'Carrier Config', 'Carrier parser mapping and optional invoice pattern'),
    sheetObj('Output Routing', 'Output Routing', 'Route specific data types to different sheets/workbooks'),
    sheetObj('Rule Config',    'Rule Config',    'Toggle business rules and carrier-specific discrepancy checks'),
    sheetObj('Addl Cost Config', 'Addl Cost Config', 'Map invoice items to haulier Additional Cost sheet columns')
  ];

  const viewerSheets = operationsSheets.filter(s => s.exists);

  // --- Console seed ---
  const consoleSeed = [];
  if (configReadyCount < 8) consoleSeed.push({ message: `${8 - configReadyCount} configuration tab(s) missing — run Initialization.`, type: 'warning' });
  if (masterRows > 0) consoleSeed.push({ message: `Master Input: ${masterRows} rows loaded.`, type: 'info' });
  if (discrepancies > 0) consoleSeed.push({ message: `${discrepancies} discrepanc${discrepancies === 1 ? 'y' : 'ies'} pending resolution.`, type: 'warning' });
  if (consoleSeed.length === 0) consoleSeed.push({ message: 'Dashboard loaded. No issues detected.', type: 'success' });

  return {
    masterRows,
    tmstRows,
    discrepancies,
    additionalCosts: addlCosts,
    estimateVarianceAlerts: estimateVarianceRows,
    carriers,
    summary: {
      readinessScore: score,
      discrepancyRate: discRate,
      configReadyCount,
      enabledCarrierCount,
      configuredCarrierCount,
      primaryCarrierCount: primaryRdcs.length,
      linkedHaulierCount
    },
    alerts,
    operationsSheets,
    configSheets,
    viewerSheets,
    consoleSeed,
    workbook: { name: ss.getName(), openUrl: ss.getUrl() }
  };
}

function uploadInvoiceWeb(base64Data, filename, mimeType, carrierKey) {
  try {
    const config = getConfig();
    if (!isCarrierEnabled_(config, carrierKey)) {
      throw new Error('Carrier is disabled in System Config. Set ' + carrierKey + '_ENABLED to YES to allow uploads.');
    }
    const folderId = config[carrierKey + '_ROOT_FOLDER'];
    if (!folderId) throw new Error("Carrier root folder not configured.");
    
    const folder = DriveApp.getFolderById(folderId);
    
    // Decode base64 
    const splitBase = base64Data.split(',');
    const data = splitBase.length > 1 ? splitBase[1] : splitBase[0];
    
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, filename);
    folder.createFile(blob);
    return { success: true, message: filename + " uploaded successfully to " + carrierKey + " root folder!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function runProcessorWeb() {
  return runMainProcess(true);
}

function runProcessorExtractOnlyWeb() {
  return runMainProcessExtractOnly(true);
}

function finalizePendingInvoicesWeb() {
  return finalizePendingInvoices(true);
}

function clearTrackerWeb() {
  return clearTrackerData(true);
}

function setupConfigWeb() {
  try {
    // Pass `true` to indicate web-app context so `setupConfigTab` can avoid UI-only calls.
    setupConfigTab(true);
    return { success: true, message: 'Configuration tabs initialized successfully.' };
  } catch (e) {
    return { success: false, message: 'Initialization error: ' + e.message };
  }
}

function validateConfigWeb() {
  try {
    return validateConfiguration_();
  } catch (e) {
    return {
      success: false,
      summary: 'Validation failed with an exception.',
      checks: [{ status: 'error', item: 'Validator', detail: e.message }]
    };
  }
}

function validateConfiguration_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checks = [];
  const add = (status, item, detail) => checks.push({ status: status, item: item, detail: detail });

  const requiredTabs = ['System Config', 'GL Config', 'RDC Aliases', 'Email Template', 'Header Config', 'Carrier Config', 'Output Routing', 'Rule Config'];
  const missingTabs = requiredTabs.filter(n => !ss.getSheetByName(n));
  if (missingTabs.length === 0) add('ok', 'Required Tabs', `All ${requiredTabs.length} required tabs found.`);
  else add('error', 'Required Tabs', `Missing: ${missingTabs.join(', ')}`);

  let config = {};
  try {
    config = getConfig();
    add('ok', 'System Config', 'System Config loaded successfully.');
  } catch (e) {
    add('error', 'System Config', e.message);
    return { success: false, summary: 'Validation incomplete due to missing System Config.', checks: checks };
  }

  const headerSheet = ss.getSheetByName('Header Config');
  if (headerSheet) {
    const cols = headerSheet.getLastColumn();
    if (cols >= 3) add('ok', 'Header Config', 'Carrier override column is present.');
    else add('warn', 'Header Config', 'Carrier override column is missing. Re-run initialization to update schema.');
  }

  const ruleSheet = ss.getSheetByName('Rule Config');
  if (ruleSheet) {
    const cols = ruleSheet.getLastColumn();
    if (cols >= 10) add('ok', 'Rule Config', 'Expanded rule schema is present.');
    else add('warn', 'Rule Config', 'Rule Config is using a legacy schema. Re-run initialization to upgrade it.');
  }

  const carrierCfg = getCarrierConfig();
  const invalidProcessors = Object.keys(carrierCfg).filter(k => ['HB', 'CRE', 'SCH'].indexOf(String(carrierCfg[k].processorType || '').toUpperCase()) === -1);
  if (invalidProcessors.length === 0) add('ok', 'Carrier Config', 'All processor types are valid (HB/CRE/SCH).');
  else add('error', 'Carrier Config', `Invalid processor type for: ${invalidProcessors.join(', ')}`);

  const regexText = String(config.INVOICE_NUMBER_REGEX || '').trim();
  if (regexText) {
    try {
      new RegExp(regexText, 'i');
      add('ok', 'Invoice Regex', 'INVOICE_NUMBER_REGEX compiles successfully.');
    } catch (e) {
      add('error', 'Invoice Regex', `Regex is invalid: ${e.message}`);
    }
  } else {
    add('warn', 'Invoice Regex', 'INVOICE_NUMBER_REGEX is blank; default extraction fallback will be used.');
  }

  const fallbackMode = String(config.INVOICE_NUMBER_FALLBACK || 'USE_FILENAME').trim().toUpperCase();
  if (fallbackMode === 'USE_FILENAME' || fallbackMode === 'REQUIRE_EXTRACTED') {
    add('ok', 'Invoice Fallback', `INVOICE_NUMBER_FALLBACK is set to ${fallbackMode}.`);
  } else {
    add('warn', 'Invoice Fallback', `Unknown fallback mode: ${fallbackMode}. Recommended: USE_FILENAME or REQUIRE_EXTRACTED.`);
  }

  const glCfg = getGlConfig() || {};
  const minCost = Number((glCfg.tolerances || {}).MIN_COST);
  const pctTol = Number((glCfg.tolerances || {}).AMOUNT_MATCH_PCT);
  const hardTol = Number((glCfg.tolerances || {}).AMOUNT_MATCH_HARD);
  const estimateAlertHard = Number((glCfg.tolerances || {}).ESTIMATE_ALERT_HARD);
  if (!isNaN(minCost) && minCost >= 0) add('ok', 'Tolerance MIN_COST', `MIN_COST = ${minCost}`);
  else add('warn', 'Tolerance MIN_COST', 'MIN_COST is invalid or missing.');
  if (!isNaN(pctTol) && pctTol >= 0) add('ok', 'Tolerance AMOUNT_MATCH_PCT', `AMOUNT_MATCH_PCT = ${pctTol}`);
  else add('warn', 'Tolerance AMOUNT_MATCH_PCT', 'AMOUNT_MATCH_PCT is invalid or missing.');
  if (!isNaN(hardTol) && hardTol >= 0) add('ok', 'Tolerance AMOUNT_MATCH_HARD', `AMOUNT_MATCH_HARD = ${hardTol}`);
  else add('warn', 'Tolerance AMOUNT_MATCH_HARD', 'AMOUNT_MATCH_HARD is invalid or missing.');
  if (!isNaN(estimateAlertHard) && estimateAlertHard >= 0) add('ok', 'Tolerance ESTIMATE_ALERT_HARD', `ESTIMATE_ALERT_HARD = ${estimateAlertHard}`);
  else add('warn', 'Tolerance ESTIMATE_ALERT_HARD', 'ESTIMATE_ALERT_HARD is invalid or missing.');

  const retention = parseInt(String(config.ARCHIVE_RETENTION_DAYS || '90').trim(), 10);
  if (isNaN(retention) || retention < 0) add('warn', 'Archive Retention', 'ARCHIVE_RETENTION_DAYS should be a non-negative integer.');
  else add('ok', 'Archive Retention', `Retention set to ${retention} day(s).`);

  Object.keys(config).forEach(key => {
    const val = String(config[key] || '').trim();
    if (!val) return;
    if (key.endsWith('_ENABLED')) {
      const upper = val.toUpperCase();
      if (['YES', 'NO', 'TRUE', 'FALSE', '1', '0', 'ON', 'OFF', 'DISABLED'].indexOf(upper) === -1) {
        add('warn', key, 'Enabled flag should be YES/NO (accepted: YES/NO/TRUE/FALSE/1/0/ON/OFF).');
      } else {
        add('ok', key, `Enabled flag set to ${upper}.`);
      }
      return;
    }
    if (key.endsWith('_ROOT_FOLDER')) {
      try {
        DriveApp.getFolderById(val).getName();
        add('ok', key, 'Folder ID is accessible.');
      } catch (e) {
        add('error', key, 'Folder ID is not accessible or invalid.');
      }
    }
    if (key.endsWith('_HAULIER_ID')) {
      try {
        SpreadsheetApp.openById(val).getName();
        add('ok', key, 'Spreadsheet ID is accessible.');
      } catch (e) {
        add('error', key, 'Spreadsheet ID is not accessible or invalid.');
      }
    }
  });

  const routing = getConfigData().outputRouting || [];
  routing.forEach((row, idx) => {
    const type = String(row[0] || '').trim();
    const targetId = String(row[2] || '').trim();
    const enabled = String(row[3] || 'YES').trim().toUpperCase();
    if (!type || enabled === 'NO' || enabled === 'FALSE') return;
    if (!targetId) return;
    try {
      SpreadsheetApp.openById(targetId).getName();
      add('ok', `Output Routing Row ${idx + 2}`, `${type}: target spreadsheet is accessible.`);
    } catch (e) {
      add('error', `Output Routing Row ${idx + 2}`, `${type}: target spreadsheet ID is invalid or inaccessible.`);
    }
  });

  const ruleRows = getConfigData().ruleConfig || [];
  const seenRuleKeys = {};
  ruleRows.forEach((row, idx) => {
    const ruleKey = String(row[0] || '').trim().toUpperCase();
    const ruleType = String(row[1] || '').trim().toUpperCase();
    const condition1 = String(row[4] || '').trim();
    const condition2 = String(row[5] || '').trim();
    if (!ruleKey) return;

    seenRuleKeys[ruleKey] = (seenRuleKeys[ruleKey] || 0) + 1;
    if (ruleType !== 'SYSTEM' && ruleType !== 'CUSTOM') {
      add('warn', `Rule Row ${idx + 2}`, `${ruleKey}: Rule Type should be SYSTEM or CUSTOM.`);
    }
    if (ruleType === 'CUSTOM') {
      if (!condition1) add('warn', `Rule Row ${idx + 2}`, `${ruleKey}: Custom rules should include Condition 1.`);
      const parsedOne = parseRuleCondition_(condition1);
      const parsedTwo = parseRuleCondition_(condition2);
      if (condition1 && (!parsedOne || parsedOne.invalid)) add('warn', `Rule Row ${idx + 2}`, `${ruleKey}: Condition 1 is invalid.`);
      if (condition2 && (!parsedTwo || parsedTwo.invalid)) add('warn', `Rule Row ${idx + 2}`, `${ruleKey}: Condition 2 is invalid.`);
    }
  });

  Object.keys(seenRuleKeys).forEach(ruleKey => {
    if (seenRuleKeys[ruleKey] > 1) add('warn', 'Rule Config', `${ruleKey} appears ${seenRuleKeys[ruleKey]} times. Prefer unique keys for maintainability.`);
  });

  const errorCount = checks.filter(c => c.status === 'error').length;
  const warnCount = checks.filter(c => c.status === 'warn').length;
  const success = errorCount === 0;
  const summary = success
    ? `Validation passed with ${warnCount} warning(s).`
    : `Validation found ${errorCount} error(s) and ${warnCount} warning(s).`;

  return { success: success, summary: summary, checks: checks };
}

// --- LIVE CONFIG READ/WRITE (used by the web UI editor) ---

function getConfigData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const readRows = (name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    return data.slice(1); // exclude header row
  };
  const readAll = (name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return null;
    return sheet.getDataRange().getValues(); // include header (used for Email Template)
  };
  const readRuleRows = () => {
    const sheet = ss.getSheetByName('Rule Config');
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    return data.slice(1).map(normalizeRuleRow_);
  };
  return {
    systemConfig:  readRows('System Config'),
    glConfig:      readRows('GL Config'),
    rdcAliases:    readRows('RDC Aliases'),
    emailTemplate: readAll('Email Template'),
    headerConfig:  readRows('Header Config'),
    carrierConfig: readRows('Carrier Config'),
    outputRouting: readRows('Output Routing'),
    ruleConfig:    readRuleRows(),
    ruleCatalog:   getRuleCatalog_(),
    glCodeTemplates: readRows('GL Code Templates'),
    addlCostConfig: readRows('Addl Cost Config')
  };
}

function saveConfigData(sheetName, rows) {
  try {
    // Only allow saving to known configuration sheets
    const editableSheets = [
      'System Config',
      'GL Config',
      'RDC Aliases',
      'Email Template',
      'Header Config',
      'Carrier Config',
      'Output Routing',
      'Rule Config',
      'GL Code Templates',
      'Addl Cost Config'
    ];
    if (editableSheets.indexOf(sheetName) === -1) {
      return { success: false, message: `Saving to sheet "${sheetName}" is not allowed.` };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `Sheet "${sheetName}" not found. Run Initialization first.` };
    }
    if (sheetName === 'Rule Config') ensureRuleConfigSchema_(sheet);
    // Clear existing data rows (preserve header row 1)
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
    // Write new data rows
    if (rows && rows.length > 0) {
      const cols = rows[0].length;
      sheet.getRange(2, 1, rows.length, cols).setValues(rows);
    }
    return { success: true, message: `"${sheetName}" saved (${rows ? rows.length : 0} rows).` };
  } catch (e) {
    return { success: false, message: 'Save failed: ' + e.message };
  }
}

// --- INVOICE RESULTS VIEWER (Web UI) ---

function getInvoiceResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];

  const masterSheet = ss.getSheetByName('Master Input');
  const tmstSheet = ss.getSheetByName('TMST');
  const costsSheet = ss.getSheetByName('Additonal Costs');
  const discSheet = ss.getSheetByName('Discrepancy Tracker');

  if (!tmstSheet || tmstSheet.getLastRow() < 2) return results;

  const tmstData = tmstSheet.getRange(2, 1, tmstSheet.getLastRow() - 1, tmstSheet.getLastColumn()).getValues();
  const masterData = masterSheet && masterSheet.getLastRow() > 1
    ? masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues()
    : [];
  const costsData = costsSheet && costsSheet.getLastRow() > 1
    ? costsSheet.getRange(2, 1, costsSheet.getLastRow() - 1, costsSheet.getLastColumn()).getValues()
    : [];
  const discData = discSheet && discSheet.getLastRow() > 1
    ? discSheet.getRange(2, 1, discSheet.getLastRow() - 1, discSheet.getLastColumn()).getValues()
    : [];

  // Group TMST rows by carrier (col 14) + RDC (col 0) as a proxy for each invoice file
  const invoiceMap = {};
  tmstData.forEach((row, idx) => {
    const rdc = String(row[0] || '').trim();
    const carrier = String(row[14] || '').trim();
    const key = carrier + '|' + rdc;
    if (!invoiceMap[key]) {
      invoiceMap[key] = { carrier, rdc, tmstRows: [], masterRows: [], costRows: [], discRows: [] };
    }
    invoiceMap[key].tmstRows.push(row);
  });

  // Attach master rows by RDC
  masterData.forEach(row => {
    const rdc = String(row[0] || '').trim();
    for (let key in invoiceMap) {
      if (key.endsWith('|' + rdc)) {
        invoiceMap[key].masterRows.push(row);
        break;
      }
    }
  });

  // Attach cost rows by RDC + carrier
  costsData.forEach(row => {
    const rdc = String(row[0] || '').trim();
    const carrier = String(row[1] || '').trim();
    const key = carrier + '|' + rdc;
    if (invoiceMap[key]) invoiceMap[key].costRows.push(row);
  });

  // Attach discrepancy rows by RDC
  discData.forEach(row => {
    const rdc = String(row[0] || '').trim();
    for (let key in invoiceMap) {
      if (key.endsWith('|' + rdc)) {
        invoiceMap[key].discRows.push(row);
        break;
      }
    }
  });

  // Build result summaries
  for (let key in invoiceMap) {
    const inv = invoiceMap[key];
    const totalTUs = inv.tmstRows.length;
    const matchedTUs = inv.tmstRows.filter(r => String(r[10]) === 'YES').length;
    const unmatchedTUs = totalTUs - matchedTUs;
    const shiftMatched = inv.tmstRows.filter(r => String(r[11]) === 'MATCHED').length;
    const storeMatched = inv.tmstRows.filter(r => String(r[12]) === 'MATCHED').length;
    const tourMatched = inv.tmstRows.filter(r => String(r[13]) === 'MATCHED').length;

    let totalCost = 0;
    let totalMiles = 0;
    let totalNY = 0;
    let totalTolls = 0;
    inv.tmstRows.forEach(r => {
      totalCost += parseFloat(r[8]) || 0;
      totalMiles += parseFloat(r[5]) || 0;
      totalNY += parseFloat(r[6]) || 0;
      totalTolls += parseFloat(r[7]) || 0;
    });

    // Delivery type breakdown
    const typeBreakdown = {};
    inv.tmstRows.forEach(r => {
      const t = String(r[9] || 'UNKNOWN').trim().toUpperCase();
      typeBreakdown[t] = (typeBreakdown[t] || 0) + 1;
    });

    // Additional costs summary
    const additionalCosts = inv.costRows.map(r => ({
      description: String(r[2] || ''),
      amount: parseFloat(r[3]) || 0,
      source: String(r[4] || '')
    }));
    const additionalCostsTotal = additionalCosts.reduce((s, c) => s + c.amount, 0);

    // Discrepancy summary
    const discrepancies = inv.discRows.map(r => ({
      date: String(r[1] || ''),
      tu: String(r[2] || ''),
      expected: String(r[3] || ''),
      actual: String(r[4] || ''),
      issue: String(r[5] || '')
    }));

    results.push({
      carrier: inv.carrier,
      rdc: inv.rdc,
      totalTUs,
      matchedTUs,
      unmatchedTUs,
      matchRate: totalTUs > 0 ? Math.round((matchedTUs / totalTUs) * 100) : 0,
      shiftMatchRate: totalTUs > 0 ? Math.round((shiftMatched / totalTUs) * 100) : 0,
      storeMatchRate: totalTUs > 0 ? Math.round((storeMatched / totalTUs) * 100) : 0,
      tourMatchRate: totalTUs > 0 ? Math.round((tourMatched / totalTUs) * 100) : 0,
      totalCost: Math.round(totalCost * 100) / 100,
      totalMiles: Math.round(totalMiles * 100) / 100,
      totalNY: Math.round(totalNY * 100) / 100,
      totalTolls: Math.round(totalTolls * 100) / 100,
      typeBreakdown,
      additionalCosts,
      additionalCostsTotal: Math.round(additionalCostsTotal * 100) / 100,
      discrepancies,
      discrepancyCount: discrepancies.length,
      items: inv.tmstRows.map(r => ({
        date: String(r[1] || ''),
        tu: String(r[3] || ''),
        shift: String(r[2] || ''),
        store: String(r[4] || ''),
        miles: parseFloat(r[5]) || 0,
        nyPay: parseFloat(r[6]) || 0,
        tolls: parseFloat(r[7]) || 0,
        totalCost: parseFloat(r[8]) || 0,
        deliveryType: String(r[9] || ''),
        tuMatch: String(r[10] || ''),
        shiftMatch: String(r[11] || ''),
        storeMatch: String(r[12] || ''),
        tourMatch: String(r[13] || '')
      }))
    });
  }

  return results;
}

// --- BACKHAUL CREDIT SUPPORT ---
// The GL Config now accepts 'Backhaul Credit' as a Rule Type.
// During GL allocation, backhaul credits are subtracted from the base bucket
// and tracked separately for the coded invoice stamp.

function applyBackhaulCredits_(glTotals, fileCostSummary, glConfig, costCenter) {
  if (!glConfig || !glConfig.backhaulCredits) return 0;
  let totalCredits = 0;
  const defaultPrefix = glConfig.defaultCategory || 'XXXXXX, XXXXXXXX';
  const defaultGlString = `${defaultPrefix}, ${costCenter}`;

  fileCostSummary.forEach(item => {
    const descLower = item.desc.toLowerCase();
    for (let credit of glConfig.backhaulCredits) {
      if (descLower.includes(credit.keyword)) {
        const creditGl = `${credit.glPrefix}, ${costCenter}`;
        if (!glTotals[creditGl]) glTotals[creditGl] = 0;
        glTotals[creditGl] -= Math.abs(item.amount);
        totalCredits += Math.abs(item.amount);
        break;
      }
    }
  });
  return totalCredits;
}

// --- OUTPUT ROUTING SUPPORT ---
// Reads the 'Output Routing' config tab to determine if certain data types
// should be written to a different spreadsheet or sheet name.

function getOutputRouting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Output Routing');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const routing = {};
  for (let i = 1; i < data.length; i++) {
    const dataType = String(data[i][0]).trim();
    const targetSheet = String(data[i][1]).trim();
    const targetSpreadsheetId = String(data[i][2]).trim();
    const enabled = String(data[i][3]).trim().toUpperCase();
    if (dataType && enabled !== 'FALSE' && enabled !== 'NO') {
      routing[dataType] = { targetSheet, targetSpreadsheetId };
    }
  }
  return routing;
}
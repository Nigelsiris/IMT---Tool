/**
 * IMT AUTOMATION LAYER
 *
 * Adds hands-off operation on top of the invoice processing engine in code.js:
 *  - Gmail ingestion: carrier invoice emails are picked up automatically, attachments are
 *    dropped into the matching carrier root folder and linked back to the source message.
 *  - Invoice Register: one row per invoice (per carrier / RDC / week) with links to the
 *    Drive file, the coded PDF and the original email thread.
 *  - Week numbering: the calendar week of the invoice is prefixed onto the file name.
 *  - Discrepancy replies: discrepancy notices are sent as a reply inside the carrier's own
 *    email thread (or drafted there when AUTO_SEND_DISCREPANCY_EMAILS = NO).
 *  - Multi-invoice workbooks: a single attachment containing several invoices is split so
 *    each invoice is tracked, coded and (if needed) queried separately.
 *  - Time-driven trigger: autoProcessCycle() runs on a schedule with a lock, a time budget
 *    and an activity log so nothing is processed twice and nothing is lost to the 6-minute
 *    Apps Script execution limit.
 */

// ─────────────────────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────────────────────

var IMT_REGISTER_SHEET = 'Invoice Register';
var IMT_AUTOMATION_LOG_SHEET = 'Automation Log';
var IMT_EMAIL_INGEST_SHEET = 'Email Ingest Config';
var IMT_TRIGGER_HANDLER = 'autoProcessCycle';
var IMT_PROP_LAST_RUN = 'IMT_AUTO_LAST_RUN';
var IMT_PROP_PROCESSED_MSGS = 'IMT_PROCESSED_MESSAGE_IDS';
var IMT_PROP_TRIGGER_MINUTES = 'IMT_TRIGGER_MINUTES';

var IMT_REGISTER_HEADERS = [
  'Registered At', 'Carrier', 'RDC', 'Week', 'Invoice Number', 'File Name', 'Total Amount',
  'TU Count', 'Discrepancies', 'Status', 'Source', 'Email From', 'Email Subject', 'Email Date',
  'Gmail Message ID', 'Gmail Thread ID', 'File URL', 'PDF URL', 'Discrepancy Email',
  'Discrepancy Email At', 'Notes'
];
var IMT_REG = {
  registeredAt: 0, carrier: 1, rdc: 2, week: 3, invoiceNumber: 4, fileName: 5, totalAmount: 6,
  tuCount: 7, discrepancies: 8, status: 9, source: 10, emailFrom: 11, emailSubject: 12,
  emailDate: 13, messageId: 14, threadId: 15, fileUrl: 16, pdfUrl: 17, discEmail: 18,
  discEmailAt: 19, notes: 20
};

var IMT_AUTOMATION_DEFAULTS = [
  ['AUTO_PROCESS_ENABLED', 'YES', 'Master switch for the scheduled automation cycle (YES/NO)'],
  ['AUTO_POLL_MINUTES', '15', 'How often the automation trigger runs (1, 5, 10, 15, 30 or a multiple of 60)'],
  ['AUTO_FINALIZE_MODE', 'AUTO', 'AUTO = stamp GL codes and export PDFs automatically. REVIEW = queue invoices in Invoice Finalization for manual review'],
  ['AUTO_RUN_TIME_BUDGET_SEC', '300', 'Stop picking up new files after this many seconds so a run always finishes cleanly (max 330)'],
  ['GMAIL_SEARCH_QUERY', 'has:attachment newer_than:30d -in:trash -in:spam', 'Gmail search used to find carrier invoice emails. Threads are re-checked for new messages (e.g. a corrected invoice replied into a discrepancy thread); individual messages are never ingested twice'],
  ['GMAIL_PROCESSED_LABEL', 'IMT/Processed', 'Label applied to email threads whose attachments were ingested'],
  ['GMAIL_REVIEW_LABEL', 'IMT/Needs Review', 'Label applied to email threads that could not be matched to a carrier or failed processing'],
  ['GMAIL_MAX_THREADS_PER_RUN', '50', 'Maximum email threads examined per automation cycle (most recently active first)'],
  ['AUTO_SEND_DISCREPANCY_EMAILS', 'YES', 'YES = reply to the carrier automatically. NO = create a draft reply in the thread for you to send'],
  ['DISCREPANCY_EMAIL_REPLY_ALL', 'NO', 'YES = reply to everyone on the original email, NO = reply only to the sender'],
  ['DISCREPANCY_EMAIL_INCLUDE_WARNINGS', 'NO', 'YES = also email non-blocking warnings (Amount Mismatch etc.) to the carrier. NO = only blocking discrepancies'],
  ['DISCREPANCY_EMAIL_CC', '', 'Optional comma-separated CC list added to every discrepancy email'],
  ['DISCREPANCY_EMAIL_BCC', '', 'Optional comma-separated BCC list added to every discrepancy email'],
  ['WEEK_NUMBER_SCHEME', 'ISO', 'ISO, US_SUNDAY or US_MONDAY week numbering for file names and the register'],
  ['FILENAME_WEEK_FORMAT', 'CW{WW} - {FILENAME}', 'File name pattern. Tokens: {WW} {W} {YYYY} {YY} {CARRIER} {RDC} {INVOICE} {FILENAME}'],
  ['SPLIT_MULTI_INVOICE_WORKBOOKS', 'YES', 'YES = split a workbook containing several invoices into one file per invoice'],
  ['UNKNOWN_RDC_FOLDER_NAME', 'Needs Review', 'Sub-folder (under the carrier root) for files whose RDC could not be detected'],
  ['DUPLICATE_FOLDER_NAME', 'Duplicates', 'Sub-folder (under the carrier root) for invoices already present in the Invoice Register'],
  ['AUTO_ARCHIVE_ON_MANUAL_RUN', 'YES', 'YES = manual runs archive and clear the tracker sheets first (legacy behaviour). Automation always appends']
];

// ─────────────────────────────────────────────────────────────────────────────
// RUN CONTEXT  (shared state for a single execution of the processing engine)
// ─────────────────────────────────────────────────────────────────────────────

var _runCtx = null;

function beginRunContext_(opts) {
  opts = opts || {};
  const config = opts.config || {};
  const budgetSec = Math.min(330, Math.max(30, parseInt(String(config.AUTO_RUN_TIME_BUDGET_SEC || '300'), 10) || 300));
  _runCtx = {
    mode: opts.mode || 'manual',            // 'manual' | 'auto'
    startedAt: Date.now(),
    deadline: Date.now() + budgetSec * 1000,
    config: config,
    outbox: [],                             // discrepancy emails to send at the end of the run
    registerRows: [],                       // Invoice Register rows to append at the end of the run
    stats: { files: 0, processed: 0, coded: 0, pending: 0, discrepancy: 0, needsReview: 0, duplicates: 0, errors: 0, remaining: 0, emailsSent: 0, emailsDrafted: 0 },
    registerIndex: null,                    // lazily built map of carrier|invoice -> status
    threadsTouched: {}                      // threadId -> { ok: bool, review: bool }
  };
  return _runCtx;
}

function endRunContext_() {
  const ctx = _runCtx;
  _runCtx = null;
  return ctx;
}

function runContext_() {
  if (!_runCtx) beginRunContext_({ config: safeGetConfig_() });
  return _runCtx;
}

function runBudgetExceeded_() {
  return !!(_runCtx && Date.now() > _runCtx.deadline);
}

function safeGetConfig_() {
  try { return getConfig(); } catch (e) { return {}; }
}

function cfgYes_(config, key, defaultYes) {
  return configYesNo_((config || {})[key], defaultYes);
}

// ─────────────────────────────────────────────────────────────────────────────
// SHEET SCHEMA
// ─────────────────────────────────────────────────────────────────────────────

function ensureAutomationSchema_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  const configSheet = ss.getSheetByName('System Config');
  if (configSheet) {
    const existing = {};
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) existing[String(data[i][0] || '').trim().toUpperCase()] = true;
    const missing = IMT_AUTOMATION_DEFAULTS.filter(row => !existing[row[0]]);
    if (missing.length) configSheet.getRange(configSheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }

  ensureOperationSheet_(ss, IMT_REGISTER_SHEET, IMT_REGISTER_HEADERS, '#c9daf8');
  const reg = ss.getSheetByName(IMT_REGISTER_SHEET);
  if (reg && reg.getLastColumn() < IMT_REGISTER_HEADERS.length) {
    reg.getRange(1, 1, 1, IMT_REGISTER_HEADERS.length).setValues([IMT_REGISTER_HEADERS]).setFontWeight('bold').setBackground('#c9daf8');
  }
  if (reg && reg.getFrozenRows() === 0) reg.setFrozenRows(1);

  ensureOperationSheet_(ss, IMT_AUTOMATION_LOG_SHEET, ['Timestamp', 'Level', 'Event', 'Detail'], '#efefef');

  let ingest = ss.getSheetByName(IMT_EMAIL_INGEST_SHEET);
  if (!ingest) {
    ingest = ss.insertSheet(IMT_EMAIL_INGEST_SHEET);
    ingest.getRange(1, 1, 1, 6).setValues([[
      'Carrier', 'Sender Match (emails or domains, comma separated)', 'Subject Keywords (optional, comma separated)',
      'Attachment Keywords (optional, comma separated)', 'Discrepancy Reply-To Override (optional)', 'Enabled'
    ]]).setFontWeight('bold').setBackground('#fde9d9');
    const carriers = [];
    try {
      const cfg = getConfig();
      Object.keys(cfg).forEach(k => { if (k.endsWith('_ROOT_FOLDER')) carriers.push(k.replace('_ROOT_FOLDER', '')); });
    } catch (e) {}
    const rows = (carriers.length ? carriers : ['CRE', 'HB', 'SCH', 'WERNER']).map(c => [c, '', '', c.toLowerCase(), '', 'YES']);
    ingest.getRange(2, 1, rows.length, 6).setValues(rows);
    ingest.setColumnWidth(1, 110); ingest.setColumnWidth(2, 360); ingest.setColumnWidth(3, 260);
    ingest.setColumnWidth(4, 260); ingest.setColumnWidth(5, 260); ingest.setColumnWidth(6, 80);
  }
}

function getEmailIngestConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IMT_EMAIL_INGEST_SHEET);
  const rules = [];
  if (!sheet) return rules;
  const data = sheet.getDataRange().getValues();
  const splitList = v => String(v || '').toLowerCase().split(',').map(s => s.trim()).filter(Boolean);
  for (let i = 1; i < data.length; i++) {
    const carrier = String(data[i][0] || '').trim().toUpperCase();
    if (!carrier) continue;
    if (!configYesNo_(data[i][5], true)) continue;
    rules.push({
      carrier: carrier,
      senders: splitList(data[i][1]),
      subjectKeywords: splitList(data[i][2]),
      attachmentKeywords: splitList(data[i][3]),
      replyTo: String(data[i][4] || '').trim()
    });
  }
  return rules;
}

// ─────────────────────────────────────────────────────────────────────────────
// LOGGING
// ─────────────────────────────────────────────────────────────────────────────

function logAutomation_(event, detail, level) {
  level = String(level || 'info').toLowerCase();
  const line = `[AUTO][${level.toUpperCase()}] ${event}: ${detail || ''}`;
  Logger.log(line);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(IMT_AUTOMATION_LOG_SHEET);
    if (!sheet) { ensureOperationSheet_(ss, IMT_AUTOMATION_LOG_SHEET, ['Timestamp', 'Level', 'Event', 'Detail'], '#efefef'); sheet = ss.getSheetByName(IMT_AUTOMATION_LOG_SHEET); }
    sheet.appendRow([new Date(), level, String(event || ''), String(detail || '').slice(0, 2000)]);
    const maxRows = 600;
    const last = sheet.getLastRow();
    if (last > maxRows + 100) sheet.deleteRows(2, last - maxRows);
  } catch (e) {
    Logger.log(`[AUTO] Could not write to ${IMT_AUTOMATION_LOG_SHEET}: ${e.message}`);
  }
}

function getAutomationLog(limit) {
  limit = Math.max(1, Math.min(300, parseInt(limit, 10) || 60));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IMT_AUTOMATION_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const last = sheet.getLastRow();
  const start = Math.max(2, last - limit + 1);
  const rows = sheet.getRange(start, 1, last - start + 1, 4).getValues();
  const tz = Session.getScriptTimeZone();
  return rows.reverse().map(r => ({
    timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'MM/dd HH:mm:ss') : String(r[0] || ''),
    level: String(r[1] || 'info'),
    event: String(r[2] || ''),
    detail: String(r[3] || '')
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// WEEK NUMBERS & FILE NAMES
// ─────────────────────────────────────────────────────────────────────────────

function coerceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'number' && value > 20000 && value < 80000) {
    // Excel serial date
    return new Date(Math.round((value - 25569) * 86400000));
  }
  const s = String(value || '').trim();
  if (!s) return null;
  // ISO first (yyyy-mm-dd, optionally with a time part) so it is never read as m/d/y.
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    const d = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
    if (!isNaN(d.getTime())) return d;
  }
  const m = s.match(/(?:^|[^\d])(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?!\d)/);
  if (m) {
    const year = parseInt(m[3].length === 2 ? '20' + m[3] : m[3], 10);
    const d = new Date(year, parseInt(m[1], 10) - 1, parseInt(m[2], 10));
    if (!isNaN(d.getTime())) return d;
  }
  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function weekNumberForScheme_(date, scheme) {
  const s = String(scheme || 'ISO').toUpperCase();
  if (s === 'US_SUNDAY') return getUsWeekNumber_(date, false);
  if (s === 'US_MONDAY') return getUsWeekNumber_(date, true);
  return getIsoWeekNumber_(date);
}

function isoWeekYear_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  return d.getUTCFullYear();
}

/**
 * Determine the invoice week. Weekly invoices normally cover one calendar week, so the most
 * frequent week among the line-item dates wins; ties go to the latest week. Falls back to the
 * carrier's stated period end, then the email date, then today.
 */
function computeInvoiceWeek_(lineDates, fallbackDates, config) {
  const scheme = String((config || {}).WEEK_NUMBER_SCHEME || 'ISO').toUpperCase();
  const counts = {};
  let best = null;
  (lineDates || []).forEach(v => {
    const d = coerceDate_(v);
    if (!d || d.getFullYear() < 2015 || d.getFullYear() > 2100) return;
    const key = isoWeekYear_(d) + '-' + weekNumberForScheme_(d, scheme);
    if (!counts[key]) counts[key] = { n: 0, date: d };
    counts[key].n++;
    if (d > counts[key].date) counts[key].date = d;
  });
  Object.keys(counts).forEach(k => {
    if (!best || counts[k].n > best.n || (counts[k].n === best.n && counts[k].date > best.date)) best = counts[k];
  });
  let source = 'line-items';
  let date = best ? best.date : null;
  if (!date) {
    for (let i = 0; i < (fallbackDates || []).length && !date; i++) {
      const d = coerceDate_(fallbackDates[i]);
      if (d) { date = d; source = 'fallback'; }
    }
  }
  if (!date) { date = new Date(); source = 'today'; }
  const week = weekNumberForScheme_(date, scheme);
  const year = scheme === 'ISO' ? isoWeekYear_(date) : date.getFullYear();
  return { week: week, year: year, label: 'CW' + (week < 10 ? '0' + week : week), date: date, source: source, scheme: scheme };
}

function stripSpreadsheetExtension_(name) {
  return String(name || '').replace(/\.(xlsx|xlsm|xls|csv)$/i, '');
}

function fileNameHasWeekPrefix_(name) {
  return /^\s*(CW|WK|WEEK)\s*\d{1,2}\b/i.test(String(name || ''));
}

/**
 * Build the display name (without extension) for an invoice file using FILENAME_WEEK_FORMAT.
 * Existing week prefixes are preserved so a file is never renamed twice.
 */
function buildWeekFileName_(originalName, weekInfo, extras, config) {
  const ext = (String(originalName || '').match(/\.(xlsx|xlsm|xls|csv)$/i) || [''])[0];
  const base = stripSpreadsheetExtension_(originalName);
  if (!weekInfo || fileNameHasWeekPrefix_(base)) return { base: base, full: base + ext, ext: ext, renamed: false };
  const fmt = String((config || {}).FILENAME_WEEK_FORMAT || 'CW{WW} - {FILENAME}');
  const ww = weekInfo.week < 10 ? '0' + weekInfo.week : String(weekInfo.week);
  extras = extras || {};
  let out = fmt
    .replace(/\{WW\}/g, ww)
    .replace(/\{W\}/g, String(weekInfo.week))
    .replace(/\{YYYY\}/g, String(weekInfo.year))
    .replace(/\{YY\}/g, String(weekInfo.year).slice(-2))
    .replace(/\{CARRIER\}/g, String(extras.carrier || ''))
    .replace(/\{RDC\}/g, String(extras.rdc || ''))
    .replace(/\{INVOICE\}/g, String(extras.invoiceNumber || ''))
    .replace(/\{FILENAME\}/g, base)
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s{2,}/g, ' ')
    .trim();
  if (!out) out = base;
  return { base: out, full: out + ext, ext: ext, renamed: out !== base };
}

// ─────────────────────────────────────────────────────────────────────────────
// FILE ⇄ EMAIL METADATA (stored in the Drive file description)
// ─────────────────────────────────────────────────────────────────────────────

function getFileEmailMeta_(file) {
  try {
    const desc = String(file.getDescription() || '').trim();
    if (!desc || desc.charAt(0) !== '{') return null;
    const meta = JSON.parse(desc);
    return (meta && meta.imt) ? meta : null;
  } catch (e) { return null; }
}

function setFileEmailMeta_(file, meta) {
  try { file.setDescription(JSON.stringify(Object.assign({ imt: 1 }, meta || {}))); } catch (e) {}
}

// ─────────────────────────────────────────────────────────────────────────────
// GMAIL INGESTION
// ─────────────────────────────────────────────────────────────────────────────

function getOrCreateLabel_(name) {
  if (!name) return null;
  try {
    return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
  } catch (e) {
    Logger.log(`[AUTO] Could not get/create label ${name}: ${e.message}`);
    return null;
  }
}

function parseEmailAddress_(fromHeader) {
  const s = String(fromHeader || '');
  const m = s.match(/<([^>]+)>/);
  return (m ? m[1] : s).trim().toLowerCase();
}

function getProcessedMessageIds_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(IMT_PROP_PROCESSED_MSGS);
    return raw ? JSON.parse(raw) : [];
  } catch (e) { return []; }
}

function saveProcessedMessageIds_(ids) {
  try {
    const trimmed = ids.slice(-3000);
    PropertiesService.getScriptProperties().setProperty(IMT_PROP_PROCESSED_MSGS, JSON.stringify(trimmed));
  } catch (e) {
    Logger.log(`[AUTO] Could not persist processed message ids: ${e.message}`);
  }
}

function isSpreadsheetAttachment_(name, contentType) {
  const n = String(name || '').toLowerCase();
  const ct = String(contentType || '').toLowerCase();
  return /\.(xlsx|xlsm|xls|csv)$/.test(n) ||
    ct.indexOf('spreadsheetml') !== -1 || ct === 'application/vnd.ms-excel' || ct === 'text/csv';
}

function isZipAttachment_(name, contentType) {
  return /\.zip$/i.test(String(name || '')) || String(contentType || '').toLowerCase().indexOf('zip') !== -1;
}

/**
 * Match a message to a carrier using the Email Ingest Config, then the carrier names found in
 * the subject / attachment names as a fallback.
 */
function matchCarrierForMessage_(message, attachmentNames, ingestRules, knownCarriers) {
  const from = parseEmailAddress_(message.getFrom());
  const subject = String(message.getSubject() || '').toLowerCase();
  const attachText = (attachmentNames || []).join(' ').toLowerCase();

  // 1. Sender based rules (strongest signal)
  for (const rule of ingestRules) {
    if (!rule.senders.length) continue;
    const senderHit = rule.senders.some(s => s.indexOf('@') === 0 ? from.endsWith(s) : (from === s || from.endsWith('@' + s) || from.endsWith('.' + s)));
    if (!senderHit) continue;
    if (rule.subjectKeywords.length && !rule.subjectKeywords.some(k => subject.indexOf(k) !== -1)) continue;
    return { carrier: rule.carrier, via: 'sender', rule: rule };
  }
  // 2. Attachment name keyword rules (file names are usually carrier specific)
  for (const rule of ingestRules) {
    if (rule.attachmentKeywords.length && rule.attachmentKeywords.some(k => attachText.indexOf(k) !== -1)) return { carrier: rule.carrier, via: 'attachment', rule: rule };
  }
  // 3. Bare carrier name anywhere in subject or attachment names
  const nameCandidates = knownCarriers.slice();
  ingestRules.forEach(r => { if (nameCandidates.indexOf(r.carrier) === -1) nameCandidates.push(r.carrier); });
  for (const c of nameCandidates) {
    const re = new RegExp('\\b' + c.toLowerCase().replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
    if (re.test(subject) || re.test(attachText)) {
      const rule = ingestRules.find(r => r.carrier === c) || null;
      return { carrier: c, via: 'name', rule: rule };
    }
  }
  // 4. Generic subject keyword rules (weakest signal)
  for (const rule of ingestRules) {
    if (rule.subjectKeywords.length && rule.subjectKeywords.some(k => subject.indexOf(k) !== -1)) return { carrier: rule.carrier, via: 'subject', rule: rule };
  }
  return null;
}

/**
 * Scan Gmail for carrier invoices and drop the attachments into the carrier root folders.
 * Returns a summary object. Safe to call repeatedly: processed messages are remembered.
 */
function ingestInvoiceEmails_(config) {
  config = config || safeGetConfig_();
  const summary = { threads: 0, messages: 0, files: 0, unmatched: 0, errors: 0, details: [] };

  const processedLabelName = String(config.GMAIL_PROCESSED_LABEL || 'IMT/Processed').trim();
  const reviewLabelName = String(config.GMAIL_REVIEW_LABEL || 'IMT/Needs Review').trim();
  const processedLabel = getOrCreateLabel_(processedLabelName);
  const reviewLabel = getOrCreateLabel_(reviewLabelName);

  // Processed threads stay in the search on purpose: a carrier may reply to the discrepancy
  // notice with a corrected invoice. Individual messages are de-duplicated by id below.
  const query = String(config.GMAIL_SEARCH_QUERY || 'has:attachment newer_than:30d -in:trash -in:spam').trim();
  const maxThreads = Math.max(1, Math.min(200, parseInt(String(config.GMAIL_MAX_THREADS_PER_RUN || '50'), 10) || 50));

  const ingestRules = getEmailIngestConfig_();
  const knownCarriers = [];
  Object.keys(config).forEach(k => { if (k.endsWith('_ROOT_FOLDER') && config[k]) knownCarriers.push(k.replace('_ROOT_FOLDER', '')); });
  const processedIds = getProcessedMessageIds_();
  const processedSet = {};
  processedIds.forEach(id => processedSet[id] = true);

  let threads = [];
  try {
    threads = GmailApp.search(query, 0, maxThreads);
  } catch (e) {
    logAutomation_('Gmail search failed', `${query} :: ${e.message}`, 'error');
    summary.errors++;
    return summary;
  }
  summary.threads = threads.length;
  if (!threads.length) return summary;

  const folderCache = {};
  const getRootFolder = carrier => {
    if (folderCache[carrier] !== undefined) return folderCache[carrier];
    try { folderCache[carrier] = DriveApp.getFolderById(String(config[carrier + '_ROOT_FOLDER'] || '').trim()); }
    catch (e) { folderCache[carrier] = null; }
    return folderCache[carrier];
  };

  threads.forEach(thread => {
    if (runBudgetExceeded_()) return;
    let threadHadFile = false;
    let threadUnmatched = false;
    let threadError = false;
    let messages = [];
    try { messages = thread.getMessages(); } catch (e) { summary.errors++; return; }

    messages.forEach(message => {
      const msgId = message.getId();
      if (processedSet[msgId]) return;
      let attachments = [];
      try { attachments = message.getAttachments({ includeInlineImages: false, includeAttachments: true }); } catch (e) { attachments = []; }

      // Expand ZIPs and keep only spreadsheet-like blobs
      const blobs = [];
      attachments.forEach(att => {
        const name = att.getName();
        const ct = att.getContentType();
        if (isSpreadsheetAttachment_(name, ct)) { blobs.push(att.copyBlob().setName(name)); return; }
        if (isZipAttachment_(name, ct)) {
          try {
            Utilities.unzip(att.copyBlob()).forEach(inner => {
              if (isSpreadsheetAttachment_(inner.getName(), inner.getContentType())) blobs.push(inner);
            });
          } catch (e) { Logger.log(`[AUTO] Could not unzip ${name}: ${e.message}`); }
        }
      });
      if (!blobs.length) { processedSet[msgId] = true; processedIds.push(msgId); return; }

      summary.messages++;
      const match = matchCarrierForMessage_(message, blobs.map(b => b.getName()), ingestRules, knownCarriers);
      if (!match) {
        threadUnmatched = true;
        summary.unmatched++;
        summary.details.push({ subject: message.getSubject(), from: message.getFrom(), status: 'UNMATCHED' });
        logAutomation_('Email not matched to a carrier', `${message.getFrom()} :: ${message.getSubject()} (${blobs.length} attachment(s))`, 'warning');
        processedSet[msgId] = true; processedIds.push(msgId);
        return;
      }
      const folder = getRootFolder(match.carrier);
      if (!folder) {
        threadError = true;
        summary.errors++;
        logAutomation_('Carrier root folder unavailable', `${match.carrier} for ${message.getSubject()}`, 'error');
        return;
      }

      const meta = {
        source: 'EMAIL',
        carrier: match.carrier,
        messageId: msgId,
        threadId: thread.getId(),
        from: message.getFrom(),
        subject: message.getSubject(),
        date: message.getDate() ? message.getDate().toISOString() : '',
        attachmentCount: blobs.length,
        matchedVia: match.via,
        replyTo: match.rule && match.rule.replyTo ? match.rule.replyTo : ''
      };

      blobs.forEach((blob, idx) => {
        try {
          const file = folder.createFile(blob);
          setFileEmailMeta_(file, Object.assign({ attachmentIndex: idx + 1 }, meta));
          summary.files++;
          threadHadFile = true;
          summary.details.push({ subject: message.getSubject(), from: message.getFrom(), carrier: match.carrier, file: file.getName(), status: 'SAVED' });
        } catch (e) {
          threadError = true;
          summary.errors++;
          logAutomation_('Attachment save failed', `${blob.getName()} :: ${e.message}`, 'error');
        }
      });
      logAutomation_('Email ingested', `${match.carrier} (${match.via}) :: ${message.getSubject()} :: ${blobs.length} file(s)`, 'info');
      processedSet[msgId] = true; processedIds.push(msgId);
    });

    try {
      if (threadHadFile && processedLabel) thread.addLabel(processedLabel);
      if ((threadUnmatched || threadError) && !threadHadFile && reviewLabel) thread.addLabel(reviewLabel);
    } catch (e) { Logger.log(`[AUTO] Label update failed: ${e.message}`); }
  });

  saveProcessedMessageIds_(processedIds);
  return summary;
}

/**
 * Mark the source email thread for manual attention (unknown RDC, processing error).
 */
function flagEmailThreadForReview_(emailMeta) {
  if (!emailMeta || !emailMeta.threadId) return;
  try {
    const config = safeGetConfig_();
    const label = getOrCreateLabel_(String(config.GMAIL_REVIEW_LABEL || 'IMT/Needs Review').trim());
    const thread = GmailApp.getThreadById(emailMeta.threadId);
    if (thread && label) thread.addLabel(label);
  } catch (e) {
    Logger.log(`[AUTO] Could not label thread for review: ${e.message}`);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// MULTI-INVOICE WORKBOOK SPLITTING
// ─────────────────────────────────────────────────────────────────────────────

function sheetLooksLikeInvoiceHeader_(sheet) {
  const n = sheet.getName().toLowerCase();
  return n.indexOf('invoice') !== -1 || n.indexOf('summary') !== -1;
}

function extractInvoiceNumberFromSheet_(sheet, extractionConfig) {
  extractionConfig = extractionConfig || {};
  let regex;
  try { regex = new RegExp(extractionConfig.regexText || 'Invoice\\s*(?:#|ID|Number)\\s*:?\\s*([A-Za-z0-9_-]+)', 'i'); }
  catch (e) { regex = /Invoice\s*(?:#|ID|Number)\s*:?\s*([A-Za-z0-9_-]+)/i; }
  const data = sheet.getDataRange().getValues();
  for (let r = 0; r < Math.min(data.length, 40); r++) {
    for (let c = 0; c < Math.min(data[r].length, 12); c++) {
      const v = String(data[r][c] || '').trim();
      if (!v) continue;
      const m = v.match(regex);
      if (m && m[1]) return m[1].trim();
      if (/^Invoice\s*(?:#|ID|Number)\s*:?$/i.test(v)) {
        for (let nc = c + 1; nc < Math.min(data[r].length, c + 4); nc++) {
          const nv = String(data[r][nc] || '').trim();
          if (nv) return nv;
        }
      }
    }
  }
  return '';
}

/**
 * Detect a workbook that bundles several invoices (several invoice/summary tabs, each with a
 * distinct invoice number). Returns [] when the workbook is a single invoice; otherwise an array
 * of { invoiceNumber, sheetNames } groups in workbook order.
 */
function detectInvoiceGroups_(invoiceSS, extractionConfig) {
  const sheets = invoiceSS.getSheets();
  if (sheets.length < 2) return [];
  const groups = [];
  let current = null;
  sheets.forEach(sheet => {
    if (sheetLooksLikeInvoiceHeader_(sheet)) {
      const num = extractInvoiceNumberFromSheet_(sheet, extractionConfig);
      if (num && (!current || current.invoiceNumber !== num)) {
        current = { invoiceNumber: num, sheetNames: [] };
        groups.push(current);
      }
    }
    if (current) current.sheetNames.push(sheet.getName());
    else {
      // Sheets before the first invoice header belong to the first group once it is created.
      if (!groups.length) groups.push({ invoiceNumber: '', sheetNames: [] });
      groups[0].sheetNames.push(sheet.getName());
    }
  });
  const distinct = {};
  groups.forEach(g => { if (g.invoiceNumber) distinct[g.invoiceNumber] = true; });
  if (Object.keys(distinct).length < 2) return [];
  // Merge the leading (header-less) group into the first real group.
  if (groups.length && !groups[0].invoiceNumber && groups.length > 1) {
    groups[1].sheetNames = groups[0].sheetNames.concat(groups[1].sheetNames);
    groups.shift();
  }
  return groups.filter(g => g.invoiceNumber);
}

/**
 * Create one temporary Google Sheet per invoice group. Each new workbook receives the sheets of
 * its group (copied with values and formatting) so the existing carrier parsers work unchanged.
 */
function splitMultiInvoiceWorkbook_(invoiceSS, groups, targetFolder, baseName) {
  const results = [];
  groups.forEach(group => {
    try {
      const name = `[TEMP] ${stripSpreadsheetExtension_(baseName)} [${group.invoiceNumber}]`;
      const newSS = SpreadsheetApp.create(name);
      group.sheetNames.forEach(sheetName => {
        const src = invoiceSS.getSheetByName(sheetName);
        if (src) src.copyTo(newSS).setName(sheetName);
      });
      const defaultSheet = newSS.getSheets()[0];
      if (newSS.getSheets().length > 1 && group.sheetNames.indexOf(defaultSheet.getName()) === -1) newSS.deleteSheet(defaultSheet);
      SpreadsheetApp.flush();
      try { moveFileToFolder(DriveApp.getFileById(newSS.getId()), targetFolder); } catch (e) {}
      results.push({ ss: newSS, sheetId: newSS.getId(), invoiceNumber: group.invoiceNumber, isTemp: true, sheetNames: group.sheetNames });
    } catch (e) {
      Logger.log(`[AUTO] Failed to split invoice group ${group.invoiceNumber}: ${e.message}`);
    }
  });
  return results;
}

// ─────────────────────────────────────────────────────────────────────────────
// INVOICE REGISTER
// ─────────────────────────────────────────────────────────────────────────────

function buildRegisterIndex_() {
  const index = { byKey: {}, byMessage: {} };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IMT_REGISTER_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return index;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, IMT_REGISTER_HEADERS.length).getValues();
  data.forEach((row, i) => {
    const key = registerKey_(row[IMT_REG.carrier], row[IMT_REG.invoiceNumber]);
    const status = String(row[IMT_REG.status] || '').toUpperCase();
    if (key) index.byKey[key] = { rowNumber: i + 2, status: status, fileName: String(row[IMT_REG.fileName] || '') };
  });
  return index;
}

function registerKey_(carrier, invoiceNumber) {
  const c = String(carrier || '').trim().toUpperCase();
  const n = String(invoiceNumber || '').trim().toUpperCase().replace(/\s+/g, '');
  if (!c || !n) return '';
  return c + '|' + n;
}

function isDuplicateInvoice_(carrier, invoiceNumber) {
  const ctx = runContext_();
  if (!ctx.registerIndex) ctx.registerIndex = buildRegisterIndex_();
  const key = registerKey_(carrier, invoiceNumber);
  if (!key) return null;
  // Rows queued in this run count as well (multi-attachment emails with the same invoice twice).
  for (const row of ctx.registerRows) {
    if (registerKey_(row[IMT_REG.carrier], row[IMT_REG.invoiceNumber]) === key) return { status: String(row[IMT_REG.status] || ''), fileName: String(row[IMT_REG.fileName] || '') };
  }
  const hit = ctx.registerIndex.byKey[key];
  if (!hit) return null;
  // Rows that were rejected can be retried.
  if (['NEEDS_REVIEW', 'ERROR', 'DUPLICATE', 'SKIPPED'].indexOf(hit.status) !== -1) return null;
  return hit;
}

/**
 * Queue an Invoice Register row. Rows are written in one batch at the end of the run.
 */
function registerInvoice_(entry) {
  const ctx = runContext_();
  const meta = entry.emailMeta || {};
  const tz = Session.getScriptTimeZone();
  let emailDate = '';
  if (meta.date) { const d = coerceDate_(meta.date); if (d) emailDate = Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm'); }
  const row = [];
  row[IMT_REG.registeredAt] = new Date();
  row[IMT_REG.carrier] = String(entry.carrier || '');
  row[IMT_REG.rdc] = String(entry.rdc || '');
  row[IMT_REG.week] = entry.week ? String(entry.week.label || '') : '';
  row[IMT_REG.invoiceNumber] = String(entry.invoiceNumber || '');
  row[IMT_REG.fileName] = String(entry.fileName || '');
  row[IMT_REG.totalAmount] = (typeof entry.totalAmount === 'number' && !isNaN(entry.totalAmount)) ? Math.round(entry.totalAmount * 100) / 100 : '';
  row[IMT_REG.tuCount] = entry.tuCount || 0;
  row[IMT_REG.discrepancies] = entry.discrepancyCount || 0;
  row[IMT_REG.status] = String(entry.status || 'PROCESSED');
  row[IMT_REG.source] = String(meta.source || entry.source || (ctx.mode === 'auto' ? 'AUTO' : 'MANUAL'));
  row[IMT_REG.emailFrom] = String(meta.from || '');
  row[IMT_REG.emailSubject] = String(meta.subject || '');
  row[IMT_REG.emailDate] = emailDate;
  row[IMT_REG.messageId] = String(meta.messageId || '');
  row[IMT_REG.threadId] = String(meta.threadId || '');
  row[IMT_REG.fileUrl] = String(entry.fileUrl || '');
  row[IMT_REG.pdfUrl] = String(entry.pdfUrl || '');
  row[IMT_REG.discEmail] = String(entry.discEmail || (entry.discrepancyCount > 0 ? 'PENDING' : ''));
  row[IMT_REG.discEmailAt] = '';
  row[IMT_REG.notes] = String(entry.notes || '');
  for (let i = 0; i < IMT_REGISTER_HEADERS.length; i++) if (row[i] === undefined) row[i] = '';
  ctx.registerRows.push(row);
  return row;
}

function flushRegisterRows_() {
  const ctx = _runCtx;
  if (!ctx || !ctx.registerRows.length) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAutomationSchema_(ss);
  const sheet = ss.getSheetByName(IMT_REGISTER_SHEET);
  const rows = ctx.registerRows.splice(0, ctx.registerRows.length);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, IMT_REGISTER_HEADERS.length).setValues(rows);
  ctx.registerIndex = null;
  return rows.length;
}

/**
 * Update the discrepancy-email columns for register rows matching a carrier + invoice number.
 * Works on both queued rows (current run) and rows already in the sheet.
 */
function markRegisterDiscrepancyEmail_(carrier, invoiceNumber, status) {
  const key = registerKey_(carrier, invoiceNumber);
  const ctx = _runCtx;
  let touched = false;
  if (ctx) {
    ctx.registerRows.forEach(row => {
      if (registerKey_(row[IMT_REG.carrier], row[IMT_REG.invoiceNumber]) === key) {
        row[IMT_REG.discEmail] = status; row[IMT_REG.discEmailAt] = new Date(); touched = true;
      }
    });
  }
  if (touched) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IMT_REGISTER_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, IMT_REGISTER_HEADERS.length).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (registerKey_(data[i][IMT_REG.carrier], data[i][IMT_REG.invoiceNumber]) === key) {
      sheet.getRange(i + 2, IMT_REG.discEmail + 1, 1, 2).setValues([[status, new Date()]]);
      return;
    }
  }
}

/**
 * Update register columns for the most recent row with a given file name (used when a pending
 * invoice is finalized from the UI).
 */
function updateRegisterByFileName_(fileName, updates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(IMT_REGISTER_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return false;
    const target = String(fileName || '').trim().toLowerCase();
    const targetBase = stripSpreadsheetExtension_(target);
    const names = sheet.getRange(2, IMT_REG.fileName + 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = names.length - 1; i >= 0; i--) {
      const n = String(names[i][0] || '').trim().toLowerCase();
      if (n === target || stripSpreadsheetExtension_(n) === targetBase) {
        const rowNumber = i + 2;
        if (updates.status !== undefined) sheet.getRange(rowNumber, IMT_REG.status + 1).setValue(updates.status);
        if (updates.pdfUrl) sheet.getRange(rowNumber, IMT_REG.pdfUrl + 1).setValue(updates.pdfUrl);
        if (updates.notes !== undefined) sheet.getRange(rowNumber, IMT_REG.notes + 1).setValue(updates.notes);
        return true;
      }
    }
  } catch (e) {
    Logger.log(`[AUTO] updateRegisterByFileName_ failed: ${e.message}`);
  }
  return false;
}

function getInvoiceRegister(options) {
  options = options || {};
  const limit = Math.max(1, Math.min(2000, parseInt(options.limit, 10) || 500));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IMT_REGISTER_SHEET);
  const tz = Session.getScriptTimeZone();
  const out = { rows: [], summary: { total: 0, thisWeek: 0, awaitingReply: 0, needsReview: 0, coded: 0, pendingReview: 0 } };
  if (!sheet || sheet.getLastRow() < 2) return out;
  const last = sheet.getLastRow();
  const start = Math.max(2, last - limit + 1);
  const data = sheet.getRange(start, 1, last - start + 1, IMT_REGISTER_HEADERS.length).getValues();
  const totalRows = last - 1;
  const currentWeek = computeInvoiceWeek_([new Date()], [], safeGetConfig_()).label;
  const fmtDate = v => v instanceof Date ? Utilities.formatDate(v, tz, 'MM/dd/yyyy HH:mm') : String(v || '');

  out.summary.total = totalRows;
  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    const status = String(r[IMT_REG.status] || '');
    const week = String(r[IMT_REG.week] || '');
    const discEmail = String(r[IMT_REG.discEmail] || '');
    if (week === currentWeek) out.summary.thisWeek++;
    if (status === 'DISCREPANCY' && discEmail !== 'RESOLVED') out.summary.awaitingReply++;
    if (status === 'NEEDS_REVIEW' || status === 'ERROR') out.summary.needsReview++;
    if (status === 'CODED') out.summary.coded++;
    if (status === 'PENDING_REVIEW') out.summary.pendingReview++;
    out.rows.push({
      rowNumber: start + i,
      registeredAt: fmtDate(r[IMT_REG.registeredAt]),
      carrier: String(r[IMT_REG.carrier] || ''),
      rdc: String(r[IMT_REG.rdc] || ''),
      week: week,
      invoiceNumber: String(r[IMT_REG.invoiceNumber] || ''),
      fileName: String(r[IMT_REG.fileName] || ''),
      totalAmount: parseFloat(r[IMT_REG.totalAmount]) || 0,
      tuCount: parseInt(r[IMT_REG.tuCount], 10) || 0,
      discrepancies: parseInt(r[IMT_REG.discrepancies], 10) || 0,
      status: status,
      source: String(r[IMT_REG.source] || ''),
      emailFrom: String(r[IMT_REG.emailFrom] || ''),
      emailSubject: String(r[IMT_REG.emailSubject] || ''),
      emailDate: fmtDate(r[IMT_REG.emailDate]),
      threadId: String(r[IMT_REG.threadId] || ''),
      threadUrl: r[IMT_REG.threadId] ? 'https://mail.google.com/mail/u/0/#all/' + String(r[IMT_REG.threadId]) : '',
      fileUrl: String(r[IMT_REG.fileUrl] || ''),
      pdfUrl: String(r[IMT_REG.pdfUrl] || ''),
      discEmail: discEmail,
      discEmailAt: fmtDate(r[IMT_REG.discEmailAt]),
      notes: String(r[IMT_REG.notes] || '')
    });
  }
  return out;
}

function updateRegisterRowWeb(rowNumber, updates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(IMT_REGISTER_SHEET);
    if (!sheet) throw new Error('Invoice Register sheet not found.');
    rowNumber = parseInt(rowNumber, 10);
    if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) throw new Error('Invalid register row.');
    updates = updates || {};
    if (updates.status !== undefined) sheet.getRange(rowNumber, IMT_REG.status + 1).setValue(String(updates.status));
    if (updates.discEmail !== undefined) sheet.getRange(rowNumber, IMT_REG.discEmail + 1).setValue(String(updates.discEmail));
    if (updates.notes !== undefined) sheet.getRange(rowNumber, IMT_REG.notes + 1).setValue(String(updates.notes));
    return { success: true, message: 'Register row updated.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// DISCREPANCY EMAILS
// ─────────────────────────────────────────────────────────────────────────────

function escapeHtml_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function formatDiscrepancyDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  const d = coerceDate_(v);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy') : String(v || '');
}

function applyEmailTokens_(text, tokens) {
  let out = String(text || '');
  Object.keys(tokens).forEach(k => { out = out.replace(new RegExp('\\{' + k + '\\}', 'g'), String(tokens[k] === undefined ? '' : tokens[k])); });
  return out;
}

/**
 * Build subject / plain / html bodies for one or more invoices that belong to the same carrier
 * (and normally the same email thread). items: [{ fileName, invoiceNumber, week, rdc, discrepancies }]
 */
function buildDiscrepancyEmailContent_(carrierName, items, emailTemplate) {
  emailTemplate = emailTemplate || {};
  const subjTemplate = emailTemplate['Subject'] || 'Discrepancy Notice: Invoice {FileName}';
  const greetTemplate = emailTemplate['Greeting'] || 'Hello {CarrierName} Team,\n\nThe following items are showing as discrepancies in our system. Please advise:';
  const outroTemplate = emailTemplate['Outro'] || 'Thank you.';

  const first = items[0] || {};
  const invoiceList = items.map(i => i.invoiceNumber || stripSpreadsheetExtension_(i.fileName)).filter(Boolean);
  const tokens = {
    FileName: items.length === 1 ? (first.fileName || '') : invoiceList.join(', '),
    CarrierName: carrierName,
    InvoiceNumber: invoiceList.join(', '),
    Week: first.week ? first.week.label : '',
    RDC: items.map(i => i.rdc).filter(Boolean).filter((v, i, a) => a.indexOf(v) === i).join(', ')
  };

  const subject = applyEmailTokens_(subjTemplate, tokens);
  const greeting = applyEmailTokens_(greetTemplate, tokens);
  const outro = applyEmailTokens_(outroTemplate, tokens);

  let plain = greeting + '\n\n';
  let html = '<div style="font-family: Arial, sans-serif; color: #333; font-size: 13px;">';
  html += '<p>' + escapeHtml_(greeting).replace(/\n/g, '<br>') + '</p>';

  items.forEach(item => {
    const title = [item.invoiceNumber ? 'Invoice ' + item.invoiceNumber : '', item.fileName, item.rdc, item.week ? item.week.label : '']
      .filter(Boolean).join(' · ');
    if (items.length > 1 || item.invoiceNumber) {
      plain += title + '\n';
      html += '<p style="margin:16px 0 6px;font-weight:bold;">' + escapeHtml_(title) + '</p>';
    }
    html += '<table style="border-collapse: collapse; width: 100%; max-width: 820px;">'
      + '<thead><tr style="background-color:#f2f2f2;">'
      + '<th style="border:1px solid #000;padding:8px;text-align:left;">Date</th>'
      + '<th style="border:1px solid #000;padding:8px;text-align:left;">TU</th>'
      + '<th style="border:1px solid #000;padding:8px;text-align:left;">Issue</th>'
      + '<th style="border:1px solid #000;padding:8px;text-align:left;">Our Record</th>'
      + '<th style="border:1px solid #000;padding:8px;text-align:left;">Billed</th>'
      + '</tr></thead><tbody>';
    (item.discrepancies || []).forEach(d => {
      const dateStr = formatDiscrepancyDate_(d[1]);
      plain += `• Date: ${dateStr} | TU: ${d[2]} | Issue: ${d[5]} | Our record: ${d[3]} | Billed: ${d[4]}\n`;
      html += '<tr>'
        + '<td style="border:1px solid #000;padding:6px;">' + escapeHtml_(dateStr) + '</td>'
        + '<td style="border:1px solid #000;padding:6px;">' + escapeHtml_(d[2]) + '</td>'
        + '<td style="border:1px solid #000;padding:6px;">' + escapeHtml_(d[5]) + '</td>'
        + '<td style="border:1px solid #000;padding:6px;">' + escapeHtml_(d[3]) + '</td>'
        + '<td style="border:1px solid #000;padding:6px;">' + escapeHtml_(d[4]) + '</td>'
        + '</tr>';
    });
    html += '</tbody></table>';
    plain += '\n';
  });

  plain += outro;
  html += '<p style="margin-top:20px;">' + escapeHtml_(outro).replace(/\n/g, '<br>') + '</p></div>';
  return { subject: subject, plainBody: plain, htmlBody: html };
}

/**
 * Queue a discrepancy notice. Notices are grouped per email thread (or per carrier when the
 * invoice did not arrive by email) and delivered once at the end of the run.
 */
function queueDiscrepancyEmail_(item) {
  const ctx = runContext_();
  ctx.outbox.push(item);
}

function deliverDiscrepancyEmail_(carrierName, items, emailTemplate, config) {
  const content = buildDiscrepancyEmailContent_(carrierName, items, emailTemplate);
  const autoSend = cfgYes_(config, 'AUTO_SEND_DISCREPANCY_EMAILS', true);
  const replyAll = cfgYes_(config, 'DISCREPANCY_EMAIL_REPLY_ALL', false);
  const cc = String(config.DISCREPANCY_EMAIL_CC || '').trim();
  const bcc = String(config.DISCREPANCY_EMAIL_BCC || '').trim();
  const options = { htmlBody: content.htmlBody };
  if (cc) options.cc = cc;
  if (bcc) options.bcc = bcc;

  const meta = (items.find(i => i.emailMeta && i.emailMeta.messageId) || {}).emailMeta || null;
  const replyToOverride = (items.find(i => i.emailMeta && i.emailMeta.replyTo) || {}).emailMeta;
  const overrideAddress = replyToOverride ? replyToOverride.replyTo : lookupCarrierReplyTo_(carrierName);

  // 1. Reply inside the original thread
  if (meta && meta.messageId) {
    try {
      const message = GmailApp.getMessageById(meta.messageId);
      if (message) {
        if (autoSend) {
          if (replyAll) message.replyAll(content.plainBody, options); else message.reply(content.plainBody, options);
          return { mode: 'SENT', target: 'thread', to: message.getFrom() };
        }
        message.createDraftReply(content.plainBody, options);
        return { mode: 'DRAFTED', target: 'thread', to: message.getFrom() };
      }
    } catch (e) {
      logAutomation_('Thread reply failed', `${carrierName} :: ${e.message}. Falling back to a new email.`, 'warning');
    }
  }

  // 2. Configured carrier address (no thread available)
  if (overrideAddress) {
    if (autoSend) {
      GmailApp.sendEmail(overrideAddress, content.subject, content.plainBody, options);
      return { mode: 'SENT', target: 'address', to: overrideAddress };
    }
    GmailApp.createDraft(overrideAddress, content.subject, content.plainBody, options);
    return { mode: 'DRAFTED', target: 'address', to: overrideAddress };
  }

  // 3. Legacy behaviour: an unaddressed draft
  GmailApp.createDraft('', content.subject, content.plainBody, options);
  return { mode: 'DRAFTED', target: 'unaddressed', to: '' };
}

function lookupCarrierReplyTo_(carrierName) {
  const key = String(carrierName || '').trim().toUpperCase();
  const rule = getEmailIngestConfig_().find(r => r.carrier === key);
  return rule && rule.replyTo ? rule.replyTo : '';
}

function flushDiscrepancyOutbox_(emailTemplate, config) {
  const ctx = _runCtx;
  if (!ctx || !ctx.outbox.length) return { sent: 0, drafted: 0 };
  config = config || ctx.config || safeGetConfig_();
  const groups = {};
  ctx.outbox.splice(0, ctx.outbox.length).forEach(item => {
    const threadKey = (item.emailMeta && item.emailMeta.threadId) ? 'T:' + item.emailMeta.threadId : 'C:' + item.carrier + ':' + (item.fileName || '');
    if (!groups[threadKey]) groups[threadKey] = { carrier: item.carrier, items: [] };
    groups[threadKey].items.push(item);
  });
  const result = { sent: 0, drafted: 0 };
  Object.keys(groups).forEach(key => {
    const group = groups[key];
    try {
      const outcome = deliverDiscrepancyEmail_(group.carrier, group.items, emailTemplate, config);
      if (outcome.mode === 'SENT') { result.sent++; ctx.stats.emailsSent++; } else { result.drafted++; ctx.stats.emailsDrafted++; }
      group.items.forEach(item => markRegisterDiscrepancyEmail_(item.carrier, item.invoiceNumber, outcome.mode + (outcome.target === 'thread' ? ' (thread)' : outcome.to ? ' (' + outcome.to + ')' : '')));
      logAutomation_('Discrepancy email ' + outcome.mode.toLowerCase(), `${group.carrier} :: ${group.items.map(i => i.invoiceNumber || i.fileName).join(', ')} -> ${outcome.to || outcome.target}`, 'info');
    } catch (e) {
      group.items.forEach(item => markRegisterDiscrepancyEmail_(item.carrier, item.invoiceNumber, 'FAILED: ' + e.message.slice(0, 80)));
      logAutomation_('Discrepancy email failed', `${group.carrier} :: ${e.message}`, 'error');
    }
  });
  return result;
}

/**
 * Re-send (or draft) the discrepancy notice for a register row using the current Discrepancy
 * Tracker contents. Used by the Invoice Register UI.
 */
function resendDiscrepancyEmailWeb(rowNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reg = ss.getSheetByName(IMT_REGISTER_SHEET);
    if (!reg) throw new Error('Invoice Register sheet not found.');
    rowNumber = parseInt(rowNumber, 10);
    if (!rowNumber || rowNumber < 2 || rowNumber > reg.getLastRow()) throw new Error('Invalid register row.');
    const row = reg.getRange(rowNumber, 1, 1, IMT_REGISTER_HEADERS.length).getValues()[0];
    const carrier = String(row[IMT_REG.carrier] || '');
    const invoiceNumber = String(row[IMT_REG.invoiceNumber] || '');
    const discrepancies = findTrackerDiscrepancies_(carrier, invoiceNumber, String(row[IMT_REG.fileName] || ''));
    if (!discrepancies.length) throw new Error('No discrepancy rows found in the Discrepancy Tracker for this invoice (they may have been archived).');
    const config = getConfig();
    const item = {
      carrier: carrier,
      invoiceNumber: invoiceNumber,
      fileName: String(row[IMT_REG.fileName] || ''),
      rdc: String(row[IMT_REG.rdc] || ''),
      week: row[IMT_REG.week] ? { label: String(row[IMT_REG.week]) } : null,
      discrepancies: discrepancies,
      emailMeta: row[IMT_REG.messageId] ? { messageId: String(row[IMT_REG.messageId]), threadId: String(row[IMT_REG.threadId] || '') } : null
    };
    const outcome = deliverDiscrepancyEmail_(carrier, [item], getEmailTemplate(), config);
    reg.getRange(rowNumber, IMT_REG.discEmail + 1, 1, 2).setValues([[outcome.mode + (outcome.target === 'thread' ? ' (thread)' : outcome.to ? ' (' + outcome.to + ')' : ''), new Date()]]);
    logAutomation_('Discrepancy email re-' + outcome.mode.toLowerCase(), `${carrier} :: ${invoiceNumber} -> ${outcome.to || outcome.target}`, 'info');
    return { success: true, message: `Discrepancy email ${outcome.mode.toLowerCase()} for ${invoiceNumber || item.fileName}.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function findTrackerDiscrepancies_(carrier, invoiceNumber, fileName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Discrepancy Tracker');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(8, sheet.getLastColumn())).getValues();
  const c = String(carrier || '').toUpperCase();
  const n = String(invoiceNumber || '').toUpperCase();
  return data.filter(r => String(r[6] || '').toUpperCase() === c && String(r[7] || '').toUpperCase() === n && String(r[5] || '').trim());
}

// ─────────────────────────────────────────────────────────────────────────────
// TRIGGERS & AUTOMATION CYCLE
// ─────────────────────────────────────────────────────────────────────────────

function getAutomationTriggers_() {
  return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === IMT_TRIGGER_HANDLER);
}

function normalizePollMinutes_(minutes) {
  const n = Math.max(1, parseInt(minutes, 10) || 15);
  if (n >= 60) return { hours: Math.max(1, Math.round(n / 60)), minutes: 0 };
  const allowed = [1, 5, 10, 15, 30];
  let best = allowed[0];
  allowed.forEach(a => { if (Math.abs(a - n) < Math.abs(best - n)) best = a; });
  return { hours: 0, minutes: best };
}

function installAutomationTriggers() {
  const config = safeGetConfig_();
  const schedule = normalizePollMinutes_(config.AUTO_POLL_MINUTES);
  getAutomationTriggers_().forEach(t => ScriptApp.deleteTrigger(t));
  const builder = ScriptApp.newTrigger(IMT_TRIGGER_HANDLER).timeBased();
  if (schedule.hours) builder.everyHours(schedule.hours).create(); else builder.everyMinutes(schedule.minutes).create();
  const effective = schedule.hours ? schedule.hours * 60 : schedule.minutes;
  PropertiesService.getScriptProperties().setProperty(IMT_PROP_TRIGGER_MINUTES, String(effective));
  ensureAutomationSchema_(SpreadsheetApp.getActiveSpreadsheet());
  logAutomation_('Trigger installed', `autoProcessCycle every ${effective} minute(s)`, 'info');
  return { success: true, message: `Automation trigger installed: runs every ${effective} minute(s).`, minutes: effective };
}

function removeAutomationTriggers() {
  const triggers = getAutomationTriggers_();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().deleteProperty(IMT_PROP_TRIGGER_MINUTES);
  logAutomation_('Trigger removed', `${triggers.length} trigger(s) deleted`, 'info');
  return { success: true, message: triggers.length ? 'Automation trigger removed.' : 'No automation trigger was installed.' };
}

function installAutomationTriggersWeb() { try { return installAutomationTriggers(); } catch (e) { return { success: false, message: e.message }; } }
function removeAutomationTriggersWeb() { try { return removeAutomationTriggers(); } catch (e) { return { success: false, message: e.message }; } }

function recordLastRun_(summary) {
  try {
    PropertiesService.getScriptProperties().setProperty(IMT_PROP_LAST_RUN, JSON.stringify(Object.assign({ at: new Date().toISOString() }, summary || {})));
  } catch (e) {}
}

function getLastRun_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(IMT_PROP_LAST_RUN);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

/**
 * Entry point for the time-driven trigger. Also callable from the UI ("Run cycle now").
 */
function autoProcessCycle() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log('[AUTO] Another run is in progress. Skipping this cycle.');
    return 'SKIPPED: another run is in progress.';
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let result = '';
  try {
    ensureAutomationSchema_(ss);
    const config = getConfig();
    if (!cfgYes_(config, 'AUTO_PROCESS_ENABLED', true)) {
      result = 'SKIPPED: AUTO_PROCESS_ENABLED is NO.';
      return result;
    }
    const ctx = beginRunContext_({ mode: 'auto', config: config });
    const ingest = ingestInvoiceEmails_(config);
    const reviewMode = String(config.AUTO_FINALIZE_MODE || 'AUTO').trim().toUpperCase() === 'REVIEW';
    const engine = runMainProcess(true, { mode: 'auto', extractOnly: reviewMode, appendMode: true, keepContext: true });
    const stats = ctx.stats;
    const summary = {
      emailsThreads: ingest.threads, emailsFiles: ingest.files, emailsUnmatched: ingest.unmatched,
      files: stats.files, processed: stats.processed, coded: stats.coded, pending: stats.pending,
      discrepancy: stats.discrepancy, needsReview: stats.needsReview, duplicates: stats.duplicates,
      errors: stats.errors + ingest.errors, remaining: stats.remaining,
      emailsSent: stats.emailsSent, emailsDrafted: stats.emailsDrafted,
      engine: engine, durationSec: Math.round((Date.now() - ctx.startedAt) / 1000)
    };
    endRunContext_();
    recordLastRun_(summary);
    result = `SUCCESS: cycle complete in ${summary.durationSec}s. Emails: ${ingest.files} file(s) from ${ingest.threads} thread(s). Files processed: ${stats.processed}/${stats.files}` +
      (stats.remaining ? ` (${stats.remaining} left for next cycle)` : '') +
      `. Coded: ${stats.coded}, discrepancies: ${stats.discrepancy}, review: ${stats.needsReview}, duplicates: ${stats.duplicates}, emails sent/drafted: ${stats.emailsSent}/${stats.emailsDrafted}.`;
    // Idle cycles (no emails, no files) are recorded in the last-run property only, to keep the log readable.
    const didWork = ingest.files || ingest.unmatched || stats.files || summary.errors;
    if (didWork) logAutomation_('Cycle complete', result, summary.errors ? 'warning' : 'info');
    return result;
  } catch (e) {
    endRunContext_();
    recordLastRun_({ error: e.message });
    logAutomation_('Cycle failed', e.message, 'error');
    return 'ERROR: ' + e.message;
  } finally {
    lock.releaseLock();
  }
}

function runAutomationCycleWeb() { return autoProcessCycle(); }

function ingestEmailsOnlyWeb() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 'SKIPPED: another run is in progress.';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAutomationSchema_(ss);
    const config = getConfig();
    beginRunContext_({ mode: 'auto', config: config });
    const s = ingestInvoiceEmails_(config);
    endRunContext_();
    const msg = `SUCCESS: inbox scan complete. ${s.files} attachment(s) saved from ${s.messages} message(s) in ${s.threads} thread(s). Unmatched: ${s.unmatched}, errors: ${s.errors}.`;
    logAutomation_('Inbox scan', msg, s.errors ? 'warning' : 'info');
    return msg;
  } catch (e) {
    endRunContext_();
    return 'ERROR: ' + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Status payload for the Automation view.
 */
function getAutomationStatus() {
  const config = safeGetConfig_();
  const triggers = getAutomationTriggers_();
  const props = PropertiesService.getScriptProperties();
  const minutes = parseInt(props.getProperty(IMT_PROP_TRIGGER_MINUTES) || '0', 10) || normalizePollMinutes_(config.AUTO_POLL_MINUTES).minutes || 60;
  const lastRun = getLastRun_();
  let nextRunApprox = '';
  if (triggers.length && lastRun && lastRun.at) {
    const next = new Date(new Date(lastRun.at).getTime() + minutes * 60000);
    nextRunApprox = Utilities.formatDate(next, Session.getScriptTimeZone(), 'MM/dd HH:mm');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = {};
  IMT_AUTOMATION_DEFAULTS.forEach(row => { settings[row[0]] = config[row[0]] !== undefined ? String(config[row[0]]) : row[1]; });
  let gmailOk = true; let gmailDetail = '';
  try { gmailDetail = Session.getEffectiveUser().getEmail(); } catch (e) { gmailOk = false; gmailDetail = e.message; }
  return {
    enabled: cfgYes_(config, 'AUTO_PROCESS_ENABLED', true),
    triggerInstalled: triggers.length > 0,
    triggerMinutes: minutes,
    lastRun: lastRun,
    lastRunDisplay: lastRun && lastRun.at ? Utilities.formatDate(new Date(lastRun.at), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm') : '',
    nextRunApprox: nextRunApprox,
    mode: String(config.AUTO_FINALIZE_MODE || 'AUTO').toUpperCase(),
    autoSend: cfgYes_(config, 'AUTO_SEND_DISCREPANCY_EMAILS', true),
    account: gmailDetail,
    gmailOk: gmailOk,
    settings: settings,
    settingDescriptions: IMT_AUTOMATION_DEFAULTS.reduce((m, r) => { m[r[0]] = r[2]; return m; }, {}),
    ingestRules: getEmailIngestConfig_(),
    schemaReady: !!ss.getSheetByName(IMT_REGISTER_SHEET) && !!ss.getSheetByName(IMT_EMAIL_INGEST_SHEET),
    log: getAutomationLog(40)
  };
}

/**
 * Save automation settings (subset of System Config) from the UI in one call.
 */
function saveAutomationSettings(settings) {
  try {
    settings = settings || {};
    const allowed = {};
    IMT_AUTOMATION_DEFAULTS.forEach(r => allowed[r[0]] = r);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAutomationSchema_(ss);
    const sheet = ss.getSheetByName('System Config');
    const data = sheet.getDataRange().getValues();
    const rowByKey = {};
    for (let i = 1; i < data.length; i++) rowByKey[String(data[i][0] || '').trim().toUpperCase()] = i + 1;
    let changed = 0;
    Object.keys(settings).forEach(key => {
      const k = String(key).toUpperCase();
      if (!allowed[k]) return;
      const value = String(settings[key] === undefined || settings[key] === null ? '' : settings[key]).trim();
      if (rowByKey[k]) sheet.getRange(rowByKey[k], 2).setValue(value);
      else sheet.appendRow([k, value, allowed[k][2]]);
      changed++;
    });
    invalidateConfigCache_();
    let triggerMsg = '';
    if (settings.AUTO_POLL_MINUTES !== undefined && getAutomationTriggers_().length) {
      const r = installAutomationTriggers();
      triggerMsg = ' ' + r.message;
    }
    return { success: true, message: `${changed} setting(s) saved.${triggerMsg}` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveEmailIngestConfigWeb(rows) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAutomationSchema_(ss);
    const sheet = ss.getSheetByName(IMT_EMAIL_INGEST_SHEET);
    const last = sheet.getLastRow();
    if (last > 1) sheet.getRange(2, 1, last - 1, 6).clearContent();
    const clean = (rows || []).map(r => [
      String(r[0] || '').trim().toUpperCase(), String(r[1] || '').trim(), String(r[2] || '').trim(),
      String(r[3] || '').trim(), String(r[4] || '').trim(), configYesNo_(r[5], true) ? 'YES' : 'NO'
    ]).filter(r => r[0]);
    if (clean.length) sheet.getRange(2, 1, clean.length, 6).setValues(clean);
    return { success: true, message: `Email ingest rules saved (${clean.length} carrier(s)).` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

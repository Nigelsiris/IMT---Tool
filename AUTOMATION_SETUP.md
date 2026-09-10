# Automation Setup

This guide covers the hands-free mode of IMT: invoice emails are picked up from Gmail, filed
per carrier and RDC with the week number in the file name, reconciled, coded, tracked in the
**Invoice Register**, and any discrepancies are sent back to the carrier as a reply in the
email thread they originally sent.

## What runs automatically

Every cycle (`autoProcessCycle`, installed as a time-driven trigger) does the following:

1. **Scan Gmail** using `GMAIL_SEARCH_QUERY`. Each message with a spreadsheet attachment
   (`.xlsx`, `.xls`, `.csv`, or a `.zip` containing them) is matched to a carrier via the
   **Email Ingest Config** tab (sender address / domain first, then attachment and subject
   keywords, then the bare carrier name). Attachments are saved to the carrier root folder with
   the email's message id, thread id, sender and subject stored on the Drive file. Threads are
   labelled `IMT/Processed`; unmatched or failed threads get `IMT/Needs Review`.
   Individual messages are never ingested twice, but a thread is still watched for new
   messages, so a corrected invoice replied into a discrepancy thread is picked up.
2. **Process the root folders.** For every new file:
   - Excel is converted to a Google Sheet.
   - The RDC is read from the file name; if that fails the workbook contents are scanned.
     Files with no RDC are moved to the `Needs Review` sub-folder and registered as
     `NEEDS_REVIEW` (rename the file with the RDC and drop it back in the root folder).
   - Workbooks that contain several invoices (several invoice / summary tabs with distinct
     invoice numbers) are split and each invoice is processed on its own.
   - Invoices already present in the register (same carrier + invoice number) are moved to the
     `Duplicates` sub-folder and registered as `DUPLICATE`.
   - The invoice week is derived from the line-item dates (majority week), falling back to the
     carrier's stated period end and then the email date. The file is renamed using
     `FILENAME_WEEK_FORMAT` (default `CW{WW} - {FILENAME}`) and moved to the RDC sub-folder.
   - Reconciliation, haulier write-back, GL coding and the coded PDF work exactly as before.
     With `AUTO_FINALIZE_MODE = REVIEW` the invoice is queued in *Invoice Finalization* instead.
3. **Register every invoice** in the `Invoice Register` sheet with status, totals, TU count,
   discrepancy count, week, links to the file / PDF / email thread and the discrepancy-email
   state.
4. **Reply to the carrier** for invoices with blocking discrepancies. The reply goes into the
   original email thread (grouped per thread when one email carried several invoices). With
   `AUTO_SEND_DISCREPANCY_EMAILS = NO` a draft reply is created in the thread instead. Invoices
   that did not arrive by email use the carrier's *Reply-To Override* from the Email Ingest
   Config, or an unaddressed draft as a last resort.
5. **Log** the cycle in the `Automation Log` sheet and on the Automation view.

Scheduled runs append to the tracker sheets; manual "Extract & Load" runs still archive and
clear them first (controlled by `AUTO_ARCHIVE_ON_MANUAL_RUN`).

## One-time setup

1. Add `automation.js` to the Apps Script project next to `code.js` and `Index.html`
   (or push with `clasp`). Replace the project manifest with `appsscript.json` from this repo
   (Apps Script editor → Project Settings → *Show "appsscript.json" manifest file in editor*).
2. **Grant the new permissions.** This build uses Gmail, URL fetch (PDF export) and triggers,
   which your earlier authorization did not cover. A web app that is missing scopes does not
   prompt; it fails with `You do not have permission to call UrlFetchApp.fetch` (or
   `GmailApp...`). Either
   - open the Apps Script editor, select the function `authorizeImt` and click **Run**, then
     approve the consent screen, or
   - open the web app: the dashboard shows an **Authorization Required** alert with a
     **Grant permissions** link that opens the same consent screen.
   Then create a new deployment version (Deploy → Manage deployments → Edit → Version: New).
   Any invoice that failed before the grant is in the carrier's `Unprocessed Temporary`
   folder; move it back to the carrier root folder and it will be processed on the next run.
3. Open the web app and click **Initialize Config**. This adds the automation settings to
   *System Config* and creates the `Invoice Register`, `Automation Log` and
   `Email Ingest Config` tabs.
4. Go to **Automation** and fill in the **Email Ingest Rules**: one row per carrier with the
   sender addresses or domains the invoices come from (for example
   `ap@hbtrucking.com, hbtrucking.com`). Save.
5. Review the **Automation Settings** (interval, AUTO vs REVIEW mode, auto-send, CC list,
   week numbering scheme, file name format) and save.
6. Click **Install / Update Schedule**. Triggers run as the account that installs them, so
   install from the account whose inbox receives the invoices.
7. Click **Run Cycle Now** once and watch the Activity Log and the Invoice Register.

Use **Scan Inbox Only** to fetch attachments without processing, and the Sheets menu
`🚀 Invoice Automation → Run Automation Cycle Now` when the web app is not open.

## Settings reference (System Config)

| Setting | Default | Meaning |
| --- | --- | --- |
| `AUTO_PROCESS_ENABLED` | `YES` | Master switch for the scheduled cycle. |
| `AUTO_POLL_MINUTES` | `15` | Trigger interval. Apps Script allows 1, 5, 10, 15, 30 minutes or whole hours; other values are rounded. |
| `AUTO_FINALIZE_MODE` | `AUTO` | `AUTO` stamps GL codes and exports PDFs. `REVIEW` queues invoices in Invoice Finalization. |
| `AUTO_RUN_TIME_BUDGET_SEC` | `300` | Files not started within this budget wait for the next cycle, so a run never hits the 6-minute limit. |
| `GMAIL_SEARCH_QUERY` | `has:attachment newer_than:30d -in:trash -in:spam` | Base Gmail search. Add `from:` or `label:` filters to narrow it. |
| `GMAIL_PROCESSED_LABEL` / `GMAIL_REVIEW_LABEL` | `IMT/Processed` / `IMT/Needs Review` | Labels applied to threads. |
| `GMAIL_MAX_THREADS_PER_RUN` | `50` | Threads examined per cycle, most recently active first. |
| `AUTO_SEND_DISCREPANCY_EMAILS` | `YES` | `NO` creates draft replies instead of sending. |
| `DISCREPANCY_EMAIL_INCLUDE_WARNINGS` | `NO` | `YES` also emails non-blocking warnings such as Amount Mismatch. |
| `DISCREPANCY_EMAIL_REPLY_ALL` | `NO` | Reply to everyone on the original email. |
| `DISCREPANCY_EMAIL_CC` / `DISCREPANCY_EMAIL_BCC` | blank | Extra recipients on every notice. |
| `WEEK_NUMBER_SCHEME` | `ISO` | `ISO`, `US_SUNDAY` or `US_MONDAY`. Also used for the Additional Cost write-back week lookup order. |
| `FILENAME_WEEK_FORMAT` | `CW{WW} - {FILENAME}` | Tokens: `{WW}` `{W}` `{YYYY}` `{YY}` `{CARRIER}` `{RDC}` `{INVOICE}` `{FILENAME}`. Files that already start with `CW..` are not renamed again. |
| `SPLIT_MULTI_INVOICE_WORKBOOKS` | `YES` | Split workbooks containing several invoices. |
| `UNKNOWN_RDC_FOLDER_NAME` | `Needs Review` | Sub-folder for files without a detectable RDC. |
| `DUPLICATE_FOLDER_NAME` | `Duplicates` | Sub-folder for already-registered invoices. |
| `AUTO_ARCHIVE_ON_MANUAL_RUN` | `YES` | Manual runs archive and clear the tracker sheets first. |

The Email Template tab now supports `{InvoiceNumber}`, `{Week}` and `{RDC}` in addition to
`{FileName}` and `{CarrierName}`.

## Invoice Register statuses

| Status | Meaning | What to do |
| --- | --- | --- |
| `CODED` | GL codes stamped and PDF exported. | Nothing. |
| `PENDING_REVIEW` | Waiting in Invoice Finalization (`REVIEW` mode). | Review and finalize; the register updates to `CODED`. |
| `DISCREPANCY` | Blocking discrepancies; notice sent or drafted in the thread. | Wait for the carrier. Use **Resend** or **Resolve** from the register. |
| `NEEDS_REVIEW` | RDC could not be detected. | Rename the file with the RDC and move it back to the carrier root folder. |
| `DUPLICATE` | Carrier + invoice number already registered. | Nothing, unless the earlier row was wrong. |
| `SKIPPED` | No cost line items found on the invoice tab. | Check the Header Config aliases for that carrier. |
| `ERROR` | Conversion or parsing failed; file moved to `Unprocessed Temporary`. | Check the Automation Log detail. |
| `RESOLVED` | Marked resolved from the register. | Nothing. |

## Tracker sheet changes

- `TMST` gained an **Invoice Number** column (16) and `Discrepancy Tracker` gained **Carrier**
  and **Invoice Number** columns (7, 8). New sheets get proper headers; existing sheets get the
  extra header labels automatically. Invoice Results now group per invoice instead of per
  carrier + RDC.
- The haulier write-back now writes the invoice and amount columns in two batched writes
  instead of one API call per cell; formulas in untouched cells are preserved.
- Haulier workbooks are only opened for RDCs that actually received an invoice, and a cycle
  with no new files returns without opening them at all.

## Troubleshooting

- **`You do not have permission to call UrlFetchApp.fetch` / `GmailApp`**: the script owner
  has not approved the new scopes. Run `authorizeImt` from the editor or use the
  **Grant permissions** link on the dashboard, then redeploy a new version (see setup step 2).
  Every run now checks this up front and stops before touching any file.

- **Nothing is ingested**: run **Scan Inbox Only** and read the Activity Log. "Email not
  matched to a carrier" means the sender is not in the Email Ingest Config; the thread is
  labelled `IMT/Needs Review`. Add the sender and remove the label to retry.
- **Replies are drafted, not sent**: `AUTO_SEND_DISCREPANCY_EMAILS` is `NO`, or the invoice did
  not arrive by email and no Reply-To Override is configured.
- **Trigger stopped**: triggers are owned by the account that installed them; reinstall from
  the Automation view after changing deployment ownership. Failed cycles appear on the
  dashboard as *Last Automation Cycle Failed*.
- **Files reprocessed every cycle**: a file that fails conversion is moved to
  `Unprocessed Temporary`; a file whose RDC is unknown goes to `Needs Review`. Only files in the
  carrier root folder are picked up, so nothing loops.

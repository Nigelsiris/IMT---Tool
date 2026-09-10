# IMT---Tool

Invoice Master Tracker: a Google Apps Script web app that ingests carrier invoices, reconciles
them against the weekly haulier reports, stamps GL codes, exports coded PDFs and tracks
discrepancies.

## Files

| File | Purpose |
| --- | --- |
| `code.js` | Processing engine, configuration tabs, web app server functions. |
| `automation.js` | Hands-free layer: Gmail ingestion, Invoice Register, week-numbered file names, discrepancy replies in the carrier's thread, scheduled trigger. |
| `Index.html` | Web app UI (Dashboard, Automation, Invoice Register, Operations, Finalization, Results, Data Viewer, Configuration). |
| `appsscript.json` | Project manifest: OAuth scopes (Sheets, Drive, Gmail, URL fetch, triggers) and the Drive advanced service. |

All files belong in the same Apps Script project.

## Documentation

- [Automation Setup](AUTOMATION_SETUP.md) — scheduled processing, email ingestion, Invoice Register, discrepancy replies.
- [Custom Rule How-To](CUSTOM_RULES_HOWTO.md) — building discrepancy rules from the UI.

## Quick start

1. After updating the code, run `authorizeImt` once from the Apps Script editor (or use the
   **Grant permissions** link on the dashboard) and redeploy a new version.
2. Open the web app, click **Initialize Config** and fill in the System Config tab
   (carrier root folders, haulier sheet IDs).
3. In **Automation**, add the sender addresses for each carrier under *Email Ingest Rules*,
   then click **Install / Update Schedule**.
4. Invoices arriving by email are now filed, reconciled, coded and registered automatically;
   carriers receive discrepancy notices as replies in their own email thread.

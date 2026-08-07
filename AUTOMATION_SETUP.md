# Controlled Gmail → Sheet sync

The dashboard remains a read-only consumer of the Google Sheet. A sheet-bound Apps Script is the sole Gmail writer.

1. Open the workbook and choose **Extensions → Apps Script**.
2. Replace `Code.gs` with this repository's `Code.gs`; set the manifest to `appsscript.json`.
3. Run `installAutomation()` once and authorize Gmail + Sheets access.
4. Confirm the workbook timezone is **Asia/Kolkata**.

The installer creates `Automation Config`, `Review Queue`, and `Sync Audit`, then scans Gmail every 15 minutes. Direct user edits in `Chatbot Projects` or `Commercials` become field-level overrides. A differing Gmail proposal is written to `Review Queue`; choose `Apply` or `Keep`. Nothing overridden is silently replaced.

Commercial calculations remain formulas in the Sheet. The public dashboard only reads rendered values and therefore reflects accepted edits on its next 60-second refresh.

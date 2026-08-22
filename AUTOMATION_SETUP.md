# Client classification refresh

The workbook automation no longer reads Gmail and never creates project rows.

1. Open the workbook and choose **Extensions → Apps Script**.
2. Replace `Code.gs` with this repository's `Code.gs`; set the manifest to `appsscript.json`.
3. Run `removeLegacyGmailAutomation()` once to remove the retired Gmail triggers.
4. Reload the workbook. Use **Client Data Refresh → Refresh legal name & industry** when required.

The refresh only updates these fields on existing `Chatbot Projects` rows:

- Client (legal client name)
- Industry
- Sub Industry

The match must be unique in `Client List`. Unmatched or ambiguous names are left unchanged. No other project data is modified.

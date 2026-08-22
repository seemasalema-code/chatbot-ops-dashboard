/**
 * Client classification refresh for the Chatbot Projects workbook.
 *
 * This script deliberately has no Gmail access and never creates, deletes, or
 * reorders projects. It only enriches existing project rows from Client List.
 */
const CLIENT_REFRESH = Object.freeze({
  PROJECTS_SHEET: 'Chatbot Projects',
  CLIENT_LIST_SHEET: 'Client List',
  PROJECT_HEADER_ROW: 1,
  CLIENT_LIST_HEADER_ROW: 1,
  PROJECT_NAME_HEADER: 'Chatbot',
  CLIENT_HEADER: 'Client',
  LEGAL_CLIENT_HEADER: 'Client Name',
  INDUSTRY_HEADER: 'Industry',
  SUB_INDUSTRY_HEADER: 'Sub Industry'
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Client Data Refresh')
    .addItem('Refresh legal name & industry', 'refreshClientDetails')
    .addSeparator()
    .addItem('Remove old Gmail automation', 'removeLegacyGmailAutomation')
    .addToUi();
}

/**
 * Updates only Client, Industry, and Sub Industry on existing project rows.
 * A row is changed only when its current client has one unique normalized match
 * in Client List. Unmatched or ambiguous rows are left untouched.
 */
function refreshClientDetails() {
  const ss = SpreadsheetApp.getActive();
  const projects = ss.getSheetByName(CLIENT_REFRESH.PROJECTS_SHEET);
  const clients = ss.getSheetByName(CLIENT_REFRESH.CLIENT_LIST_SHEET);
  if (!projects || !clients) throw new Error('Chatbot Projects or Client List sheet is missing.');

  const projectHeaders = headerMap_(projects, CLIENT_REFRESH.PROJECT_HEADER_ROW);
  const clientHeaders = headerMap_(clients, CLIENT_REFRESH.CLIENT_LIST_HEADER_ROW);
  const projectNameCol = requiredColumn_(projectHeaders, CLIENT_REFRESH.PROJECT_NAME_HEADER);
  const clientCol = requiredColumn_(projectHeaders, CLIENT_REFRESH.CLIENT_HEADER);
  const industryCol = requiredColumn_(projectHeaders, CLIENT_REFRESH.INDUSTRY_HEADER);
  const subIndustryCol = requiredColumn_(projectHeaders, CLIENT_REFRESH.SUB_INDUSTRY_HEADER);
  const legalClientCol = requiredColumn_(clientHeaders, CLIENT_REFRESH.LEGAL_CLIENT_HEADER);
  const sourceIndustryCol = requiredColumn_(clientHeaders, CLIENT_REFRESH.INDUSTRY_HEADER);
  const sourceSubIndustryCol = requiredColumn_(clientHeaders, CLIENT_REFRESH.SUB_INDUSTRY_HEADER);

  const sourceCount = Math.max(0, clients.getLastRow() - CLIENT_REFRESH.CLIENT_LIST_HEADER_ROW);
  if (!sourceCount) throw new Error('Client List has no data.');
  const sourceRows = clients.getRange(
    CLIENT_REFRESH.CLIENT_LIST_HEADER_ROW + 1, 1, sourceCount, clients.getLastColumn()
  ).getDisplayValues();

  const lookup = {};
  sourceRows.forEach(row => {
    const legalName = String(row[legalClientCol - 1] || '').trim();
    if (!legalName) return;
    const key = normalizeClient_(legalName);
    const record = {
      legalName,
      industry: String(row[sourceIndustryCol - 1] || '').trim(),
      subIndustry: String(row[sourceSubIndustryCol - 1] || '').trim()
    };
    if (!lookup[key]) lookup[key] = [];
    lookup[key].push(record);
  });

  const firstDataRow = CLIENT_REFRESH.PROJECT_HEADER_ROW + 1;
  const rowCount = Math.max(0, projects.getLastRow() - CLIENT_REFRESH.PROJECT_HEADER_ROW);
  if (!rowCount) return;
  const projectRows = projects.getRange(firstDataRow, 1, rowCount, projects.getLastColumn()).getDisplayValues();
  let updated = 0;

  projectRows.forEach((row, index) => {
    const projectName = String(row[projectNameCol - 1] || '').trim();
    const currentClient = String(row[clientCol - 1] || '').trim();
    if (!projectName || !currentClient) return;
    const matches = lookup[normalizeClient_(currentClient)] || [];
    if (matches.length !== 1) return;
    const match = matches[0];
    const sheetRow = firstDataRow + index;
    projects.getRange(sheetRow, clientCol).setValue(match.legalName);
    projects.getRange(sheetRow, industryCol).setValue(match.industry);
    projects.getRange(sheetRow, subIndustryCol).setValue(match.subIndustry);
    updated++;
  });

  SpreadsheetApp.getUi().alert(
    'Client refresh complete',
    updated + ' existing project row(s) updated. No projects were added or removed.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** Removes any triggers created by the retired Gmail importer. */
function removeLegacyGmailAutomation() {
  const obsoleteHandlers = ['syncGmailProjects', 'captureManualOverride'];
  ScriptApp.getProjectTriggers()
    .filter(trigger => obsoleteHandlers.indexOf(trigger.getHandlerFunction()) !== -1)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  SpreadsheetApp.getUi().alert('Old Gmail automation removed.');
}

function headerMap_(sheet, rowNumber) {
  const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  values.forEach((value, index) => {
    const header = String(value || '').trim();
    if (header) map[header] = index + 1;
  });
  return map;
}

function requiredColumn_(headers, name) {
  const column = headers[name];
  if (!column) throw new Error('Required column not found: ' + name);
  return column;
}

function normalizeClient_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\b(private|pvt|public|limited|ltd|company|co|incorporated|inc)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

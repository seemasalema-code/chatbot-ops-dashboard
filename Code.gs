/**
 * Gmail -> Chatbot Projects controlled sync.
 * Install as a bound Apps Script project on the workbook, set the project
 * timezone to Asia/Kolkata, then run installAutomation() once as the owner.
 */
const OPS = Object.freeze({
  TZ: 'Asia/Kolkata',
  PROJECTS: 'Chatbot Projects',
  COMMERCIALS: 'Commercials',
  REVIEW: 'Review Queue',
  AUDIT: 'Sync Audit',
  CONFIG: 'Automation Config',
  LOOKBACK_DAYS: 30,
  PROJECT_HEADERS: {
    client: 'Client', project: 'Chatbot', salesOwner: 'Sales Connect', botType: 'Bot Type',
    aiEnabled: 'AI enabled', vendor: 'Vendor', status: 'Status',
    briefReceivedDate: 'Brief Received Date', commercialsSharedDate: 'Commercials Shared Date',
    notes: 'Notes'
  }
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gmail Project Sync')
    .addItem('Run sync now', 'syncGmailProjects')
    .addItem('Review pending changes', 'openReviewQueue')
    .addSeparator().addItem('Install / repair automation', 'installAutomation').addToUi();
}

function installAutomation() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(OPS.TZ);
  ensureControlSheets_(ss);
  ScriptApp.getProjectTriggers().filter(t => ['syncGmailProjects','captureManualOverride'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncGmailProjects').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('captureManualOverride').forSpreadsheet(ss).onEdit().create();
  PropertiesService.getDocumentProperties().setProperty('INSTALLED_AT', now_());
  SpreadsheetApp.getUi().alert('Controlled Gmail sync installed. Dates use Asia/Kolkata.');
}

function ensureControlSheets_(ss) {
  const specs = [
    [OPS.CONFIG, ['Key','Value','Description'], [
      ['GMAIL_QUERY', 'newer_than:30d -in:spam -in:trash (feasibility OR commercial OR demo OR requirement OR BRD OR go-live)', 'Gmail query used by the 15-minute sync'],
      ['AUTO_APPLY_CONFIDENCE', '0.85', 'Only non-overridden proposals at or above this confidence apply automatically'],
      ['LAST_SUCCESSFUL_SYNC', '', 'Asia/Kolkata timestamp of the latest completed scan']
    ]],
    [OPS.REVIEW, ['Review ID','Detected At','Project Key','Sheet','Row','Field','Current Value','Proposed Value','Email Subject','Email Date','Email Link','Reason','Confidence','Decision','Decision At'], []],
    [OPS.AUDIT, ['Timestamp','Action','Project Key','Sheet','Row','Field','Old Value','New Value','Source','Message ID'], []]
  ];
  specs.forEach(([name, headers, rows]) => {
    let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    if (rows.length && sh.getLastRow() === 1) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
    sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
  });
}

function captureManualOverride(e) {
  if (!e || !e.range || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  const sh = e.range.getSheet();
  if (![OPS.PROJECTS, OPS.COMMERCIALS].includes(sh.getName()) || e.range.getRow() <= (sh.getName() === OPS.COMMERCIALS ? 2 : 1)) return;
  const headerRow = sh.getName() === OPS.COMMERCIALS ? 2 : 1;
  const field = String(sh.getRange(headerRow, e.range.getColumn()).getDisplayValue()).trim();
  if (!field || /helper|total vendor cost|margin %|^margin$/i.test(field)) return;
  const projectKey = projectKeyForRow_(sh, e.range.getRow());
  const key = overrideKey_(sh.getName(), projectKey, field);
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify({value:e.range.getDisplayValue(), at:now_()}));
  audit_('MANUAL_OVERRIDE', projectKey, sh.getName(), e.range.getRow(), field, e.oldValue || '', e.value || '', 'sheet edit', '');
}

function syncGmailProjects() {
  const lock = LockService.getDocumentLock(); if (!lock.tryLock(1000)) return;
  try {
    const ss = SpreadsheetApp.getActive(); ensureControlSheets_(ss); ss.setSpreadsheetTimeZone(OPS.TZ);
    applyReviewDecisions_();
    const cfg = config_();
    const props = PropertiesService.getDocumentProperties();
    const last = props.getProperty('LAST_GMAIL_MS');
    if (!last) {
      props.setProperty('LAST_GMAIL_MS', String(Date.now()));
      setConfig_('LAST_SUCCESSFUL_SYNC', now_() + ' (checkpoint created; new mail only)');
      return;
    }
    const query = cfg.GMAIL_QUERY + (last ? ' after:' + Utilities.formatDate(new Date(+last - 86400000), OPS.TZ, 'yyyy/MM/dd') : '');
    let newest = +(last || 0);
    GmailApp.search(query, 0, 100).forEach(thread => thread.getMessages().forEach(message => {
      if (message.getDate().getTime() <= +(last || 0)) return;
      newest = Math.max(newest, message.getDate().getTime());
      const proposal = classifyEmail_(message);
      if (proposal) applyProposal_(proposal, Number(cfg.AUTO_APPLY_CONFIDENCE || .85));
    }));
    if (newest) props.setProperty('LAST_GMAIL_MS', String(newest));
    setConfig_('LAST_SUCCESSFUL_SYNC', now_());
  } finally { lock.releaseLock(); }
}

function classifyEmail_(m) {
  const subject = m.getSubject() || '', body = m.getPlainBody().slice(0, 12000), text = (subject + '\n' + body).toLowerCase();
  const participants = [m.getFrom(), m.getTo(), m.getCc()].join(', ');
  const cleanSubject = subject.replace(/^(re|fw|fwd):\s*/ig,'').replace(/\[(internal|urgent)\]/ig,'').trim();
  if (!cleanSubject) return null;
  const fields = {};
  if (/\b(brd|brief|requirement|requirements|use\s*case|query)\b/.test(text)) { fields.briefReceivedDate = dateOnly_(m.getDate()); fields.status = 'Discovery'; }
  if (/feasibilit/.test(text)) fields.status = /please|kindly|check|confirm|share/.test(text) ? 'Feasibility Awaited' : 'Feasibility Received';
  if (/commercial|pricing|quote|cost/.test(text)) fields.status = /attached|sharing|please find|final|agreed|go ahead at/.test(text) ? 'Commercials Received' : 'Commercials Awaited';
  if (/\bdemo\b/.test(text)) fields.status = /calendar|invite|scheduled| at \d/.test(text) ? 'Demo Scheduled' : 'Demo Requested';
  if (/get on a call|requirement call|discuss.*requirement|share.*availability/.test(text)) fields.status = 'Requirement Call Pending';
  if (/\b(rcs)\b/.test(text)) fields.botType = 'RCS'; else if (/whatsapp|\bwa bot\b/.test(text)) fields.botType = 'WhatsApp';
  if (/\b(ai|llm|generative|genai|chatgpt)\b/.test(text)) fields.aiEnabled = 'Yes';
  const vendor = vendorFrom_(participants + ' ' + text); if (vendor) fields.vendor = vendor;
  const sales = salesOwnerFrom_(participants); if (sales) fields.salesOwner = sales;
  const me = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (me && String(m.getFrom()).toLowerCase().includes(me) && /commercial|pricing|quote/.test(text)) fields.commercialsSharedDate = dateOnly_(m.getDate());
  if (!Object.keys(fields).length) return null;
  return { projectKey: normalize_(cleanSubject), project: cleanSubject, client: inferClient_(cleanSubject), fields,
    messageId:m.getId(), subject, date:m.getDate(), url:'https://mail.google.com/mail/u/0/#all/' + m.getId(), confidence:.88 };
}

function applyProposal_(p, threshold) {
  const sh = SpreadsheetApp.getActive().getSheetByName(OPS.PROJECTS), headers = headerMap_(sh, 1);
  let row = findProjectRow_(sh, p);
  if (!row) {
    row = sh.getLastRow()+1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
    copyRowStructure_(sh, Math.max(2, row-1), row);
    setCell_(sh,row,headers[OPS.PROJECT_HEADERS.client],p.client);
    setCell_(sh,row,headers[OPS.PROJECT_HEADERS.project],p.project);
  }
  Object.entries(p.fields).forEach(([logical,value]) => {
    const field = OPS.PROJECT_HEADERS[logical], col = headers[field]; if (!field || !col) return;
    const current = sh.getRange(row,col).getDisplayValue(); if (String(current) === String(value)) return;
    const key = overrideKey_(sh.getName(), projectKeyForRow_(sh,row), field);
    const overridden = PropertiesService.getDocumentProperties().getProperty(key);
    if (overridden || p.confidence < threshold) queueReview_(p, sh, row, field, current, value, overridden ? 'Manual override exists' : 'Below auto-apply threshold');
    else { setCell_(sh,row,col,value); audit_('AUTO_APPLY',p.projectKey,sh.getName(),row,field,current,value,'Gmail',p.messageId); }
  });
}

function applyReviewDecisions_() {
  const ss=SpreadsheetApp.getActive(), q=ss.getSheetByName(OPS.REVIEW); if (!q || q.getLastRow()<2) return;
  const values=q.getRange(2,1,q.getLastRow()-1,15).getValues();
  values.forEach((r,i)=>{
    const decision=String(r[13]||'').toLowerCase(); if (!['apply','keep'].includes(decision) || r[14]) return;
    if (decision==='apply') {
      const sh=ss.getSheetByName(r[3]), headers=headerMap_(sh, sh.getName()===OPS.COMMERCIALS?2:1), col=headers[r[5]];
      if (sh && col) { const old=sh.getRange(r[4],col).getDisplayValue(); setCell_(sh,r[4],col,r[7]); PropertiesService.getDocumentProperties().deleteProperty(overrideKey_(r[3],r[2],r[5])); audit_('REVIEW_APPLY',r[2],r[3],r[4],r[5],old,r[7],'Review Queue',''); }
    }
    q.getRange(i+2,15).setValue(now_());
  });
}

function queueReview_(p, sh, row, field, current, proposed, reason) {
  const q=SpreadsheetApp.getActive().getSheetByName(OPS.REVIEW), id=Utilities.getUuid();
  q.appendRow([id,now_(),projectKeyForRow_(sh,row),sh.getName(),row,field,current,proposed,p.subject,dateOnly_(p.date),p.url,reason,p.confidence,'Pending','']);
  q.getRange(q.getLastRow(),14).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Apply','Keep'],true).build());
}

function openReviewQueue(){ SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(OPS.REVIEW)); }
function headerMap_(sh,row){ const v=sh.getRange(row,1,1,sh.getLastColumn()).getDisplayValues()[0],m={}; v.forEach((x,i)=>{if(x.trim())m[x.trim()]=i+1;}); return m; }
function findProjectRow_(sh,p){ const vals=sh.getRange(2,1,Math.max(1,sh.getLastRow()-1),2).getDisplayValues(); const key=normalize_(p.project); for(let i=0;i<vals.length;i++) if(normalize_(vals[i][1])===key) return i+2; return 0; }
function projectKeyForRow_(sh,row){ const header=sh.getName()===OPS.COMMERCIALS?2:1,m=headerMap_(sh,header); return normalize_(sh.getRange(row,m.Chatbot||2).getDisplayValue() || sh.getRange(row,m.Client||1).getDisplayValue()); }
function overrideKey_(sheet,key,field){ return 'OVERRIDE|' + [sheet,key,field].map(normalize_).join('|'); }
function normalize_(s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim(); }
function dateOnly_(d){ return Utilities.formatDate(new Date(d),OPS.TZ,'yyyy-MM-dd'); }
function now_(){ return Utilities.formatDate(new Date(),OPS.TZ,'yyyy-MM-dd HH:mm:ss'); }
function setCell_(sh,row,col,value){ if(col) sh.getRange(row,col).setValue(value); }
function copyRowStructure_(sh,from,to){ if(from>=2){ const src=sh.getRange(from,1,1,sh.getLastColumn()),dst=sh.getRange(to,1); src.copyTo(dst,SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false); src.copyTo(dst,SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false); } }
function inferClient_(s){ return s.split(/\s[-|:]\s/)[0].trim(); }
function vendorFrom_(s){ const x=s.toLowerCase(); if(x.includes('kevit.io'))return 'Kevit'; if(x.includes('limechat.ai'))return 'Limechat'; if(x.includes('talk.ai'))return 'Talk.ai'; return ''; }
function salesOwnerFrom_(s){ const matches=[...String(s).matchAll(/([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s+[^,]*@netcore(?:cloud)?\.(?:com|ai)/g)]; return matches.length?matches[0][1]:''; }
function config_(){ const sh=SpreadsheetApp.getActive().getSheetByName(OPS.CONFIG),o={}; if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues().forEach(r=>o[r[0]]=r[1]); return o; }
function setConfig_(key,value){ const sh=SpreadsheetApp.getActive().getSheetByName(OPS.CONFIG),v=sh.getRange(1,1,sh.getLastRow(),1).getDisplayValues().flat(),i=v.indexOf(key); if(i>=0)sh.getRange(i+1,2).setValue(value);else sh.appendRow([key,value,'']); }
function audit_(action,key,sheet,row,field,oldValue,newValue,source,messageId){ const sh=SpreadsheetApp.getActive().getSheetByName(OPS.AUDIT); if(sh)sh.appendRow([now_(),action,key,sheet,row,field,oldValue,newValue,source,messageId]); }

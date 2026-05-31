// ═══════════════════════════════════════════════════════════════
//  NEXUS · RescueTap Operations — Apps Script Backend v2.0
//  Paste this entire file into Extensions → Apps Script
//  Then: Deploy → New Deployment → Web App → Anyone → Deploy
// ═══════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── LEADERSHIP (can toggle automations) ──
const LEADERS = ['abion', 'ayo', 'shina', 'darlington', 'chizaram'];

// ── n8n WEBHOOK URLs (paste yours after deploying n8n) ──
const N8N = {
  ASSIGNMENT : 'PASTE_N8N_ASSIGNMENT_WEBHOOK_URL',
  OVERDUE    : 'PASTE_N8N_OVERDUE_WEBHOOK_URL',
  EOD_NUDGE  : 'PASTE_N8N_EOD_NUDGE_WEBHOOK_URL',
};

// ═══════════════════════════════════
//  HTTP ROUTER — doGet / doPost
// ═══════════════════════════════════

function doGet(e) {
  const action = (e.parameter.action || 'read');
  try {
    if (action === 'read')           return readAll();
    if (action === 'readTab')        return readTab(e.parameter.tab);
    if (action === 'getAutomations') return getAutomations();
    if (action === 'seed')           return seedDemoData();
    return json({ error: 'Unknown GET action: ' + action });
  } catch(err) {
    return json({ error: err.message });
  }
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  try {
    if (action === 'appendTask')        return appendRow('1. Master Tasks', body.data);
    if (action === 'appendKPI')         return appendRow('5. KPI Log', body.data);
    if (action === 'appendPartner')     return appendRow('3. Partnerships', body.data);
    if (action === 'appendCOS')         return appendRow('2. COS', body.data);
    if (action === 'appendOrg')         return appendRow('4. Org Database', body.data);
    if (action === 'appendMarketing')   return appendRow('6. Marketing', body.data);
    if (action === 'updateTaskStatus')  return updateTaskStatus(body.id, body.status, body.user);
    if (action === 'setAutomation')     return setAutomation(body.key, body.enabled, body.user);
    if (action === 'triggerOverdue')    return triggerOverdueCheck();
    return json({ error: 'Unknown POST action: ' + action });
  } catch(err) {
    return json({ error: err.message });
  }
}

// ═══════════════════════════════════
//  READ FUNCTIONS
// ═══════════════════════════════════

function readAll() {
  return json({
    tasks     : sheetToObjects('1. Master Tasks').slice(0, 200),
    cos       : sheetToObjects('2. COS'),
    partners  : sheetToObjects('3. Partnerships'),
    orgs      : sheetToObjects('4. Org Database'),
    kpi       : sheetToObjects('5. KPI Log').slice(-30),
    marketing : sheetToObjects('6. Marketing'),
    automations: getAutomationsData()
  });
}

function readTab(tabName) {
  return json({ rows: sheetToObjects(tabName) });
}

function getAutomations() {
  return json(getAutomationsData());
}

function getAutomationsData() {
  const sheet = getOrCreateSheet('8. Automation Controls', [
    'Key', 'Label', 'Description', 'Enabled', 'LastToggled', 'ToggledBy'
  ]);
  const rows = sheetToObjects('8. Automation Controls');
  // If empty, seed defaults
  if (!rows.length) {
    const defaults = [
      ['email_assignments', 'Task Assignment Emails',    'Email alert when a task is assigned to someone',      'TRUE',  '', ''],
      ['overdue_reminders', 'Overdue Task Reminders',    'Remind assignees at 2, 5, and 7 days since last update','TRUE','', ''],
      ['eod_nudges',        'EOD Report Nudges',         'Ping team members who haven\'t submitted EOD by 5PM',  'TRUE',  '', ''],
      ['n8n_webhooks',      'n8n Webhook Triggers',      'Fire webhooks to n8n for all automation flows',        'TRUE',  '', ''],
    ];
    defaults.forEach(r => sheet.appendRow(r));
    return defaults.map(r => ({ Key: r[0], Label: r[1], Description: r[2], Enabled: r[3]==='TRUE', LastToggled: r[4], ToggledBy: r[5] }));
  }
  return rows.map(r => ({ ...r, Enabled: r['Enabled'] === 'TRUE' || r['Enabled'] === true }));
}

// ═══════════════════════════════════
//  WRITE FUNCTIONS
// ═══════════════════════════════════

function appendRow(tabName, data) {
  const sheet = SS.getSheetByName(tabName);
  if (!sheet) return json({ error: 'Tab not found: ' + tabName });
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  // Auto-stamp ID if blank
  if (!row[0]) row[0] = tabName.split('.')[0].trim() + '-' + Date.now();
  sheet.appendRow(row);
  return json({ success: true, id: row[0] });
}

function updateTaskStatus(taskId, newStatus, user) {
  const sheet = SS.getSheetByName('1. Master Tasks');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol  = headers.indexOf('ID');
  const statCol = headers.indexOf('Status');
  const updCol  = headers.indexOf('Last_Updated');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(taskId)) {
      sheet.getRange(i + 1, statCol + 1).setValue(newStatus);
      if (updCol >= 0) sheet.getRange(i + 1, updCol + 1).setValue(new Date().toISOString());
      return json({ success: true, taskId, newStatus });
    }
  }
  return json({ error: 'Task not found: ' + taskId });
}

// ═══════════════════════════════════
//  AUTOMATION TOGGLE
// ═══════════════════════════════════

function setAutomation(key, enabled, user) {
  // Leadership check
  const caller = (user || '').toLowerCase();
  if (!LEADERS.some(l => caller.includes(l))) {
    return json({ error: 'Not authorised. Only leadership can toggle automations.' });
  }

  const sheet = SS.getSheetByName('8. Automation Controls');
  if (!sheet) return json({ error: 'Automation Controls tab missing' });

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol  = headers.indexOf('Key');
  const enCol   = headers.indexOf('Enabled');
  const tsCol   = headers.indexOf('LastToggled');
  const byCol   = headers.indexOf('ToggledBy');

  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      sheet.getRange(i + 1, enCol  + 1).setValue(enabled ? 'TRUE' : 'FALSE');
      sheet.getRange(i + 1, tsCol  + 1).setValue(new Date().toISOString());
      sheet.getRange(i + 1, byCol  + 1).setValue(user);
      return json({ success: true, key, enabled });
    }
  }
  return json({ error: 'Automation key not found: ' + key });
}

function isAutomationEnabled(key) {
  try {
    const rows = sheetToObjects('8. Automation Controls');
    const row  = rows.find(r => r['Key'] === key);
    return row ? (row['Enabled'] === 'TRUE' || row['Enabled'] === true) : true;
  } catch(e) { return true; }
}

// ═══════════════════════════════════
//  TRIGGERS (called by time-based triggers)
// ═══════════════════════════════════

// AUTO-STAMP Last_Updated on every edit to Master Tasks
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== '1. Master Tasks') return;
  if (e.range.getRow() === 1) return;

  // Stamp Last_Updated (find column index dynamically)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updIdx  = headers.indexOf('Last_Updated');
  if (updIdx >= 0) sheet.getRange(e.range.getRow(), updIdx + 1).setValue(new Date().toISOString());

  // If Assigned_To_Email column just filled → fire assignment email via n8n
  const emailIdx = headers.indexOf('Assigned_To_Email');
  if (e.range.getColumn() === emailIdx + 1 && e.value) {
    if (!isAutomationEnabled('email_assignments')) return;
    if (!isAutomationEnabled('n8n_webhooks')) return;
    const row = sheet.getRange(e.range.getRow(), 1, 1, headers.length).getValues()[0];
    const payload = {};
    headers.forEach((h, i) => payload[h] = row[i]);
    pingN8N(N8N.ASSIGNMENT, payload);
  }
}

// DAILY OVERDUE CHECK — set this as a Time-driven trigger (daily, 8 AM)
function dailyOverdueCheck() {
  if (!isAutomationEnabled('overdue_reminders')) return;
  if (!isAutomationEnabled('n8n_webhooks')) return;
  triggerOverdueCheck();
}

function triggerOverdueCheck() {
  const today    = new Date();
  const overdue  = sheetToObjects('1. Master Tasks').filter(task => {
    if (!task['Assigned_To_Email']) return false;
    if (['Done', 'Cancelled'].includes(task['Status'])) return false;
    const last = new Date(task['Last_Updated'] || task['Start Date'] || today);
    const days = Math.floor((today - last) / 86400000);
    task.daysSinceUpdate = days;
    return days === 2 || days === 5 || days === 7;
  });

  if (overdue.length) pingN8N(N8N.OVERDUE, { tasks: overdue, triggeredAt: today.toISOString() });
  return json({ success: true, overdueCount: overdue.length });
}

// EOD NUDGE — set as Time-driven trigger (daily, 4:30 PM)
function dailyEODNudge() {
  if (!isAutomationEnabled('eod_nudges')) return;
  if (!isAutomationEnabled('n8n_webhooks')) return;

  const today      = new Date().toDateString();
  const reported   = new Set(
    sheetToObjects('5. KPI Log')
      .filter(r => new Date(r['Date']).toDateString() === today)
      .map(r => r['Team Member'])
  );

  const allMembers = [
    'Ayo T. Joshua','Shina Samson','Israel Idowu','Darlington','Chizaram',
    'Adetola','Titi','Naruka','Bukayo','Joshua (Content)','Seth','Nora',
    'Rashidat','Muyi','Perpetual','Christiana','Mark','Nehemiah',
    'Joshua Alex','Mr Sam','Joshua (Tech Outreach)'
  ];

  const missing = allMembers.filter(m => !reported.has(m));
  if (missing.length) pingN8N(N8N.EOD_NUDGE, { missing, date: today, total: allMembers.length, submitted: reported.size });
  return json({ success: true, missing });
}

// ═══════════════════════════════════
//  DEMO SEED DATA
// ═══════════════════════════════════

function seedDemoData() {
  seedTasks();
  seedCOS();
  seedPartners();
  seedKPI();
  seedMarketing();
  return json({ success: true, message: 'Demo data seeded across all tabs.' });
}

function seedTasks() {
  const sheet = SS.getSheetByName('1. Master Tasks');
  // Clear existing data rows (keep headers)
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).clearContent();

  const tasks = [
    ['T-001','Finalize RescueTap onboarding deck for Safe School pitch','HIGH','Ayo T. Joshua','Shina Samson','COS','2026-05-28','2026-06-05','In Progress','60','ayo@rescuetap.org',new Date().toISOString(),'Key deck for ACH meeting',''],
    ['T-002','Follow up with ACH Umanah re: Safe School MoU signing','HIGH','Darlington','Ayo T. Joshua','Partnerships','2026-05-30','2026-06-03','Not Started','0','darlington@rescuetap.org',new Date().toISOString(),'',''],
    ['T-003','Create WhatsApp groups for Kano, Kaduna, and Katsina states','MEDIUM','Chizaram','Naruka','COS','2026-05-29','2026-06-07','In Progress','40','chizaram@rescuetap.org',new Date().toISOString(),'Use template message v3',''],
    ['T-004','Upload June content calendar to Google Drive','MEDIUM','Joshua (Content)','Titi','Marketing','2026-05-30','2026-06-01','Not Started','0','joshua.c@rescuetap.org',new Date().toISOString(),'',''],
    ['T-005','Deploy RescueTap v2.1 to Play Store','URGENT','Mr Sam','Israel Idowu','Tech/Product','2026-05-28','2026-06-02','In Progress','75','mrsam@rescuetap.org',new Date().toISOString(),'APK ready, pending review',''],
    ['T-006','Submit NITDA grant application — Phase 1','HIGH','Ayo T. Joshua','ABION','Grants','2026-05-25','2026-06-10','In Progress','30','ayo@rescuetap.org',new Date().toISOString(),'Requires ABION final approval',''],
    ['T-007','Recruit state reps for Rivers, Bayelsa, Delta','MEDIUM','Nehemiah','Mark','COS','2026-06-01','2026-06-15','Not Started','0','nehemiah@rescuetap.org',new Date().toISOString(),'',''],
    ['T-008','Design press release for June 12 launch campaign','MEDIUM','Titi','Rashidat','Marketing','2026-05-30','2026-06-08','Not Started','0','titi@rescuetap.org',new Date().toISOString(),'',''],
    ['T-009','Onboard Perpetual to KPI log process','LOW','Muyi','','Admin','2026-06-01','2026-06-03','Not Started','0','','',new Date().toISOString(),'New team member'],
    ['T-010','Weekly reshare / repost — all platforms 9AM','HIGH','Joshua Alex','Seth','Content & Growth','2026-06-02','2026-06-02','Not Started','0','joshuaalex@rescuetap.org',new Date().toISOString(),'Daily recurring task',''],
  ];

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  tasks.forEach(t => {
    const row = headers.map((h,i) => t[i] !== undefined ? t[i] : '');
    sheet.appendRow(row);
  });
}

function seedCOS() {
  const sheet = SS.getSheetByName('2. COS');
  const last  = sheet.getLastRow();
  if (last > 1) sheet.getRange(2,1,last-1,sheet.getLastColumn()).clearContent();

  const states = [
    ['Lagos','Bukayo Adeyemi','08012345678','bukayo@rescuetap.org','Yes','143','Active','2026-05-30',''],
    ['Abuja (FCT)','Adetola Bello','08023456789','adetola@rescuetap.org','Yes','98','Active','2026-05-29',''],
    ['Kano','Naruka Ibrahim','08034567890','naruka@rescuetap.org','Yes','67','Active','2026-05-28',''],
    ['Rivers','Seth Williams','08045678901','seth@rescuetap.org','No','0','Pending','','Rep not yet confirmed'],
    ['Oyo','Christiana Ojo','08056789012','christiana@rescuetap.org','Yes','54','Active','2026-05-27',''],
    ['Kaduna','Joshua Alex','08067890123','joshuaalex@rescuetap.org','No','0','In Progress','','Group being created'],
    ['Enugu','Rashidat Usman','08078901234','rashidat@rescuetap.org','Yes','41','Active','2026-05-26',''],
    ['Delta','Mark Eze','08089012345','mark@rescuetap.org','No','0','Pending','',''],
    ['Borno','Nehemiah Danjuma','08090123456','nehemiah@rescuetap.org','No','0','Pending','',''],
    ['Anambra','Perpetual Nwosu','08001234567','perpetual@rescuetap.org','Yes','29','Active','2026-05-25',''],
  ];

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  states.forEach(s => {
    const row = headers.map((h,i) => s[i] !== undefined ? s[i] : '');
    sheet.appendRow(row);
  });
}

function seedPartners() {
  const sheet = SS.getSheetByName('3. Partnerships');
  const last  = sheet.getLastRow();
  if (last > 1) sheet.getRange(2,1,last-1,sheet.getLastColumn()).clearContent();

  const partners = [
    ['Darlington','ACH Umanah Foundation','Dr. ACH Umanah','ach@umanah.org','NGO','MoU Pending','2026-05-28','2026-06-03','Draft Sent','Safe School partnership',''],
    ['Ayo T. Joshua','NITDA','Mr Bello Kazaure','nitda@gov.ng','Government','Active','2026-05-20','2026-06-10','Signed','Grant partner',''],
    ['Shina Samson','Red Cross Nigeria','Amina Suleiman','amina@redcross.ng','NGO','In Discussion','2026-05-15','2026-06-05','None','Emergency response MoU',''],
    ['Chizaram','Lagos State Emergency','Engr. Tunde','tunde@lasema.gov.ng','Government','MoU Pending','2026-05-22','2026-06-08','Draft Sent','State integration',''],
    ['Darlington','Dangote Foundation','CSR Desk','csr@dangote.com','Corporate','Cold','2026-04-30','2026-06-15','None','Funding pitch pending',''],
  ];

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  partners.forEach(p => {
    const row = headers.map((h,i) => p[i] !== undefined ? p[i] : '');
    sheet.appendRow(row);
  });
}

function seedKPI() {
  const sheet = SS.getSheetByName('5. KPI Log');
  const last  = sheet.getLastRow();
  if (last > 1) sheet.getRange(2,1,last-1,sheet.getLastColumn()).clearContent();

  const today = new Date();
  const fmt = d => d.toISOString().split('T')[0];
  const d0 = fmt(today);
  const d1 = fmt(new Date(today - 86400000));
  const d2 = fmt(new Date(today - 172800000));

  const logs = [
    [d0,'Sunday','Ayo T. Joshua','Deputy COO','Reviewed NITDA grant draft, sent to ABION for approval','Grant submission moved to 90%','Awaiting ABION sign-off','Complete grant final page','Y'],
    [d0,'Sunday','Shina Samson','Tech Lead','Tested v2.1 APK, logged 3 bugs to Mr Sam','Play Store review queue cleared','None','Monitor Play Store approval','Y'],
    [d0,'Sunday','Joshua (Content)','Content','Posted Sunday recap reel, 9AM reshare done','Engagement up 12%','None','Prepare June 1 content','Y'],
    [d1,'Saturday','Darlington','Outreach','Called ACH Umanah office, confirmed meeting for June 3','MoU process advancing','Decision-maker on leave','Follow-up email sent','Y'],
    [d1,'Saturday','Chizaram','COS','Created Kano WhatsApp group, added 67 members','COS Kano live','Low phone credit for team','Activate Kaduna group','Y'],
    [d2,'Friday','Naruka','COS','Recruited 3 new Kano state reps','COS Kano team strengthened','','Onboard reps to group','Y'],
    [d2,'Friday','Mark','COS','Scouted Delta state contacts','5 contacts identified','No Delta rep confirmed','Send intro messages','Y'],
  ];

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  logs.forEach(l => {
    const row = headers.map((h,i) => l[i] !== undefined ? l[i] : '');
    sheet.appendRow(row);
  });
}

function seedMarketing() {
  const sheet = SS.getSheetByName('6. Marketing');
  const last  = sheet.getLastRow();
  if (last > 1) sheet.getRange(2,1,last-1,sheet.getLastColumn()).clearContent();

  const posts = [
    ['M-001','Reel','June 12 Democracy Day — RescueTap Safety Campaign','Instagram,TikTok','Titi','2026-06-08','In Progress','','Hero creative for June launch'],
    ['M-002','Thread','How RescueTap works — 5-step explainer','Twitter/X','Joshua (Content)','2026-06-02','Not Started','',''],
    ['M-003','Graphic','COS State Rep recruitment poster','WhatsApp,Facebook','Rashidat','2026-06-01','Done','https://drive.google.com/...',''],
    ['M-004','Blog Post','RescueTap and the future of emergency response in Nigeria','Website','Seth','2026-06-05','Not Started','','SEO target: emergency app Nigeria'],
    ['M-005','Story','Daily 9AM reshare — RescueTap app download CTA','Instagram,Facebook','Joshua Alex','2026-06-01','In Progress','','Daily recurring'],
  ];

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  posts.forEach(p => {
    const row = headers.map((h,i) => p[i] !== undefined ? p[i] : '');
    sheet.appendRow(row);
  });
}

// ═══════════════════════════════════
//  UTILS
// ═══════════════════════════════════

function sheetToObjects(tabName) {
  const sheet = SS.getSheetByName(tabName);
  if (!sheet) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows
    .filter(r => r.some(c => c !== ''))
    .map(r => Object.fromEntries(headers.map((h, i) => [h, r[i] instanceof Date ? r[i].toISOString().split('T')[0] : r[i]])));
}

function getOrCreateSheet(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function pingN8N(url, payload) {
  if (!url || url.startsWith('PASTE_')) return;
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log('n8n ping failed: ' + e.message); }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════
//  SETUP: Install time-based triggers
//  Run this ONCE manually after deploying
// ═══════════════════════════════════

function installTriggers() {
  // Remove existing to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Daily overdue check at 8 AM
  ScriptApp.newTrigger('dailyOverdueCheck')
    .timeBased().everyDays(1).atHour(8).create();

  // Daily EOD nudge at 4:30 PM
  ScriptApp.newTrigger('dailyEODNudge')
    .timeBased().everyDays(1).atHour(16).create();

  Logger.log('Triggers installed: dailyOverdueCheck (8AM) + dailyEODNudge (4:30PM)');
}

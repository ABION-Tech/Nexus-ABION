// ═══════════════════════════════════════════════════════════════
//  NEXUS · ABION Industries — Apps Script Backend v3.1
//  Fixes from v3.0:
//    ✓ Added FOUNDER_EMAIL + SHEET_LINK constants
//    ✓ runFollowUpAutomation() sends direct Gmail at day 2/5/7
//    ✓ checkMissingKPI() sends direct Gmail EOD summary (no n8n dep)
//    ✓ dailyOverdueCheck now sends emails AND pings n8n
//    ✓ Partners tab header aligned ('Contact Name' consistent)
//    ✓ doPost CORS: all POST bodies parsed via e.postData.contents
//      (no Content-Type issues — Apps Script handles this server-side)
// ═══════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ─── YOUR DETAILS — fill these in ─────────────────────────────
const FOUNDER_EMAIL  = 'ayo@abion.ng';              // ← Ayo T. Joshua's email (google login resolves this to Ayo via the '9. Team Members' roster, not Oye — fixed mislabeled comment)
const COO_EMAIL      = 'abionnexus@gmail.com';           // ← Femi Balogun (test inbox)
const DEPUTY_EMAIL   = 'amara@abion.ng';         // ← Ayo T. Joshua (test inbox)
const TEAM_LEAD_EMAIL= 'ayo@abion.ng'; // ← team lead email (update if multiple)
// Display-name equivalents of the four constants above, for push routing
// (broadcastPushTo matches by name, not email) — best-guess from the email
// comments above since they're inconsistent (Ayo T. Joshua's name shows up
// against 3 different "roles"); double-check these match reality.
const FOUNDER_NAME   = 'Ayo T. Joshua';
const COO_NAME       = 'Femi Balogun';
const DEPUTY_NAME    = 'Ayo T. Joshua';
const TEAM_LEAD_NAME = 'Ayo T. Joshua';
const APP_LINK       = 'https://nexus-abion.vercel.app/';
const SHEET_LINK     = 'https://docs.google.com/spreadsheets/d/' + SS.getId();
const RESEND_FROM    = 'NEXUS · ABION <nexus@mail.abion.ng>';

// ─── Resend Accounts — priority-ordered pool with usage caps + failover ──
// '10. Resend Accounts' sheet columns:
//   Priority | Account Name | Key_Property | From_Email | Daily_Limit | Used_Today | Last_Reset_Date | Enabled
// The actual API key is NEVER stored in the sheet — Key_Property just names
// the Script Property (Project Settings → Script Properties) holding it, so
// secrets stay out of a document that gets shared/exported. Daily_Limit +
// Used_Today implement a simple per-account daily cap (reset automatically
// when Last_Reset_Date rolls to a new local day) so a maxed-out free-tier
// account gets skipped rather than erroring the whole send. If this sheet
// doesn't exist yet (pre-migration), sendViaResend falls back to the
// original single RESEND_API_KEY/RESEND_FROM constants unchanged.
const RESEND_ACCOUNTS_TAB = '10. Resend Accounts';

// One-time migration: creates the Resend Accounts tab, seeded with a single
// "Primary" row pointing at the existing RESEND_API_KEY property/RESEND_FROM
// constant, so nothing about current sends changes until more rows (backup
// Resend accounts, lower priority) are added by hand. Safe to re-run.
function migrateAddResendAccountsTab() {
  if (SS.getSheetByName(RESEND_ACCOUNTS_TAB)) { Logger.log('"' + RESEND_ACCOUNTS_TAB + '" already exists — skipping.'); return; }
  const sheet = SS.insertSheet(RESEND_ACCOUNTS_TAB);
  const headers = ['Priority','Account Name','Key_Property','From_Email','Daily_Limit','Used_Today','Last_Reset_Date','Enabled'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange(2,1,1,headers.length).setValues([[1, 'Primary', 'RESEND_API_KEY', RESEND_FROM, 100, 0, '', true]]);
  sheet.autoResizeColumns(1, headers.length);
  Logger.log('Created "' + RESEND_ACCOUNTS_TAB + '" seeded with the existing primary account. ' +
    'Add more rows for failover accounts — Priority 2, 3, etc. — each with its own ' +
    'Key_Property pointing to a Script Property holding that account\'s Resend API key.');
}

// Reads enabled accounts sorted by Priority, rolling Used_Today over to 0
// for any account whose Last_Reset_Date isn't today (both in memory and
// written back to the sheet). Returns [] entries with everything
// attemptResendSend / incrementAccountUsage need. If the tab hasn't been
// migrated yet, returns a single synthetic "legacy" account describing the
// original constants, so behavior is unchanged pre-migration.
function getResendAccountsForSending() {
  const sheet = SS.getSheetByName(RESEND_ACCOUNTS_TAB);
  if (!sheet) {
    return [{ rowIndex: -1, name: 'Primary (legacy)', keyProperty: 'RESEND_API_KEY', fromEmail: RESEND_FROM, dailyLimit: Infinity, usedToday: 0, priority: 1 }];
  }
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  const today = localDateISO();
  const accounts = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row.some(c => c !== '')) continue;
    const enabledCell = row[col['Enabled']];
    const enabled = enabledCell === true || String(enabledCell).trim().toUpperCase() === 'TRUE';
    if (!enabled) continue;
    let usedToday = Number(row[col['Used_Today']]) || 0;
    const lastResetRaw = row[col['Last_Reset_Date']];
    const lastReset = lastResetRaw ? (lastResetRaw instanceof Date ? localDateISO(lastResetRaw) : String(lastResetRaw)) : '';
    if (lastReset !== today) {
      // New day — roll the counter back to 0, both here and on the sheet,
      // so the cap applies per-day rather than accumulating forever.
      usedToday = 0;
      sheet.getRange(i + 1, col['Used_Today'] + 1).setValue(0);
      sheet.getRange(i + 1, col['Last_Reset_Date'] + 1).setValue(today);
    }
    accounts.push({
      rowIndex   : i + 1,
      name       : row[col['Account Name']] || ('Account row ' + (i + 1)),
      keyProperty: row[col['Key_Property']] || 'RESEND_API_KEY',
      fromEmail  : row[col['From_Email']] || RESEND_FROM,
      dailyLimit : row[col['Daily_Limit']] === '' || row[col['Daily_Limit']] == null ? Infinity : Number(row[col['Daily_Limit']]),
      usedToday  : usedToday,
      priority   : Number(row[col['Priority']]) || 999,
    });
  }
  accounts.sort((a, b) => a.priority - b.priority);
  return accounts;
}

// Best-effort +1 to an account's Used_Today after a confirmed send. Never
// throws for the caller — a failed counter write shouldn't undo a real send.
function incrementAccountUsage(rowIndex) {
  if (rowIndex === -1) return; // synthetic legacy account — no sheet row to touch
  try {
    const sheet = SS.getSheetByName(RESEND_ACCOUNTS_TAB);
    if (!sheet) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const usedCol = headers.indexOf('Used_Today');
    if (usedCol === -1) return;
    const cell = sheet.getRange(rowIndex, usedCol + 1);
    cell.setValue((Number(cell.getValue()) || 0) + 1);
  } catch (e) { Logger.log('incrementAccountUsage error: ' + e.message); }
}

// Single low-level Resend API call for one specific account's key/from
// address. Pure send attempt, no account bookkeeping — that's the caller's
// job (sendViaResend), so this can be reused/tested independently.
function attemptResendSend(apiKey, fromEmail, toList, subject, htmlBody, plainBody, options) {
  const payload = {
    from: fromEmail,
    to: toList,
    subject: subject,
    html: htmlBody || '<pre style="font-family:sans-serif">' + (plainBody || '') + '</pre>',
  };
  if (plainBody) payload.text = plainBody;
  if (options.replyTo) payload.reply_to = options.replyTo;
  if (options.cc) {
    const cc = Array.isArray(options.cc) ? options.cc.filter(Boolean) : [options.cc].filter(Boolean);
    if (cc.length) payload.cc = cc;
  }
  try {
    const resp = UrlFetchApp.fetch('https://api.resend.com/emails', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) return true;
    Logger.log('Resend FAILED (' + code + '): ' + resp.getContentText());
    return false;
  } catch (e) {
    Logger.log('Resend error: ' + e.message);
    return false;
  }
}

// Last-resort fallback once every Resend account is exhausted, disabled, or
// erroring: send via the script owner's own Gmail account (GmailApp — same
// OAuth scope testFullFlow() already exercises). Not as polished (no custom
// "from" display name/domain — Gmail sends as the account owner) but keeps
// the notification actually landing in an inbox instead of silently
// vanishing. Skips outright if the daily Gmail quota (shared with MailApp)
// is already exhausted, rather than throwing.
function sendViaGmailFallback(toList, subject, htmlBody, plainBody, options) {
  options = options || {};
  try {
    const quota = MailApp.getRemainingDailyQuota();
    if (quota <= 0) { Logger.log('sendViaGmailFallback SKIPPED — Gmail daily quota exhausted too'); return false; }
    const gmailOptions = { name: 'NEXUS · ABION', htmlBody: htmlBody || undefined };
    if (options.cc) {
      const cc = Array.isArray(options.cc) ? options.cc.filter(Boolean) : [options.cc].filter(Boolean);
      if (cc.length) gmailOptions.cc = cc.join(',');
    }
    if (options.replyTo) gmailOptions.replyTo = options.replyTo;
    GmailApp.sendEmail(toList.join(','), subject, plainBody || 'See HTML version.', gmailOptions);
    Logger.log('sendViaGmailFallback — sent to ' + toList.join(',') + ' via Gmail (all Resend accounts exhausted/failed)');
    return true;
  } catch (e) {
    Logger.log('sendViaGmailFallback error: ' + e.message);
    return false;
  }
}

// ─── Resend email sender — priority-ordered account failover, then Gmail ──
// Tries each enabled Resend account in the '10. Resend Accounts' sheet in
// Priority order, skipping any already at its Daily_Limit for today. First
// account that accepts the send wins (its Used_Today is bumped). If every
// account is skipped, disabled, missing its key, or actively rejects the
// send, falls back to the script owner's Gmail account as a last resort so
// the message still goes out rather than silently vanishing. Returns true
// only if something actually delivered it.
function sendViaResend(to, subject, htmlBody, plainBody, options) {
  options = options || {};
  const toList = Array.isArray(to) ? to : [to];
  const accounts = getResendAccountsForSending();

  for (const acct of accounts) {
    if (acct.usedToday >= acct.dailyLimit) {
      Logger.log('sendViaResend — "' + acct.name + '" at daily limit (' + acct.usedToday + '/' + acct.dailyLimit + '), trying next');
      continue;
    }
    const apiKey = PropertiesService.getScriptProperties().getProperty(acct.keyProperty);
    if (!apiKey) {
      Logger.log('sendViaResend — "' + acct.name + '" has no key set for property "' + acct.keyProperty + '", trying next');
      continue;
    }
    const delivered = attemptResendSend(apiKey, acct.fromEmail, toList, subject, htmlBody, plainBody, options);
    if (delivered) {
      incrementAccountUsage(acct.rowIndex);
      Logger.log('sendViaResend — sent via "' + acct.name + '" (' + (acct.usedToday + 1) + '/' + (acct.dailyLimit === Infinity ? '∞' : acct.dailyLimit) + ' today) — ' + subject);
      return true;
    }
    Logger.log('sendViaResend — "' + acct.name + '" rejected the send, trying next');
  }

  Logger.log('sendViaResend — all Resend accounts exhausted/failed, falling back to Gmail for: ' + subject);
  return sendViaGmailFallback(toList, subject, htmlBody, plainBody, options);
}


// ─── Leadership & RBAC roles ──────────────────────────────────
const LEADERS  = ['ayo','femi'];
const MANAGERS = ['ayo','femi'];
// Was: a hardcoded array, frozen at whatever the team looked like when it
// was written — confirmed stale and would silently miss anyone added
// later too. Used in
// ingestMeetingSummary()'s Fireflies assignee-matching and
// getMembersExpectedToday()'s EOD compliance check — both would silently
// fail to recognize anyone not on this frozen list. Now derived live from
// the sheet on every call, same single source of truth as everywhere else.
function getAllTeamNames() {
  return sheetToObjects('9. Team Members').map(m => String(m.Name || '').trim()).filter(Boolean);
}

const RBAC = {
  editDeadline    : MANAGERS,
  toggleAutomation: LEADERS,
  deleteRow       : LEADERS,
  get viewAll()   { return getAllTeamNames(); }, // was a stale hardcoded array; unused elsewhere but kept live rather than dangling
};

// ─── Leader passphrases (server-side only — never sent to frontend) ──
const LEADER_PASSPHRASES = {
  'AYO'    : 'falcon-nexus-7',
  'FEMI'   : 'shield-ops-4',
  'RONKE'  : 'ember-north-2',
  'BASSEY' : 'crown-rescue-1',
};

// Canonical display name for each legacy passphrase key — must match the
// full name that person gets when they sign in with Google (i.e. their
// name on "9. Team Members", the sole source of truth for names/emails).
// Without this, the legacy
// login stored the raw key ("AYO") as the chat/author name while Google
// login stored the full roster name ("Ayo T. Joshua") — same person,
// two different identities in chat, unread counts, push subs, etc.
const LEADER_DISPLAY_NAMES = {
  'AYO'    : 'Ayo T. Joshua',
  'FEMI'   : 'Femi Balogun',
  'RONKE'  : 'Ronke',
  'BASSEY' : 'Bassey Etim',
};

// Some people are logged in the Team Members roster more than once under
// slightly different names (e.g. a legacy short "Ayo" row alongside the
// canonical "Ayo T. Joshua" row, one Gmail-matched and one leaderlegacy-
// matched) — that produced two separate chat identities for the same
// person. This maps any known alias to the one canonical display name, so
// no matter which roster row/login path resolves them, everything downstream
// (chat authorship, task filters, notifications) treats them as one person.
// Add more "SHORTNAME": "Canonical Full Name" pairs here as duplicates surface.
const NAME_ALIASES = {
  'AYO'            : 'Ayo T. Joshua',
  'AYO T JOSHUA'   : 'Ayo T. Joshua',
  'AYO JOSHUA'     : 'Ayo T. Joshua',
  // ayo@abion.ng is Ayo's designated leader inbox (see
  // FOUNDER_EMAIL) — a roster row also exists under
  // "Othniel" for the same address, which must resolve to the same
  // leader identity rather than falling back to team-member access.
  'OTHNIEL'        : 'Ayo T. Joshua',
};
function canonicalizeName(name) {
  if (!name) return name;
  const key = String(name).trim().toUpperCase().replace(/\.+/g, '');
  return NAME_ALIASES[key] || name;
}

// True if `name` — however it was authenticated, legacy passphrase key or
// full Google/roster display name — belongs to a designated leader.
function isLeaderName(name) {
  if (!name) return false;
  const n = String(name).trim().toUpperCase();
  return Object.keys(LEADER_DISPLAY_NAMES).some(key =>
    key === n || LEADER_DISPLAY_NAMES[key].toUpperCase() === n
  );
}

// ─── Dynamic team email lookup ─────────────────────────────────
// '9. Team Members' is the sole source of truth — no hardcoded bootstrap
// map (mirrors the frontend's getTeamNames()/getTeamEmail(), which went
// sheet-only for the same reason: a frozen hardcoded list silently misses
// anyone added later, and risks a stale duplicate sitting next to a real
// roster row under a slightly different spelling).
// Always call this instead of reading sheetToObjects directly, so
// newly-added members resolve everywhere (assignment emails, CC, digests,
// reminders) the moment they're added.
function getTeamEmailMap() {
  const map = {};
  sheetToObjects('9. Team Members').forEach(m => {
    if (m.Name && m.Email) map[String(m.Name).trim()] = String(m.Email).trim();
  });
  return map;
}

// ─── Multi-value Assigned To / Support Person helpers ──────────
// 'Assigned To' and 'Support Person' can now hold more than one person,
// stored as a delimited string (comma / slash / ampersand / plus — same
// delimiter set the read paths already split on elsewhere in this file,
// e.g. the "My Tasks" filter and editRow permission check). Emails are
// stored comma-only, since email addresses can legitimately contain '+'.
function splitPeopleNames(str) {
  return String(str || '').split(/[,\/&+]+/).map(s => s.trim()).filter(Boolean);
}
function splitEmailList(str) {
  return String(str || '').split(',').map(s => s.trim()).filter(Boolean);
}
// "Ayo" / "Ayo and Chidi" / "Ayo, Chidi and Emeka"
function formatNameList(names) {
  const list = (names || []).filter(Boolean);
  if (list.length === 0) return '';
  if (list.length === 1) return list[0];
  if (list.length === 2) return list[0] + ' and ' + list[1];
  return list.slice(0, -1).join(', ') + ' and ' + list[list.length - 1];
}

// Is `nameUpper` (already canonicalized + uppercased) the assignee, support
// person, or supervisor on this task? Centralized so every access check
// treats "Assigned To" / "Support Person" / "Supervisor" as the multi-value
// lists they actually are, without duplicating this logic at every call site.
// Supervisor was previously compared as a single un-split value — a task
// with more than one supervisor never matched anyone but whichever name sat
// first in the field. Now goes through the same split-then-canonicalize-each
// check as the other two roles.
function isPersonOnTask(nameUpper, task) {
  if (!nameUpper || !task) return false;
  const assignedList = splitPeopleNames(task['Assigned To'])
    .map(n => canonicalizeName(n).toUpperCase());
  if (assignedList.indexOf(nameUpper) !== -1) return true;
  const supportList = splitPeopleNames(task['Support Person'])
    .map(n => canonicalizeName(n).toUpperCase());
  if (supportList.indexOf(nameUpper) !== -1) return true;
  const supervisorList = splitPeopleNames(task['Supervisor'])
    .map(n => canonicalizeName(n).toUpperCase());
  if (supervisorList.indexOf(nameUpper) !== -1) return true;
  return false;
}


// ─── Trusted Editors ─────────────────────────────────────────// Users here can edit leader-only fields (Deadline, Assigned To, Priority etc.)
// without a leader token. Add names in UPPER CASE to match passphrase keys,
// OR as full display names — both are checked.
const TRUSTED_EDITORS = [
  'Amara Chukwu',  // can edit leader-only fields without leader login
  // add more names here as needed
];

// ─── Per-tab trusted-editor grants (leader-editable, no redeploy needed) ──
// Column on "9. Team Members" holding this tab's "Y"/"N" grant. Leaders
// toggle these live from the Team Directory page (see updateTeamPermission
// below) instead of editing TRUSTED_EDITORS in code and redeploying.
const TAB_TRUST_COLUMN = {
  '1. Master Tasks' : 'Trusted_MasterTasks',
  '2. COS'           : 'Trusted_COS',
  '3. Partnerships'  : 'Trusted_Partnerships',
  // Virtual "tab" — not a real sheet, gates whether a non-leader can see/use
  // the global "+ NEW" button at all (frontend reads this flag off their
  // own roster row; see PERM_TABS / can-add-new class in index.html).
  'NewButton'        : 'Trusted_NewButton',
  // Virtual "tab" — gates the "Hand off to Agent" button on an approved
  // task (see handoffTaskToAgent below). Leaders always have it; this is
  // for extending that to specific non-leaders without a redeploy, same
  // live-toggle pattern as the three real grants above.
  'AIAgent'          : 'Trusted_AIAgent',
};

// True if `name` may bypass this tab's locked fields / row-ownership
// restriction — either via the legacy blanket TRUSTED_EDITORS list above
// (kept for back-compat with existing grants like Amara's), or a
// per-tab grant recorded on the "9. Team Members" roster.
function isTrustedForTab(name, tabName) {
  if (!name) return false;
  const canon = canonicalizeName(name).toUpperCase();
  const legacyTrusted = TRUSTED_EDITORS.some(te =>
    te.toUpperCase() === canon || te.toUpperCase() === (canon.split(' ')[0] || '').toUpperCase()
  );
  if (legacyTrusted) return true;
  const col = TAB_TRUST_COLUMN[tabName];
  if (!col) return false;
  const roster = sheetToObjects('9. Team Members');
  const row = roster.find(m => canonicalizeName(String(m.Name || '')).toUpperCase() === canon);
  return !!(row && String(row[col] || '').trim().toUpperCase() === 'Y');
}

// Session tokens: in-memory map { token -> { name, expires } }
// (Apps Script instances are short-lived; token validated each request)
const LEADER_TOKEN_PROP = 'nexus_leader_tokens'; // stored in PropertiesService

// ─── Team member auth — Google Sign-In ─────────────────────────
// Team tokens (30-day sessions) issued only after a Google ID token is
// verified AND its email matches a row in "9. Team Members". Replaces
// self-reported names for chat/push/status so nobody can act as a
// teammate without actually signing in as them.
const TEAM_TOKEN_PROP  = 'nexus_team_tokens'; // stored in PropertiesService
// ─── Team invite gate ───────────────────────────────────────────
// A roster row's Status is Pending until confirmed — googleLogin refuses to
// issue a session for a Pending row (see gate in googleLogin below). A row
// moves Pending → Active one of two ways: (1) the invitee clicks the
// confirmation link addTeamMember emailed them (confirmInvite, matched
// against a one-time token here), or (2) a leader calls approveTeamMember
// directly (covers a bounced/ignored invite email). Invite tokens expire
// after 7 days; an expired/used token just falls through to "invite not
// found" rather than silently activating anything.
const TEAM_INVITE_PROP = 'nexus_team_invites'; // stored in PropertiesService
const TEAM_INVITE_TTL_MS = 7 * 24 * 60 * 60 * 1000;
// ← Paste the OAuth Client ID from Google Cloud Console (APIs & Services →
//   Credentials → OAuth client ID → Web application). Must match the
//   data-client_id used in index.html's Google Sign-In button.
const GOOGLE_CLIENT_ID = '54897471878-72tukkee6t8p3js811ireepf16p4hmq7.apps.googleusercontent.com';
const N8N = {
  ASSIGNMENT : 'PASTE_N8N_ASSIGNMENT_WEBHOOK_URL',
  OVERDUE    : 'PASTE_N8N_OVERDUE_WEBHOOK_URL',
  EOD_NUDGE  : 'PASTE_N8N_EOD_NUDGE_WEBHOOK_URL',
};

// ─── Chat + Web Push ────────────────────────────────────────────
// Vercel serverless function that actually sends the encrypted Web Push
// (holds the VAPID private key as an env var — never put it in Apps Script).
// Deploy api/send-push.js from the nexus repo, then this URL just works
// since NEXUS is already hosted at APP_LINK.
const VERCEL_PUSH_URL = APP_LINK.replace(/\/$/, '') + '/api/send-push';

// ─── ID column per tab ────────────────────────────────────────
const TAB_ID_COL = {
  '1. Master Tasks' : 'ID',
  '2. COS'          : 'State',
  '3. Partnerships' : 'Organisation',
  '4. Org Database' : 'Org ID',
  '5. KPI Log'      : '_rowIndex',
  '6. Marketing'    : 'ID',
  '7. Meetings'     : 'Meeting_ID',
  '16. Pending AI Tasks' : 'ID',
};

// ─── Exact sheet headers ──────────────────────────────────────
const TAB_HEADERS = {
  '1. Master Tasks' : [
    'ID','Task Description','Priority','Assigned To','Support Person',
    'Business / Dept','Start Date','Deadline','Deadline_Time','Depends On','Status','% Progress',
    'Assigned_To_Email','Support_Person_Email','Last_Updated','Notes','Proof Link','Calendar_Event_ID'
  ],
  '2. COS'          : ['State','Admin Name','Admin Phone','Admin Email','WhatsApp Group Created','Members Count','Status','% Progress','Last Report Date','Notes'],
  '3. Partnerships' : [
    'Assigned To','Organisation','Contact Name','Contact Email','Category',
    'Current Status','% Progress','Last Contact Date','Follow-Up Date','Document Status','Remarks','Last_Updated'
  ],
  '4. Org Database' : ['Org ID','Visitor Name','Position','Company/Org','Email','Phone','State','Staff Size','Date Visited','Contact Owner','Letter Status','Onboarding Status','WhatsApp Added','Follow-Up Date','Notes'],
  '5. KPI Log'      : ['Date','Day','Team Member','Role','What I Completed (with proof)','What I Moved Forward','Blockers',"Tomorrow's Plan",'Report Submitted (Y/N)'],
  '6. Marketing'    : ['ID','Content Type','Topic / Caption','Platform','Assigned To','Due Date','Status','Post Link','Notes'],
  '7. Meetings'     : ['Meeting_ID','Title','Date','Time','Duration_Mins','Location_Type','Location_Address','Zoom_Link','Agenda','Team_Invitees','External_Emails','Organiser','Created_At'],
  '8. Automation Controls' : ['Key','Label','Description','Enabled','LastToggled','ToggledBy'],
  '9. Team Members'        : ['Name','Email','Role','Department','EOD_Days','Notes','Trusted_MasterTasks','Trusted_COS','Trusted_Partnerships','Trusted_NewButton','Trusted_AIAgent','Picture','Status'],
  '10. Chat Messages'      : ['ID','Channel','TaskID','Author','AuthorEmail','Message','Timestamp','MsgType'],
  '11. Push Subscriptions' : ['Name','Endpoint','Keys_JSON','UpdatedAt'],
  '12. Chat Read State'    : ['Name','Channel','LastReadTimestamp'],
  '13. Automation Log'     : ['Timestamp','Key','Label','Enabled','ToggledBy'],
  // Auto-tracked, non-self-reported activity — one row per real change a
  // person makes on a task (status flip, % progress, proof link). This is
  // the "can't lie about it" record: it's written by the server itself off
  // real edits, never typed in by the team member, and used to sanity-check
  // manual EOD reports (see computeActivityFlags()).
  '14. Auto Activity'      : ['Timestamp','Team Member','Task ID','Task','Field','Old Value','New Value'],
  // Audit trail for the per-tab trust grants above (TAB_TRUST_COLUMN) — one
  // row per change a leader makes on the Team Directory permissions UI.
  '15. Permission Log'     : ['Timestamp','GrantedBy','GrantedTo','Tab','OldValue','NewValue'],
  // Review queue for Fireflies-extracted meeting action items (see
  // ingestMeetingSummary). Nothing from this pipeline ever writes to
  // '1. Master Tasks' automatically — every row here needs a leader to
  // approve (or reject) it first via approvePendingAITask/rejectPendingAITask.
  '16. Pending AI Tasks'   : [
    'ID','Meeting_ID','Task Description','Suggested Assignee','Matched Assignee',
    'Suggested Deadline','Priority','Status','Source','Created_At','Reviewed_By','Reviewed_At','Meeting_Summary',
    'Goal_ID','Depends_On_Index',
    // v3.3 — Review Queue enrichment (NEXUS build log 5-ish): same
    // exact-match-against-roster pattern as Suggested/Matched Assignee
    // above, just for the two other people-fields a task can carry.
    // Suggested Department is NOT matched against anything — it's
    // constrained at the prompt level to DEPARTMENT_OPTIONS instead.
    'Suggested Supervisor','Matched Supervisor',
    'Suggested Support Person','Matched Support Person',
    'Suggested Department',
  ],
  // Sub-task checklist items, one row per item, linked to a parent Master
  // Tasks row via 'Task ID'. Checklist completion feeds into that task's
  // % Progress by interpolation (see PROGRESS_CEILING / computeInterpolatedProgress)
  // — it never changes which status a task is in, only where in the current
  // status's range the number sits.
  '17. Task Checklists'    : ['ID','Task ID','Description','Done','Created_At','Created_By'],
  // In-app notification feed — everything notifiable (assignments, chat,
  // escalations, deadline reminders, EOD nudges, partner follow-ups, digests)
  // gets a row here for whoever has opted into in-app notifications, so a
  // leader managing 100+ people can see it all from the bell icon instead
  // of their inbox. See NOTIFY_PREF_COL / notifyInApp() below.
  '18. Notifications'      : ['ID','Recipient','Type','Title','Body','Url','TaskID','Tag','CreatedAt','Read'],
  // Goal Decomposition feature — a Goal outlives the review step, so
  // "is this goal done" stays answerable after its child tasks are
  // approved into '1. Master Tasks'. TaskIDs is a comma-list populated
  // as each child task is approved; PendingCount caches how many
  // children are still sitting unreviewed in '16. Pending AI Tasks'.
  // Reasoning — the AI's own short explanation of how it approached this
  // goal's breakdown (see AI_GOAL_DECOMPOSE_SYSTEM_PROMPT's "reasoning"
  // field). Stored once at decomposition/regeneration time so the "show
  // your workings" panel on the goal card doesn't need a second AI call.
  '19. Goals' : ['ID','Description','CreatedBy','CreatedAt','Status','TaskIDs','PendingCount','Notes','Reasoning'],
};

// Fixed department list — used to constrain the AI's suggestedDepartment
// output (see AI_GOAL_DECOMPOSE_SYSTEM_PROMPT) to the same options the
// Master Tasks "Business / Dept" dropdown offers in index.html's
// FORM_SCHEMAS, so a suggestion is never something a leader can't select.
const DEPARTMENT_OPTIONS = ['Technology','Partnerships','Marketing','Product','Operations','Finance','HR','Content & Media','Data & Research'];

// ═════════════════════════════════════
//  HTTP ROUTER
// ═════════════════════════════════════

function doGet(e) {
  const action = (e && e.parameter) ? (e.parameter.action || 'read') : 'read';
  try {
    // Data-bearing reads now require a verified session (leader or
    // Google-matched team member) — previously anyone with the web app
    // URL could pull every sheet with no login at all.
    const READ_ACTIONS = ['read', 'readTab', 'readFiltered'];
    if (READ_ACTIONS.includes(action) && !resolveIdentity(e.parameter.token)) {
      return json({ error: 'Login required.' });
    }
    if (action === 'read')           return readAll(e.parameter.token);
    if (action === 'readTab')        return readTab(e.parameter.tab);
    if (action === 'readFiltered')   return readFiltered(e.parameter.tab, e.parameter);
    if (action === 'getAutomations') return json(getAutomationsData());
    if (action === 'getChatMessages')return getChatMessages(e.parameter.token, e.parameter.channel, e.parameter.taskId, e.parameter.sinceId);
    if (action === 'getChatSummary') return getChatSummary(e.parameter.token);
    if (action === 'getChatThreads') return getChatThreads(e.parameter.token);
    if (action === 'getPendingAITasks') return getPendingAITasks(e.parameter.token);
    if (action === 'getPolicyRegistry') return getPolicyRegistry(e.parameter.token);
    if (action === 'getGoals')          return getGoals(e.parameter.token);
    if (action === 'getGoalDetail')     return getGoalDetail(e.parameter.id, e.parameter.token);
    if (action === 'getTaskChecklist') return getTaskChecklist(e.parameter.taskId, e.parameter.token);
    if (action === 'getNotifications') return getNotifications(e.parameter.token);
    if (action === 'getAIInsights')    return getAIInsights(e.parameter.token, e.parameter.mode, e.parameter.force, e.parameter.range, e.parameter.rangeStart, e.parameter.rangeEnd, e.parameter.customPrompt);
    if (action === 'getAISettings')    return getAISettings(e.parameter.token);
    if (action === 'whoAmI')         return json(resolveIdentity(e.parameter.token) || { error: 'Login required.' });
    if (action === 'confirmInvite')  return confirmInvite(e.parameter.token);
    return json({ error: 'Unknown GET action: ' + action });
  } catch(err) { return json({ error: err.message, stack: err.stack }); }
}

function doPost(e) {
  // NOTE: Apps Script receives POST body via e.postData.contents regardless of
  // whether the client sends Content-Type or not. The CORS issue is on the
  // CLIENT side — see frontend notes. Server-side parsing is always fine here.
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(pe) { return json({ error: 'Invalid JSON body: ' + pe.message }); }
  const action = body.action;
  try {
    if (action === 'appendTask') {
      const denied = requireNewButtonAccess(body.token); // row creation itself was previously unguarded — see requireNewButtonAccess()
      if (denied) return denied;
      enforceLeaderOnlyFields(body.data, body.token, '1. Master Tasks');
      return appendRow('1. Master Tasks', body.data);
    }
    if (action === 'appendKPI')         return appendRow('5. KPI Log', body.data);
    if (action === 'appendEOD')         return appendRow('5. KPI Log', body.data);  // alias
    if (action === 'appendPartner') {
      const denied = requireNewButtonAccess(body.token);
      if (denied) return denied;
      enforceLeaderOnlyFields(body.data, body.token, '3. Partnerships'); // was previously unguarded — Assigned To / Document Status could be set by anyone on creation
      return appendRow('3. Partnerships', body.data);
    }
    if (action === 'appendCOS') {
      const denied = requireNewButtonAccess(body.token);
      if (denied) return denied;
      enforceLeaderOnlyFields(body.data, body.token, '2. COS'); // was previously unguarded — Members Count / Status could be set by anyone on creation
      return appendRow('2. COS', body.data);
    }
    if (action === 'appendOrg')         return appendRow('4. Org Database', body.data);
    if (action === 'appendMarketing')   return appendRow('6. Marketing', body.data);
    if (action === 'appendMeeting')     return appendMeeting(body.data);
    if (action === 'ingestMeetingSummary') return ingestMeetingSummary(body);
    if (action === 'approvePendingAITask') return approvePendingAITask(body.id, body.data, body.token);
    if (action === 'decomposeGoal')       return decomposeGoal(body);
    if (action === 'archiveGoal')         return archiveGoal(body.id, body.token);
    if (action === 'regenerateGoalTasks') return regenerateGoalTasks(body.id, body.token);
    if (action === 'rejectPendingAITask')  return rejectPendingAITask(body.id, body.token);
    if (action === 'handoffTaskToAgent')   return handoffTaskToAgent(body);
    // Called by n8n itself, not the frontend — see agentWorkflowCallback's
    // own doc comment for the shape n8n should POST back.
    if (action === 'agentWorkflowCallback') return agentWorkflowCallback(body);
    if (action === 'addChecklistItem')     return addChecklistItem(body.taskId, body.description, body.token);
    if (action === 'toggleChecklistItem')  return toggleChecklistItem(body.itemId, body.taskId, body.done, body.token);
    if (action === 'deleteChecklistItem')  return deleteChecklistItem(body.itemId, body.taskId, body.token);
    if (action === 'updateRow')         return updateRow(body.tabName, body.id, body.data, body.token);
    // Thin wrappers around updateRow for the Acknowledge / In Progress
    // buttons on a task's auto-posted "Task Assigned" chat message — reuse
    // every existing safeguard (RBAC ownership check, % Progress recompute,
    // activity log, and the Acknowledged-transition auto-message) rather
    // than reimplementing any of it.
    if (action === 'chatAckTask')       return updateRow('1. Master Tasks', body.taskId, { Status: 'Acknowledged' }, body.token);
    if (action === 'chatStartTask')     return updateRow('1. Master Tasks', body.taskId, { Status: 'In Progress' }, body.token);
    if (action === 'updateTask')        return updateRow('1. Master Tasks', body.id, body.data, body.token); // alias used by frontend FORM_SCHEMAS
    if (action === 'batchAppend')       return batchAppend(body.tabName, body.rows, body.token);
    if (action === 'updateTaskStatus')  return updateTaskStatus(body.id, body.status, body.token);
    if (action === 'setAutomation')     return setAutomation(body.key, body.enabled, body.token);
    if (action === 'triggerOverdue')    return requireLeader(body.token) || triggerOverdueCheck();
    if (action === 'addRecurringTask')  return addRecurringTask(body.data, body.token);
    if (action === 'syncCalendar')      return createCalendarEvents(body.user);
    if (action === 'checkDependencies') return requireLeader(body.token) || updateDependencyStatuses();
    if (action === 'sendDeadlineReport')return requireLeader(body.token) || checkDeadlinesAndNotify(true);
    if (action === 'leaderLogin')       return leaderLogin(body.name, body.passphrase);
    if (action === 'leaderLogout')      return leaderLogout(body.token);
    if (action === 'googleLogin')       return googleLogin(body.idToken);
    if (action === 'teamLogout')        return teamLogout(body.token);
    if (action === 'changeLeaderPassword') return changeLeaderPassword(body.token, body.currentPassphrase, body.newPassphrase);
    if (action === 'deleteRow')         return deleteRowAction(body.tabName, body.id, body.token);
    if (action === 'resendAssignment')  return resendAssignmentEmail(body.id, body.token);
    if (action === 'addTeamMember')     return addTeamMember(body.data, body.token);
    if (action === 'approveTeamMember') return approveTeamMember(body.name, body.token);
    if (action === 'updateTeamPermission') return updateTeamPermission(body.name, body.tabName, body.grant, body.token);
    if (action === 'setAssignNotifyMode') return setAssignNotifyMode(body.token, body.name, body.mode);
    if (action === 'getShareLink')      return requireAnyAuth(body.token) || getShareLink(body.url);
    if (action === 'sendChatMessage')   return sendChatMessage(body.data, body.token);
    if (action === 'markChatRead')      return markChatRead(body.token, body.channel);
    if (action === 'savePushSubscription')   return savePushSubscription(body.token, body.subscription);
    if (action === 'removePushSubscription') return removePushSubscription(body.token);
    if (action === 'saveNotifyPref')            return saveNotifyPref(body.token, body.pref);
    if (action === 'saveDigestNotifyPref')      return saveDigestNotifyPref(body.token, body.pref);
    if (action === 'markNotificationRead')      return markNotificationRead(body.token, body.id);
    if (action === 'markAllNotificationsRead')  return markAllNotificationsRead(body.token);
    if (action === 'sendStateSecurityReport') return requireAnyAuth(body.token) || sendStateSecurityReport(body.state, body.token);
    if (action === 'queryAIInChat')     return queryAIInChat(body);
    if (action === 'queryAIAnalystFollowup') return queryAIAnalystFollowup(body);
    if (action === 'createTaskFromChat')     return createTaskFromChat(body);
    if (action === 'saveAISettings')    return saveAISettings(body);
    if (action === 'testAIConnection')  return testAIConnection(body);
    return json({ error: 'Unknown POST action: ' + action });
  } catch(err) { return json({ error: err.message, stack: err.stack }); }
}

// ═════════════════════════════════════
//  AI ANALYST — NVIDIA NIM (@AI in task chat)
//  Session 1 of the AI Analyst build (see HANDOFF_AI_ANALYST_SESSION.md).
//  Scope for this pass: task-chat @AI only — tightest context, cheapest on
//  tokens, most immediately demoable. General/team-chat @AI and the
//  dedicated AI Analyst view (reports / attention queue / predictions) are
//  later sessions in the same build order and are NOT implemented yet.
//
//  Required Script Properties (Project Settings → Script Properties):
//    NIM_API_KEY       — NVIDIA Developer Program API key (required)
//    NIM_MODEL_CHAT    — defaults to 'z-ai/glm-5.2' if unset
//  (NIM_MODEL_REPORT is for the Analyst-view report generator, a later
//  session — add it then, not needed for this endpoint.)
// ═════════════════════════════════════

const AI_CHAT_SYSTEM_PROMPT =
  'You are NEXUS AI — an internal operations analyst embedded in a team\'s ' +
  'task discussion thread.\n' +
  'You have access to this specific task\'s live status, progress, and ' +
  'recent history.\n' +
  'Your job is to help the operations leader understand what is happening ' +
  'and what to do next.\n' +
  'Be direct. Be specific. Name actual people, dates, and numbers where the ' +
  'data warrants it.\n' +
  'Never invent data. If a signal is not in the data provided, say nothing ' +
  'about it.\n' +
  'Output plain text only. No markdown headers. No bullet symbols — use ' +
  'numbers or dashes if you must list things.\n' +
  'Keep the answer under 150 words.';

const AI_REPORT_SYSTEM_PROMPT =
  'You are NEXUS AI, generating a weekly operations report for the ' +
  'leadership team from the real, already-computed data below.\n' +
  'Structure your answer in exactly four sections, in this order: ' +
  'Summary, What Shipped, What\'s Stuck, Key Concerns.\n' +
  'Plain text only. No markdown headers, no bullet symbols — use dashes if ' +
  'you need to list things. Label each section with its name on its own line.\n' +
  'Never invent data. If a section has nothing substantive to report, say ' +
  'so briefly in that section rather than filling space.\n' +
  'Maximum 400 words total.';

const AI_ATTENTION_SYSTEM_PROMPT =
  'You are NEXUS AI. Below is a pre-ranked list of the top attention items ' +
  'for an operations leader, already scored and ordered by urgency in code ' +
  '— your job is to narrate WHY each one matters and suggest exactly ONE ' +
  'recommended next action per item, not to re-rank or second-guess the order.\n' +
  'Never invent data or numbers beyond what is given. Name the actual people ' +
  'and tasks mentioned.\n' +
  'Output plain text only. No markdown headers. Number the items 1 to 5, ' +
  'matching the order given.\n' +
  'Keep each item to 1-2 sentences.';

const AI_PREDICTIONS_SYSTEM_PROMPT =
  'You are NEXUS AI. Below are deterministic signals computed in code — ' +
  'velocity, overload, repeat-miss, and escalation-trend — for an ' +
  'operations leader planning next week.\n' +
  'Narrate what these signals mean in plain language and what to watch next ' +
  'week. Do not invent probabilities or numbers beyond what is given.\n' +
  'If a signal category is empty ("None."), say so plainly rather than ' +
  'inventing a concern for it.\n' +
  'Output plain text only. No markdown headers. No bullet symbols — use ' +
  'dashes if you must list things.\n' +
  'Maximum 250 words.';

// NEW (v3.3) — Post-approval agent handoff. Given ONE already-approved task
// (its live status/context, same block @AI already uses in task chat) plus
// a short leader instruction, do the work directly rather than describe it.
// Reuses callNIMApi_ (fast chat model) — this is a short, scoped ask, not a
// deep multi-step plan, so latency matters more than the deeper-reasoning
// model runGoalDecomposition_ uses.
const AI_AGENT_HANDOFF_SYSTEM_PROMPT =
  'You are NEXUS Agent — an execution-focused AI assistant that a team ' +
  'leader has just handed ONE already-approved task, plus a short ' +
  'instruction about what to do with it.\\n' +
  'Do the work directly — draft the requested text, produce the requested ' +
  'plan, or answer the requested question — do not just describe what you ' +
  'would do.\\n' +
  'Stay strictly scoped to this one task and instruction. Never invent ' +
  'facts about the task, the team, or anyone in it beyond what is given.\\n' +
  'You cannot take real-world actions (sending email, editing a live file, ' +
  'changing task status) — if the instruction implies one, say so plainly ' +
  'and hand back the best draft or text a human could use to do it ' +
  'themselves instead.\\n' +
  'Output plain text only. No markdown headers. Use dashes if you must ' +
  'list things.\\n' +
  'Keep the answer under 300 words unless the instruction clearly needs ' +
  'more room (e.g. a full draft document).';

// NEW (v3.2) — Goal Decomposition. Given a leader's free-text goal plus
// the live team roster, produce a concrete, assignable task breakdown.
// Reuses callNIMApiReport_ (deeper-reasoning model), not the fast chat
// model — this is a plan a leader is about to act on, so quality matters
// more than latency here.
const AI_GOAL_DECOMPOSE_SYSTEM_PROMPT =
  'You are NEXUS AI, breaking a leader\'s goal down into a concrete task ' +
  'list for their operations team.\n' +
  'Given the goal and the current team roster below, produce 4 to 10 ' +
  'discrete, assignable tasks that together would accomplish the goal. ' +
  'Each task must be:\n' +
  '- Concrete enough that one person could start it without asking what ' +
  'it means\n' +
  '- Independently completable, not "and also handle everything else"\n' +
  '- Given a suggested priority (URGENT, HIGH, MEDIUM, or LOW) based on ' +
  'how blocking it is to the rest of the goal\n\n' +
  'If the goal names a specific person, match them against the roster ' +
  'EXACTLY as spelled there — never guess or invent a name that is not ' +
  'listed. Leave suggestedAssignee blank if no one is named or no exact ' +
  'match is found. The same exact-match-or-blank rule applies to ' +
  'suggestedSupervisor (who should be checking in on this task, if it ' +
  'clearly needs oversight beyond the assignee) and suggestedSupportPerson ' +
  '(a second person who should help but isn\'t the owner) — leave either ' +
  'blank rather than guessing if the task doesn\'t obviously need one.\n\n' +
  'For suggestedDepartment, pick the single closest match from exactly ' +
  'this list, or leave it an empty string if none clearly fits: ' +
  'Technology, Partnerships, Marketing, Product, Operations, Finance, HR, ' +
  'Content & Media, Data & Research. Never output a department not on ' +
  'this list.\n\n' +
  'If tasks have a natural dependency order, express it via a ' +
  'dependsOnIndex field pointing at another task\'s position in this same ' +
  'array (0-indexed) — do not describe sequencing in prose. Most tasks ' +
  'will have no dependency; leave dependsOnIndex null in that case.\n\n' +
  'Never invent deadlines unless the goal itself implies a timeframe — if ' +
  'so, distribute reasonable dates across the tasks; otherwise leave ' +
  'suggestedDeadline as an empty string and let a human set it.\n\n' +
  'Also include a top-level "reasoning" string: 1-3 plain-text sentences ' +
  'explaining how you approached splitting this specific goal up (why this ' +
  'many tasks, what the ordering logic is, what you deliberately left out ' +
  'or left for a human to decide). This is shown to the leader as your ' +
  'working, not hidden — write it for a human reader, not as a log line.\n\n' +
  'Output ONLY a JSON object, no prose, no markdown code fences, in ' +
  'exactly this shape:\n' +
  '{"reasoning":"...","tasks":[{"description":"...","suggestedAssignee":"",' +
  '"suggestedSupervisor":"","suggestedSupportPerson":"","suggestedDepartment":"",' +
  '"priority":"MEDIUM","suggestedDeadline":"","dependsOnIndex":null}]}';

// NEW — AI Analyst chat-driven task creation. Lightweight sibling of Goal
// Decomposition: instead of a leader describing a broad goal in the
// dedicated modal (4–10 tasks, batched under a Goal), this handles ONE
// task described inline in an ongoing AI Analyst chat via a `/task` prefix.
// Reuses callNIMApi_ (fast chat model, not the report model) — this is a
// quick inline action inside a live conversation, not a plan the leader is
// about to walk away and act on, so latency matters more than it does for
// Goal Decomposition.
const AI_CHAT_TASK_EXTRACT_SYSTEM_PROMPT =
  'You are NEXUS AI, helping a leader turn a short instruction into a ' +
  'single task for their operations team.\n' +
  'Given the leader\'s message and the current team roster below, extract ' +
  'ONE task. It must be:\n' +
  '- Concrete enough that one person could start it without asking what ' +
  'it means\n' +
  '- A single, independently completable unit of work — not a multi-part ' +
  'ask\n\n' +
  'If the message actually describes multiple distinct tasks, or a ' +
  'broader outcome that would need several tasks to achieve, do NOT ' +
  'invent a split. Instead respond with {"multiple": true} and nothing ' +
  'else, so the caller can tell the leader to describe one task at a time ' +
  'or use "Describe a Goal" for a broader objective.\n\n' +
  'If the message names a specific person, match them against the roster ' +
  'EXACTLY as spelled there — never guess or invent a name that is not ' +
  'listed. Leave suggestedAssignee blank if no one is named or no exact ' +
  'match is found.\n\n' +
  'Never invent a deadline unless the message itself implies one.\n\n' +
  'Output ONLY a JSON object, no prose, no markdown code fences, in ' +
  'exactly this shape:\n' +
  '{"description":"...","suggestedAssignee":"","priority":"MEDIUM",' +
  '"suggestedDeadline":""}\n' +
  'or, if the message describes multiple tasks:\n' +
  '{"multiple": true}';

// Leader-only. Frontend detects a message starting with "/task " in the AI
// Analyst chat input (mirrors the existing "@AI" convention in task chat —
// see queryAIInChat below), strips the prefix, and fires this action
// instead of a normal queryAIAnalystFollowup call. Writes exactly one row
// to '16. Pending AI Tasks' (Source: 'AI Analyst', no Goal_ID — always an
// independent row, never a goal-batch member) and returns a confirmation
// string for the frontend to render as a chat bubble in the same thread.
//
// NOTE: unlike queryAIInChat, this does NOT call postAIChatMessage_ — the
// AI Analyst follow-up thread (unlike task/team chat) is never persisted
// server-side; the frontend holds it in memory and resends recent turns as
// `history` (see queryAIAnalystFollowup above). So the confirmation is
// just returned in the response for the frontend to append itself, exactly
// like every other AI Analyst chat turn.
function createTaskFromChat(body) {
  const identity = resolveIdentity(body.token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'This feature is available to leaders only.' });

  const message = String(body.message || '').trim();
  if (!message) return json({ error: 'No task description provided.' });

  const liveTeam = getAllTeamNames();
  const rosterBlock = liveTeam.length ? liveTeam.join(', ') : '(no team members on roster yet)';
  const raw = callNIMApi_(AI_CHAT_TASK_EXTRACT_SYSTEM_PROMPT, 'MESSAGE:\n' + message + '\n\nTEAM ROSTER:\n' + rosterBlock);

  let parsed;
  try {
    parsed = JSON.parse(String(raw || '').trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim());
  } catch (e) {
    return json({ error: 'Could not parse that into a task — try rephrasing.' });
  }

  if (parsed.multiple) {
    return json({ error: 'That sounds like more than one task — describe them one at a time, or use "Describe a Goal" for a broader objective.' });
  }
  if (!parsed.description) return json({ error: 'Could not extract a task from that message.' });

  const rosterCanon = liveTeam.map(n => canonicalizeName(n).trim().toUpperCase());
  const rawAssignee = String(parsed.suggestedAssignee || '').trim();
  const canon = rawAssignee ? canonicalizeName(rawAssignee).trim().toUpperCase() : '';
  const matchIdx = canon ? rosterCanon.indexOf(canon) : -1;
  const matchedAssignee = matchIdx >= 0 ? liveTeam[matchIdx] : '';

  const sheet = getOrCreateSheet('16. Pending AI Tasks', TAB_HEADERS['16. Pending AI Tasks']);
  const headers = TAB_HEADERS['16. Pending AI Tasks'];
  const id = 'PAT-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4);
  const row = {
    'ID': id, 'Meeting_ID': '', 'Task Description': parsed.description,
    'Suggested Assignee': rawAssignee, 'Matched Assignee': matchedAssignee,
    'Suggested Deadline': parsed.suggestedDeadline || '', 'Priority': parsed.priority || 'MEDIUM',
    'Status': 'Pending', 'Source': 'AI Analyst', 'Created_At': new Date().toISOString(),
    'Reviewed_By': '', 'Reviewed_At': '', 'Meeting_Summary': '', 'Goal_ID': '', 'Depends_On_Index': '',
  };
  assertKnownPolicySource_(row['Source']);
  sheet.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));

  const confirmMsg = 'Queued for review: "' + parsed.description + '"'
    + (matchedAssignee ? ' → suggested for ' + matchedAssignee : ' (no assignee match — pick one in the Review Queue)') + '.';

  return json({ success: true, taskId: id, matchedAssignee: matchedAssignee, confirmMsg: confirmMsg });
}

// Leader-only. Frontend detects a message starting with "@AI" inside a task
// thread OR the general team thread, strips the prefix, and fires this
// action instead of a normal sendChatMessage. The AI's answer is written
// back into the SAME thread it was asked in as a distinct chat row
// (Author: 'NEXUS AI', MsgType: 'ai_response') so it persists on reload
// exactly like any other message.
function queryAIInChat(body) {
  const identity = resolveIdentity(body.token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'This feature is available to leaders only.' });

  const query = String(body.query || '').trim();
  if (!query) return json({ error: 'No question provided.' });

  const taskId = body.taskId || null;
  let context;
  if (taskId) {
    // Same access model as the thread itself — if you can't see the
    // thread, you can't ask the AI about it either.
    if (!hasTaskThreadAccess(identity, taskId)) return json({ error: 'Access denied.' });
    context = buildTaskChatContext_(taskId);
  } else {
    context = buildOrgChatContext_();
  }

  const raw = callNIMApi_(AI_CHAT_SYSTEM_PROMPT, context + '\n\nQuestion: ' + query);
  const answer = raw && raw.trim()
    ? raw.trim()
    : 'Sorry — I couldn\'t reach the AI service just now. Please try again in a moment.';

  postAIChatMessage_(taskId ? 'task' : 'team', taskId, answer);

  return json({ success: true, response: answer });
}

// Aggregates everything about one task into a compact plain-text block for
// the LLM: current status/progress/assignment, recent escalations (from the
// Notifications sheet, Type === 'escalation'), and the last few messages in
// this exact thread for conversational continuity. Deliberately scoped to
// ONE task — no org-wide sheet reads — so this stays fast and token-cheap.
function buildTaskChatContext_(taskId) {
  const task = sheetToObjects('1. Master Tasks').find(t => String(t['ID']) === String(taskId));
  if (!task) return 'No task found with ID ' + taskId + '. Say so plainly if asked about it.';

  const lines = [];
  lines.push('TASK #' + taskId + ': ' + (task['Task Description'] || '(no description)'));
  lines.push('Status: ' + (task['Status'] || 'Unknown'));
  lines.push('% Progress: ' + (task['% Progress'] || 0) + '%');
  lines.push('Assigned To: ' + (task['Assigned To'] || 'Unassigned'));
  if (task['Support Person']) lines.push('Support Person: ' + task['Support Person']);
  lines.push('Priority: ' + (task['Priority'] || 'Unset'));
  lines.push('Business / Dept: ' + (task['Business / Dept'] || 'Unset'));
  lines.push('Deadline: ' + (task['Deadline'] || 'None') + (task['Deadline_Time'] ? ' ' + task['Deadline_Time'] : ''));
  if (task['Depends On']) lines.push('Depends On: ' + task['Depends On']);
  if (task['Notes']) lines.push('Latest Notes: ' + String(task['Notes']).substring(0, 500));
  lines.push('Last Updated: ' + (task['Last_Updated'] || 'Unknown'));

  const escalations = sheetToObjects('18. Notifications')
    .filter(n => n.Type === 'escalation' && String(n.TaskID) === String(taskId))
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt))
    .slice(0, 5);
  if (escalations.length) {
    lines.push('');
    lines.push('Recent Escalations (most recent first):');
    escalations.forEach(e => lines.push('- ' + e.CreatedAt + ': ' + (e.Title || e.Body || '')));
  } else {
    lines.push('No escalations recorded for this task.');
  }

  const msgs = sheetToObjects('10. Chat Messages')
    .filter(m => m.Channel === 'task' && String(m.TaskID) === String(taskId) && m.Author !== 'NEXUS AI')
    .sort((a, b) => new Date(a.Timestamp) - new Date(b.Timestamp))
    .slice(-10);
  if (msgs.length) {
    lines.push('');
    lines.push('Recent Discussion (oldest to newest):');
    msgs.forEach(m => lines.push(m.Author + ': ' + String(m.Message).substring(0, 200)));
  }

  return lines.join('\n');
}

// Org-wide snapshot for @AI in the general/team chat — coarser and cheaper
// than the task-chat context on purpose (per handoff: task chat context
// is naturally scoped, org chat needs deliberate compression). Covers:
// overdue tasks (name + assignee, compressed), escalation count in the
// last 7 days, recent EOD blockers, and a lightweight at-risk flag.
//
// NOTE on scope: the "at-risk" flag below is a simple time-vs-progress
// check, NOT the full velocity/consecutive-miss/overload model described
// for the dedicated AI Analyst view's Predictions mode (a later build
// step) — that one is meant to be a deterministic, more rigorous signal
// set computed once and reused across the Predictions dashboard. This is
// just enough for the AI to have something honest to say about risk in a
// quick chat answer, not a substitute for that later feature.
function buildOrgChatContext_() {
  const tasks = sheetToObjects('1. Master Tasks');
  const today = new Date();
  const OPEN_STATUSES_EXCLUDE = ['Done', 'Cancelled'];

  const overdue = tasks.filter(t => {
    if (!t['Deadline'] || OPEN_STATUSES_EXCLUDE.includes(t['Status'])) return false;
    return new Date(t['Deadline']) < today;
  });

  const lines = [];
  lines.push('ORG-WIDE SNAPSHOT');
  lines.push('Overdue tasks: ' + overdue.length);
  if (overdue.length) {
    const names = overdue.slice(0, 15).map(t =>
      String(t['Task Description'] || '').substring(0, 50) + ' (' + (t['Assigned To'] || 'unassigned') + ')'
    );
    lines.push('Overdue list: ' + names.join('; ') + (overdue.length > 15 ? '; +' + (overdue.length - 15) + ' more' : ''));
  }

  const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  const recentEscalations = sheetToObjects('18. Notifications')
    .filter(n => n.Type === 'escalation' && new Date(n.CreatedAt) >= weekAgo);
  lines.push('Escalations in the last 7 days: ' + recentEscalations.length);

  const eodRows = sheetToObjects('5. KPI Log')
    .filter(r => r['Blockers'] && String(r['Blockers']).trim())
    .sort((a, b) => new Date(b['Date']) - new Date(a['Date']))
    .slice(0, 5);
  if (eodRows.length) {
    lines.push('');
    lines.push('Recent EOD blockers:');
    eodRows.forEach(r => lines.push('- ' + (r['Team Member'] || 'Unknown') + ': ' + String(r['Blockers']).substring(0, 150)));
  } else {
    lines.push('No recent EOD blockers reported.');
  }

  const atRisk = tasks.filter(t => {
    if (!t['Start Date'] || !t['Deadline'] || OPEN_STATUSES_EXCLUDE.includes(t['Status'])) return false;
    const start = new Date(t['Start Date']), end = new Date(t['Deadline']);
    const total = end - start;
    if (total <= 0) return false;
    const pctTimeElapsed = (today - start) / total;
    const pctDone = (parseFloat(t['% Progress']) || 0) / 100;
    return pctTimeElapsed > 0.5 && pctDone < 0.5;
  });
  lines.push('');
  lines.push('Tasks at risk of slipping (past halfway on time, under halfway on progress): ' + atRisk.length);
  if (atRisk.length) {
    lines.push(atRisk.slice(0, 10).map(t =>
      String(t['Task Description'] || '').substring(0, 50) + ' (' + (t['Assigned To'] || 'unassigned') + ', ' + (t['% Progress'] || 0) + '%)'
    ).join('; '));
  }

  return lines.join('\n');
}

// System-generated message into a task's chat thread, same pattern as
// postSystemTaskMessage (task_assigned) — but Author is 'NEXUS AI' and
// MsgType is 'ai_response' so the frontend can render a visually distinct
// bubble. No push notification: the leader who asked is already looking at
// the thread and will see the reply render in place.
function postAIChatMessage_(channel, taskId, message) {
  try {
    const sheet = getOrCreateSheet('10. Chat Messages', TAB_HEADERS['10. Chat Messages']);
    const id = 'MSG-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
    const row = {
      ID: id, Channel: channel, TaskID: channel === 'task' ? String(taskId || '') : '',
      Author: 'NEXUS AI', AuthorEmail: '',
      Message: String(message).substring(0, 2000),
      Timestamp: new Date().toISOString(),
      MsgType: 'ai_response',
    };
    const headers = TAB_HEADERS['10. Chat Messages'];
    sheet.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
    try { CacheService.getScriptCache().remove(CHAT_THREADS_CACHE_KEY); } catch(ce) {}
    return id;
  } catch (e) {
    Logger.log('postAIChatMessage_ error: ' + e.message);
    return null;
  }
}

// ═════════════════════════════════════
//  AGENT HANDOFF — NEXUS is the Router + Human-Approval layer; the actual
//  Agent that DOES things lives in n8n (see NEXUS_ARCHITECTURE notes). This
//  block is the "Router → Agent → Output" leg of that diagram:
//    approved task (human-approved input) → n8n webhook (Agent + Tools)
//    → result posted back into this task's chat (Output), either
//    synchronously (n8n replies in the same HTTP response) or
//    asynchronously (n8n calls agentWorkflowCallback below once its
//    workflow finishes, e.g. after a multi-step research/outreach graph).
//
//  Required Script Properties (Project Settings → Script Properties):
//    N8N_AGENT_WEBHOOK_URL   — the n8n Webhook node URL this posts to (required)
//    N8N_AGENT_AUTH_HEADER   — e.g. 'Authorization' (optional)
//    N8N_AGENT_AUTH_VALUE    — e.g. 'Bearer xyz' (optional, paired with above)
//    N8N_CALLBACK_SECRET     — shared secret n8n must echo back on the async
//                              callback so a stranger can't post fake agent
//                              messages into a task's chat (required for async)
// ═════════════════════════════════════

// Leader (or AIAgent-trusted) only. Frontend's "Hand off to Agent" button
// on an approved task's chat thread — takes a short free-text instruction,
// posts it into the thread as a log line, then routes the task's live
// context + that instruction to the configured n8n webhook. Whatever n8n's
// workflow actually does (research, drafting, sending, CRM updates, etc.)
// is entirely up to the workflow on the other end — NEXUS's job here is
// just: package the approved task, send it, and surface the result back
// into the same thread everyone else can already see.
function handoffTaskToAgent(body) {
  const identity = resolveIdentity(body.token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader && !isTrustedForTab(identity.name, 'AIAgent')) {
    return json({ error: 'You do not have permission to hand tasks off to the agent.' });
  }

  const taskId = String(body.taskId || '').trim();
  if (!taskId) return json({ error: 'No task specified.' });
  const instructions = String(body.instructions || '').trim();
  if (!instructions) return json({ error: 'Add a short instruction for the agent first.' });

  const task = sheetToObjects('1. Master Tasks').find(t => String(t['ID']) === taskId);
  if (!task) return json({ error: 'Task not found: ' + taskId });

  // Same access model as the thread itself — if you can't see the thread,
  // you can't hand its task off either.
  if (!hasTaskThreadAccess(identity, taskId)) return json({ error: 'Access denied.' });

  const webhookUrl = PropertiesService.getScriptProperties().getProperty('N8N_AGENT_WEBHOOK_URL');
  const shortDesc = String(task['Task Description'] || '').substring(0, 60);
  const context = buildTaskChatContext_(taskId);

  if (!webhookUrl) {
    // No n8n workflow wired up yet — fall back to a direct NIM call so the
    // button isn't dead in the meantime. Same drafting-only behavior as
    // before n8n was chosen as the real execution layer: it can produce
    // text, not take real actions. Once N8N_AGENT_WEBHOOK_URL is set, this
    // branch stops firing and every handoff goes through n8n instead.
    postSystemTaskMessage(taskId, shortDesc, 'Handed off to agent by ' + identity.name + ': ' + instructions, [], 'agent_handoff');
    const raw = callNIMApi_(AI_AGENT_HANDOFF_SYSTEM_PROMPT, context + '\n\nLEADER INSTRUCTIONS FOR THE AGENT:\n' + instructions);
    const answer = raw && raw.trim()
      ? raw.trim()
      : 'Sorry — I couldn\'t reach the fallback assistant just now. Set N8N_AGENT_WEBHOOK_URL to route this through a real agent workflow instead.';
    postAgentChatMessage_(taskId, answer + '\n\n(No agent workflow configured yet — this came from the built-in assistant, not n8n.)');
    return json({ success: true, response: answer, async: false, fallback: true });
  }

  const handoffId = 'HANDOFF-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
  postSystemTaskMessage(taskId, shortDesc, 'Handed off to agent by ' + identity.name + ': ' + instructions, [], 'agent_handoff');

  const payload = {
    handoffId: handoffId,
    taskId: taskId,
    taskContext: context,
    instructions: instructions,
    requestedBy: identity.name,
    // n8n's workflow calls this back (action:'agentWorkflowCallback') once
    // it's done, if it doesn't reply synchronously in the webhook response.
    callbackUrl: ScriptApp.getService().getUrl(),
  };

  const props = PropertiesService.getScriptProperties();
  const headers = {};
  const authHeader = props.getProperty('N8N_AGENT_AUTH_HEADER');
  const authValue  = props.getProperty('N8N_AGENT_AUTH_VALUE');
  if (authHeader && authValue) headers[authHeader] = authValue;

  let resp;
  try {
    resp = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('handoffTaskToAgent webhook fetch error: ' + e.message);
    postAgentChatMessage_(taskId, 'Could not reach the agent workflow just now (' + e.message + '). Try again in a moment.');
    return json({ error: 'Could not reach the agent workflow: ' + e.message });
  }

  const code = resp.getResponseCode();
  let resultText = '';
  try {
    const parsed = JSON.parse(resp.getContentText());
    // n8n's "Respond to Webhook" node shape varies by workflow — accept a
    // few common field names rather than dictating one.
    resultText = String(parsed.result || parsed.message || parsed.output || parsed.text || '').trim();
  } catch (e) {
    // Not JSON — a workflow that just returns plain text is fine too.
    resultText = String(resp.getContentText() || '').trim();
  }

  if (code >= 200 && code < 300 && resultText) {
    // Workflow replied synchronously with a finished result — this only
    // fits SHORT workflows (n8n's own webhook response window applies);
    // anything longer should ack immediately and use the async callback
    // path below instead.
    postAgentChatMessage_(taskId, resultText);
    return json({ success: true, response: resultText, async: false });
  }

  if (code >= 200 && code < 300) {
    // Accepted but no immediate result — normal for a real multi-step
    // agent graph (research → score → draft → human approval → send).
    // The workflow is expected to call agentWorkflowCallback() below when
    // it actually finishes.
    postSystemTaskMessage(taskId, shortDesc, 'Agent workflow started — result will post here once it finishes.', [], 'agent_handoff');
    return json({ success: true, async: true, handoffId: handoffId });
  }

  // Non-2xx from n8n — surface it plainly rather than pretending it worked.
  const errMsg = 'Agent workflow returned an error (HTTP ' + code + ').';
  postAgentChatMessage_(taskId, errMsg + (resultText ? ' ' + resultText.substring(0, 300) : ''));
  return json({ error: errMsg });
}

// Called BY n8n (not the frontend) when an async agent workflow finishes —
// pass this as `callbackUrl` in the handoff payload above, with
// action:'agentWorkflowCallback' in the POST body n8n sends back, e.g.:
//   { action:'agentWorkflowCallback', secret:'<N8N_CALLBACK_SECRET>',
//     taskId:'<taskId>', handoffId:'<handoffId>', result:'<final text>' }
// Authenticated by a shared secret (not a leader token — n8n isn't a
// logged-in NEXUS user) so a stranger can't post fake agent messages into
// a task's chat just by knowing the URL.
function agentWorkflowCallback(body) {
  const expected = PropertiesService.getScriptProperties().getProperty('N8N_CALLBACK_SECRET');
  if (!expected) return json({ error: 'No N8N_CALLBACK_SECRET configured — async callbacks are disabled until one is set.' });
  if (String(body.secret || '') !== expected) return json({ error: 'Unauthorized.' });

  const taskId = String(body.taskId || '').trim();
  const result = String(body.result || '').trim();
  if (!taskId) return json({ error: 'No taskId provided.' });
  if (!result) return json({ error: 'No result provided.' });

  postAgentChatMessage_(taskId, result);
  return json({ success: true });
}

// Agent's reply into a task thread — same shape as postAIChatMessage_ (the
// @AI answer poster) but a distinct Author/MsgType ('NEXUS Agent' /
// 'agent_response') so the frontend can render it as its own bubble style,
// visually separate from a plain @AI Q&A answer.
function postAgentChatMessage_(taskId, message) {
  try {
    const sheet = getOrCreateSheet('10. Chat Messages', TAB_HEADERS['10. Chat Messages']);
    const id = 'MSG-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
    const row = {
      ID: id, Channel: 'task', TaskID: String(taskId || ''),
      Author: 'NEXUS Agent', AuthorEmail: '',
      Message: String(message).substring(0, 2000),
      Timestamp: new Date().toISOString(),
      MsgType: 'agent_response',
    };
    const headers = TAB_HEADERS['10. Chat Messages'];
    sheet.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
    try { CacheService.getScriptCache().remove(CHAT_THREADS_CACHE_KEY); } catch(ce) {}
    return id;
  } catch (e) {
    Logger.log('postAgentChatMessage_ error: ' + e.message);
    return null;
  }
}

// ── Shared range helper for all 3 AI Analyst modes ──────────────────
// `bounds` is { start: Date, end: Date } computed client-side from the
// range dropdown (this/last week, this/last month) or null for "All time".
// `end` doubles as the "as of" reference date for every calc below instead
// of `new Date()`, so picking "Last week" answers "what did this look like
// as of last week" rather than filtering today's snapshot by an unrelated
// window. `windowDays` is the fallback lookback (in days, ending at asOf)
// used only when bounds is null, matching each mode's original default.
function resolveAsOf_(bounds) {
  return bounds && bounds.end ? new Date(bounds.end) : new Date();
}
function resolveWindow_(bounds, windowDays) {
  const asOf = resolveAsOf_(bounds);
  const start = bounds && bounds.start ? new Date(bounds.start) : new Date(asOf.getTime() - windowDays * 24 * 60 * 60 * 1000);
  return { start: start, end: asOf };
}

// ── Mode 1: Generate Report ─────────────────────────────────────────
// Aggregates task status counts, overdue detail, EOD blockers, and
// escalations (by department) into a compact plain-text block, then asks
// the report model for a 4-section narrative. Uses NIM_MODEL_REPORT
// (deeper reasoning), not the chat model — see callNIMApiReport_.
// `bounds` (optional) scopes the EOD-blocker/escalation window and the
// "as of" date used for the overdue-task list — see resolveWindow_ above.
function aggregateReportData_(bounds) {
  const tasks = sheetToObjects('1. Master Tasks');
  const win = resolveWindow_(bounds, 7);
  const today = win.end; // "as of" date for overdue calc
  const FROZEN = ['Done', 'Cancelled'];

  const statusCounts = {};
  tasks.forEach(t => {
    const s = t['Status'] || 'Unknown';
    statusCounts[s] = (statusCounts[s] || 0) + 1;
  });

  const overdue = tasks.filter(t => t['Deadline'] && !FROZEN.includes(t['Status']) && new Date(t['Deadline']) < today)
    .map(t => ({
      desc: String(t['Task Description'] || '').substring(0, 60),
      assignee: t['Assigned To'] || 'Unassigned',
      daysOverdue: Math.floor((today - new Date(t['Deadline'])) / 86400000),
      note: String(t['Notes'] || '').substring(0, 100),
    }))
    .sort((a, b) => b.daysOverdue - a.daysOverdue);

  const eodBlockers = sheetToObjects('5. KPI Log')
    .filter(r => r['Date'] && new Date(r['Date']) >= win.start && new Date(r['Date']) < win.end && r['Blockers'] && String(r['Blockers']).trim())
    .map(r => (r['Team Member'] || 'Unknown') + ': ' + String(r['Blockers']).substring(0, 150));

  const escalations = sheetToObjects('18. Notifications').filter(n => n.Type === 'escalation' && new Date(n.CreatedAt) >= win.start && new Date(n.CreatedAt) < win.end);
  const tasksById = buildTasksById_();
  const escByDept = {};
  escalations.forEach(e => {
    const t = tasksById[String(e.TaskID)];
    const dept = t ? (t['Business / Dept'] || 'Unknown') : 'Unknown';
    escByDept[dept] = (escByDept[dept] || 0) + 1;
  });

  const windowLabel = bounds ? 'SELECTED RANGE (' + win.start.toDateString() + ' – ' + win.end.toDateString() + ')' : 'LAST 7 DAYS';

  const lines = [];
  lines.push('TASK STATUS COUNTS');
  Object.keys(statusCounts).forEach(s => lines.push('- ' + s + ': ' + statusCounts[s]));

  lines.push('');
  lines.push('OVERDUE TASKS AS OF ' + today.toDateString() + ' (' + overdue.length + ' total, most overdue first)');
  overdue.slice(0, 20).forEach(o => lines.push('- ' + o.desc + ' | ' + o.assignee + ' | ' + o.daysOverdue + ' days overdue' + (o.note ? ' | last note: ' + o.note : '')));
  if (overdue.length > 20) lines.push('...and ' + (overdue.length - 20) + ' more overdue tasks.');

  lines.push('');
  lines.push('EOD BLOCKERS — ' + windowLabel + ' (' + eodBlockers.length + ' reported)');
  eodBlockers.slice(0, 25).forEach(b => lines.push('- ' + b));
  if (!eodBlockers.length) lines.push('- None reported.');

  lines.push('');
  lines.push('ESCALATIONS — ' + windowLabel + ' (' + escalations.length + ' total, by department)');
  const depts = Object.keys(escByDept);
  if (depts.length) depts.forEach(d => lines.push('- ' + d + ': ' + escByDept[d]));
  else lines.push('- None.');

  return lines.join('\n');
}

// `customPrompt` (optional): a leader-supplied line of extra instructions,
// appended (never substituted) so the format/anti-fabrication rules in
// AI_REPORT_SYSTEM_PROMPT always still apply.
function generateAIReport_(bounds, customPrompt) {
  const data = aggregateReportData_(bounds);
  const sys = customPrompt ? AI_REPORT_SYSTEM_PROMPT + '\n\nAdditional instructions from the leader for this run: ' + String(customPrompt).trim() : AI_REPORT_SYSTEM_PROMPT;
  const raw = callNIMApiReport_(sys, data);
  return raw && raw.trim() ? raw.trim() : 'Report generation failed — the AI service did not return a response. Check Logger for details and try again.';
}

// ── Mode 2: Attention Queue ─────────────────────────────────────────
// Apps Script computes and ranks every candidate item; the LLM only
// narrates why each matters and suggests one action — it never re-ranks.
const PRIORITY_WEIGHT = { URGENT: 4, HIGH: 3, MEDIUM: 2, LOW: 1 };

function aggregateAttentionData_(bounds) {
  const tasks = sheetToObjects('1. Master Tasks');
  const today = resolveAsOf_(bounds);
  const FROZEN = ['Done', 'Cancelled'];
  const items = [];

  // Overdue tasks: score = days_overdue × priority_weight
  tasks.forEach(t => {
    if (!t['Deadline'] || FROZEN.includes(t['Status'])) return;
    const deadline = new Date(t['Deadline']);
    if (deadline >= today) return;
    const daysOverdue = Math.floor((today - deadline) / 86400000);
    const weight = PRIORITY_WEIGHT[t['Priority']] || PRIORITY_WEIGHT.MEDIUM;
    items.push({
      score: daysOverdue * weight,
      detail: '[Overdue task] ' + String(t['Task Description'] || '').substring(0, 60) +
        ' — ' + (t['Assigned To'] || 'Unassigned') + ' — ' + daysOverdue + ' days overdue, priority ' + (t['Priority'] || 'MEDIUM'),
    });
  });

  // Escalations open 48h+ (as of `today`) — no resolution-state tracking in
  // the schema, so "open" here is a proxy: still unresolved by virtue of
  // age since fired.
  const twoDaysAgo = new Date(today.getTime() - 48 * 60 * 60 * 1000);
  const tasksById = buildTasksById_();
  sheetToObjects('18. Notifications')
    .filter(n => n.Type === 'escalation' && new Date(n.CreatedAt) <= twoDaysAgo)
    .forEach(n => {
      const t = tasksById[String(n.TaskID)];
      const hoursOpen = Math.floor((today - new Date(n.CreatedAt)) / 3600000);
      items.push({
        score: hoursOpen,
        detail: '[Escalation 48h+] ' + String(n.Title || n.Body || 'Escalation').substring(0, 70) +
          (t ? ' — task: ' + String(t['Task Description'] || '').substring(0, 40) : '') + ' — open ' + hoursOpen + 'h',
      });
    });

  // Team members with no EOD submission in 3+ days as of `today` (only
  // flags people with at least one prior EOD on file — no baseline means
  // no comparison, not an assumed "never reported")
  const eodRows = sheetToObjects('5. KPI Log');
  sheetToObjects('9. Team Members')
    .filter(m => String(m['Status'] || '').toLowerCase() !== 'inactive')
    .forEach(m => {
      const name = canonicalizeName(m['Name']);
      const myReports = eodRows.filter(r => canonicalizeName(r['Team Member']) === name && r['Date'] && new Date(r['Date']) <= today);
      if (!myReports.length) return;
      const lastDate = myReports.map(r => new Date(r['Date'])).sort((a, b) => b - a)[0];
      const daysSince = Math.floor((today - lastDate) / 86400000);
      if (daysSince >= 3) {
        items.push({ score: daysSince * 2, detail: '[No EOD 3+ days] ' + name + ' — last EOD ' + daysSince + ' days ago' });
      }
    });

  // Tasks with zero progress update in 5+ days as of `today`
  tasks.forEach(t => {
    if (!t['Last_Updated'] || FROZEN.includes(t['Status'])) return;
    const daysSince = Math.floor((today - new Date(t['Last_Updated'])) / 86400000);
    if (daysSince >= 5) {
      items.push({
        score: daysSince,
        detail: '[Stale 5+ days] ' + String(t['Task Description'] || '').substring(0, 60) +
          ' — ' + (t['Assigned To'] || 'Unassigned') + ' — ' + daysSince + ' days since last update',
      });
    }
  });

  // Partner follow-ups past their next-action date, as of `today`
  sheetToObjects('3. Partnerships').forEach(p => {
    if (!p['Follow-Up Date']) return;
    const dueDate = new Date(p['Follow-Up Date']);
    if (dueDate >= today) return;
    const daysLate = Math.floor((today - dueDate) / 86400000);
    items.push({
      score: daysLate * 2,
      detail: '[Partner follow-up overdue] ' + (p['Organisation'] || 'Unknown org') +
        ' — ' + (p['Assigned To'] || 'Unassigned') + ' — ' + daysLate + ' days past follow-up date',
    });
  });

  items.sort((a, b) => b.score - a.score);
  return items.slice(0, 5);
}

function generateAttentionQueue_(bounds, customPrompt) {
  const items = aggregateAttentionData_(bounds);
  if (!items.length) return 'Nothing urgent right now — no overdue tasks, stale threads, missed EODs, or overdue follow-ups above the tracked thresholds.';

  const asOfLabel = bounds ? ' (as of ' + resolveAsOf_(bounds).toDateString() + ')' : '';
  const lines = ['TOP ATTENTION ITEMS' + asOfLabel + ' (already ranked by urgency — narrate and recommend one action each, do not re-rank):'];
  items.forEach((it, i) => lines.push((i + 1) + '. ' + it.detail));

  const sys = customPrompt ? AI_ATTENTION_SYSTEM_PROMPT + '\n\nAdditional instructions from the leader for this run: ' + String(customPrompt).trim() : AI_ATTENTION_SYSTEM_PROMPT;
  const raw = callNIMApi_(sys, lines.join('\n'));
  return raw && raw.trim() ? raw.trim() : 'Attention queue generation failed — the AI service did not return a response. Check Logger for details and try again.';
}

// ── Mode 3: Predictions ─────────────────────────────────────────
// Code-first, LLM-last: every signal below is computed deterministically;
// the LLM only narrates what they mean. No fabricated probabilities.
function computePredictionSignals_(bounds) {
  const tasks = sheetToObjects('1. Master Tasks');
  const today = resolveAsOf_(bounds);
  const FROZEN = ['Done', 'Cancelled'];
  const lines = [];

  // Velocity = (progress_pct/100) / (days_elapsed/total_days)
  const likelyToSlip = [];
  const atRisk = [];
  tasks.forEach(t => {
    if (!t['Start Date'] || !t['Deadline'] || FROZEN.includes(t['Status'])) return;
    const start = new Date(t['Start Date']), end = new Date(t['Deadline']);
    const totalDays = (end - start) / 86400000;
    const daysElapsed = (today - start) / 86400000;
    if (totalDays <= 0 || daysElapsed <= 0) return;
    const pct = (parseFloat(t['% Progress']) || 0) / 100;
    const timeFrac = Math.min(daysElapsed / totalDays, 1);
    const velocity = pct / timeFrac;
    const entry = String(t['Task Description'] || '').substring(0, 60) + ' — ' + (t['Assigned To'] || 'Unassigned') + ' — velocity ' + velocity.toFixed(2);
    if (velocity < 0.4) likelyToSlip.push(entry);
    else if (velocity < 0.7) atRisk.push(entry);
  });

  // Overload flag: 5+ active tasks per person
  const activeByPerson = {};
  tasks.forEach(t => {
    if (FROZEN.includes(t['Status'])) return;
    splitPeopleNames(t['Assigned To'] || '').forEach(name => {
      const n = canonicalizeName(name);
      if (n) activeByPerson[n] = (activeByPerson[n] || 0) + 1;
    });
  });
  const overloaded = Object.keys(activeByPerson).filter(n => activeByPerson[n] >= 5).map(n => n + ' (' + activeByPerson[n] + ' active tasks)');

  // Repeat-miss proxy: 3+ tasks currently overdue for the same person.
  // NOTE: the schema doesn't retain a history of past deadline outcomes,
  // so this is a same-moment proxy for "consecutive missed deadlines", not
  // a true streak across time. Documented limitation, not hidden from the
  // leader reading the output.
  const overdueByPerson = {};
  tasks.forEach(t => {
    if (!t['Deadline'] || FROZEN.includes(t['Status']) || new Date(t['Deadline']) >= today) return;
    splitPeopleNames(t['Assigned To'] || '').forEach(name => {
      const n = canonicalizeName(name);
      if (n) overdueByPerson[n] = (overdueByPerson[n] || 0) + 1;
    });
  });
  const repeatMissers = Object.keys(overdueByPerson).filter(n => overdueByPerson[n] >= 3).map(n => n + ' (' + overdueByPerson[n] + ' overdue tasks right now)');

  // Escalation frequency trend over the last 4 weeks
  const weekBuckets = [0, 0, 0, 0]; // [this week, 1wk ago, 2wk ago, 3wk ago]
  sheetToObjects('18. Notifications').filter(n => n.Type === 'escalation').forEach(n => {
    const days = Math.floor((today - new Date(n.CreatedAt)) / 86400000);
    const week = Math.floor(days / 7);
    if (week >= 0 && week < 4) weekBuckets[week]++;
  });
  let trend = 'stable';
  if (weekBuckets[0] > weekBuckets[1] && weekBuckets[1] >= weekBuckets[2]) trend = 'rising';
  else if (weekBuckets[0] < weekBuckets[1] && weekBuckets[1] <= weekBuckets[2]) trend = 'falling';

  lines.push('VELOCITY SIGNALS');
  lines.push('Likely to slip (velocity < 0.4): ' + (likelyToSlip.length ? likelyToSlip.join('; ') : 'None.'));
  lines.push('At risk (velocity 0.4-0.7): ' + (atRisk.length ? atRisk.join('; ') : 'None.'));
  lines.push('');
  lines.push('OVERLOAD SIGNALS (5+ active tasks per person)');
  lines.push(overloaded.length ? overloaded.join('; ') : 'None.');
  lines.push('');
  lines.push('REPEAT-MISS SIGNALS (3+ currently-overdue tasks for the same person — a same-moment proxy, not a true streak across time)');
  lines.push(repeatMissers.length ? repeatMissers.join('; ') : 'None.');
  lines.push('');
  lines.push('ESCALATION TREND (last 4 weeks, most recent first): ' + weekBuckets.join(', ') + ' — trend: ' + trend);

  return lines.join('\n');
}

function generatePredictions_(bounds, customPrompt) {
  const signals = computePredictionSignals_(bounds);
  const sys = customPrompt ? AI_PREDICTIONS_SYSTEM_PROMPT + '\n\nAdditional instructions from the leader for this run: ' + String(customPrompt).trim() : AI_PREDICTIONS_SYSTEM_PROMPT;
  const raw = callNIMApi_(sys, signals);
  return raw && raw.trim() ? raw.trim() : 'Predictions generation failed — the AI service did not return a response. Check Logger for details and try again.';
}

// ── Caching layer ─────────────────────────────────────────
// Caches the last generated output per mode in Script Properties with a
// timestamp. A repeat request for the same mode within 30 minutes returns
// the cached copy instead of spending another NIM call — keeps real API
// calls well under the 40 RPM / effectively-unlimited-credit NIM tier even
// with several leaders clicking around. `force` bypasses the cache (the
// Analyst view's Refresh button sets this).
// NOTE: this deployment is single-org-per-script (each org runs its own
// Apps Script project + Sheet), so the cache key doesn't need an orgId —
// unlike a shared multi-tenant backend would.
const AI_CACHE_TTL_MS = 30 * 60 * 1000;

function getCachedOrFetch_(cacheKey, fetchFn, force) {
  const props = PropertiesService.getScriptProperties();
  const key = 'AI_CACHE_' + cacheKey;
  if (!force) {
    const cached = props.getProperty(key);
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        if (Date.now() - parsed.ts < AI_CACHE_TTL_MS) {
          return { text: parsed.text, generatedAt: parsed.ts, cached: true };
        }
      } catch (pe) { /* corrupt cache entry — fall through to a fresh fetch */ }
    }
  }
  const text = fetchFn();
  try { props.setProperty(key, JSON.stringify({ ts: Date.now(), text: text })); }
  catch (se) { Logger.log('getCachedOrFetch_ cache write error (' + cacheKey + '): ' + se.message); }
  return { text: text, generatedAt: Date.now(), cached: false };
}

// GET endpoint for the dedicated AI Analyst view — leader-only, same nav
// gate as Team Directory / Review Queue. mode is one of 'report' /
// 'attention' / 'predictions'; force='1' bypasses the 30-min cache.
// rangeStart/rangeEnd (ISO strings, optional): a leader-picked window —
// now honored by all three modes, not just Report (see resolveAsOf_/
// resolveWindow_). customPrompt (optional): extra instructions appended to
// that mode's system prompt for this run only — nothing is persisted.
function getAIInsights(token, mode, force, rangeKey, rangeStart, rangeEnd, customPrompt) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'Leader only.' });

  const generators = { report: generateAIReport_, attention: generateAttentionQueue_, predictions: generatePredictions_ };
  const fn = generators[mode];
  if (!fn) return json({ error: 'Unknown AI Analyst mode: ' + mode });

  const bounds = (rangeStart && rangeEnd) ? { start: new Date(rangeStart), end: new Date(rangeEnd) } : null;
  const prompt = customPrompt ? String(customPrompt).trim().substring(0, 500) : '';
  // Cache key folds in range + prompt so different selections don't share
  // a cached answer from a different one; a hash keeps the property key
  // short (Script Properties values aren't the limit here, keys are fine
  // too, but a raw custom prompt could contain characters best avoided in
  // a property key).
  const cacheKey = mode + '_' + (rangeKey || 'all') + (prompt ? '_p' + Utilities.base64Encode(Utilities.newBlob(prompt).getBytes()).substring(0, 24) : '');

  const result = getCachedOrFetch_(cacheKey, () => fn(bounds, prompt), force === '1' || force === true);
  return json({ success: true, mode: mode, text: result.text, generatedAt: result.generatedAt, cached: result.cached });
}

// ── AI Analyst follow-up chat ─────────────────────────────────────────
// Lets a leader ask a clarifying question about a Report / Attention Queue
// / Predictions output that's already on screen, instead of only ever
// getting a single fixed block of text. Re-aggregates the SAME underlying
// data (respecting the same range, if one was selected) so the follow-up
// answer is grounded in real numbers, not just the model guessing from the
// prior prose. Not cached — each question is a fresh, cheap chat-model
// call — and nothing here is persisted server-side; the frontend keeps
// the thread and resends recent turns as `history` for continuity.
const AI_FOLLOWUP_SYSTEM_PROMPT =
  'You are NEXUS AI, continuing a conversation about the operations output you already generated below.\n' +
  'Answer the leader\'s follow-up question using ONLY the data and prior output given — never invent facts, ' +
  'names, or numbers beyond what is present.\n' +
  'If the answer is not in the data, say so plainly rather than guessing.\n' +
  'Output plain text only. No markdown headers, no bullet symbols — use dashes if you must list things.\n' +
  'Keep the answer under 200 words.';

function queryAIAnalystFollowup(body) {
  const identity = resolveIdentity(body.token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'Leader only.' });

  const mode = body.mode;
  const contextFns = { report: aggregateReportData_, predictions: computePredictionSignals_ };
  if (!['report', 'attention', 'predictions'].includes(mode)) return json({ error: 'Unknown AI Analyst mode: ' + mode });

  const question = String(body.question || '').trim();
  if (!question) return json({ error: 'No question provided.' });
  const priorText = String(body.priorText || '').trim();

  const bounds = (body.rangeStart && body.rangeEnd) ? { start: new Date(body.rangeStart), end: new Date(body.rangeEnd) } : null;
  let dataBlock;
  if (mode === 'attention') {
    const items = aggregateAttentionData_(bounds);
    dataBlock = items.length ? items.map((it, i) => (i + 1) + '. ' + it.detail).join('\n') : 'Nothing urgent right now.';
  } else {
    dataBlock = contextFns[mode](bounds);
  }

  // Last few turns only (bounded token cost) — the frontend sends the
  // running thread for this mode, we don't store it.
  const history = Array.isArray(body.history) ? body.history.slice(-6) : [];
  const historyBlock = history.length
    ? '\n\nEARLIER FOLLOW-UP Q&A THIS SESSION:\n' + history.map(h => 'Q: ' + String(h.q || '') + '\nA: ' + String(h.a || '')).join('\n\n')
    : '';

  const userPrompt =
    'UNDERLYING DATA:\n' + dataBlock +
    '\n\n' + mode.toUpperCase() + ' OUTPUT ALREADY SHOWN TO THE LEADER:\n' + (priorText || '(none generated yet)') +
    historyBlock +
    '\n\nLEADER\'S FOLLOW-UP QUESTION:\n' + question;

  const raw = callNIMApi_(AI_FOLLOWUP_SYSTEM_PROMPT, userPrompt);
  const answer = raw && raw.trim() ? raw.trim() : 'Sorry — I couldn\'t reach the AI service just now. Please try again in a moment.';
  return json({ success: true, answer: answer });
}

// Thin wrapper around NVIDIA NIM's OpenAI-compatible chat completions
// endpoint. Model name is a Script Property (not hardcoded) because NIM
// models can be deprecated with only a few days' notice — swapping the
// property is a one-line fix, no redeploy of logic needed.
// Returns '' on any failure (missing key, non-200, malformed response) —
// callers decide how to surface that (queryAIInChat falls back to a plain
// "couldn't reach the AI service" message rather than erroring the request).
function callNIMApi_(systemPrompt, userPrompt) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('NIM_API_KEY');
  if (!apiKey) {
    Logger.log('callNIMApi_ error: NIM_API_KEY Script Property is not set.');
    return '';
  }
  const model = props.getProperty('NIM_MODEL_CHAT') || 'z-ai/glm-5.2';
  const baseUrl = (props.getProperty('NIM_BASE_URL') || 'https://integrate.api.nvidia.com/v1').replace(/\/+$/, '');

  const payload = {
    model: model,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user',   content: userPrompt   }
    ],
    max_tokens: 600,
    temperature: 0.3 // low — consistent, not creative; this is ops reporting, not prose
  };

  try {
    const res = UrlFetchApp.fetch(
      baseUrl + '/chat/completions',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + apiKey },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );
    const code = res.getResponseCode();
    if (code !== 200) {
      Logger.log('callNIMApi_ HTTP ' + code + ': ' + res.getContentText().substring(0, 500));
      return '';
    }
    const data = JSON.parse(res.getContentText());
    return (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
  } catch (e) {
    Logger.log('callNIMApi_ error: ' + e.message);
    return '';
  }
}

// Same shape as callNIMApi_ but uses NIM_MODEL_REPORT (a deeper-reasoning
// model, per the handoff — currently MiniMax M2.7) and a higher max_tokens
// budget, since a 400-word 4-section report needs more headroom than a
// ~150-word chat answer. Confirm the exact model string on
// build.nvidia.com before relying on the fallback default below.
function callNIMApiReport_(systemPrompt, userPrompt) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('NIM_API_KEY');
  if (!apiKey) {
    Logger.log('callNIMApiReport_ error: NIM_API_KEY Script Property is not set.');
    return '';
  }
  const model = props.getProperty('NIM_MODEL_REPORT') || 'z-ai/glm-5.2';
  const baseUrl = (props.getProperty('NIM_BASE_URL') || 'https://integrate.api.nvidia.com/v1').replace(/\/+$/, '');

  const payload = {
    model: model,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user',   content: userPrompt   }
    ],
    max_tokens: 900,
    temperature: 0.3
  };

  try {
    const res = UrlFetchApp.fetch(
      baseUrl + '/chat/completions',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + apiKey },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );
    const code = res.getResponseCode();
    if (code !== 200) {
      Logger.log('callNIMApiReport_ HTTP ' + code + ': ' + res.getContentText().substring(0, 500));
      return '';
    }
    const data = JSON.parse(res.getContentText());
    return (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
  } catch (e) {
    Logger.log('callNIMApiReport_ error: ' + e.message);
    return '';
  }
}

// ─────────────────────────────────────
//  AI SETTINGS — Settings modal (bottom-left, leader-only)
//  Lets a leader set NIM_BASE_URL / NIM_API_KEY / NIM_MODEL_CHAT from the
//  UI instead of the Apps Script editor's Script Properties page. All
//  three endpoints below are leader-gated via resolveIdentity(); the API
//  key is never echoed back to the client — getAISettings only reports
//  whether one is currently saved.
// ─────────────────────────────────────

// GET — populates the modal when it's opened.
function getAISettings(token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'Leader only.' });

  const props = PropertiesService.getScriptProperties();
  return json({
    success: true,
    baseUrl: props.getProperty('NIM_BASE_URL') || 'https://integrate.api.nvidia.com/v1',
    model: props.getProperty('NIM_MODEL_CHAT') || 'z-ai/glm-5.2',
    modelReport: props.getProperty('NIM_MODEL_REPORT') || 'z-ai/glm-5.2',
    hasKey: !!props.getProperty('NIM_API_KEY'),
    // Meeting-notetaker ingest secret (see ingestMeetingSummary) — reported
    // the same way as hasKey: never echo the value, just whether one's set,
    // so a leader can tell the pipeline is configured without ever opening
    // the Apps Script editor.
    hasIngestSecret: !!props.getProperty('MEETING_INGEST_SECRET')
  });
}

// POST — the modal's "Save" button. apiKey is optional: blank keeps
// whatever key is already stored, so changing just the model or base URL
// doesn't require re-pasting the key.
function saveAISettings(body) {
  const identity = resolveIdentity(body.token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'Leader only.' });

  const baseUrl = String(body.baseUrl || '').trim().replace(/\/+$/, '');
  const model   = String(body.model   || '').trim();
  const modelReport = String(body.modelReport || '').trim();
  const apiKey  = String(body.apiKey  || '').trim();
  // Optional, same "blank keeps the saved value" pattern as apiKey — a
  // leader changing the model shouldn't have to re-paste the n8n secret.
  const ingestSecret = String(body.ingestSecret || '').trim();
  if (!baseUrl) return json({ error: 'API Base URL is required.' });
  if (!model)   return json({ error: 'Chat Model is required.' });

  const props = PropertiesService.getScriptProperties();
  props.setProperty('NIM_BASE_URL', baseUrl);
  props.setProperty('NIM_MODEL_CHAT', model);
  if (modelReport) props.setProperty('NIM_MODEL_REPORT', modelReport);
  if (apiKey) props.setProperty('NIM_API_KEY', apiKey);
  if (ingestSecret) props.setProperty('MEETING_INGEST_SECRET', ingestSecret);

  return json({ success: true });
}

// POST — the modal's "Test connection" button. Fires a minimal, cheap
// request against whatever's currently typed in the form — WITHOUT
// saving anything — so a bad key, wrong base URL, or unrecognized model
// string is caught before it's persisted. Blank fields fall back to
// whatever is already saved, so e.g. a new model can be tested against
// the existing key without retyping it.
function testAIConnection(body) {
  const identity = resolveIdentity(body.token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'Leader only.' });

  const props = PropertiesService.getScriptProperties();
  const baseUrl = String(body.baseUrl || props.getProperty('NIM_BASE_URL') || 'https://integrate.api.nvidia.com/v1').trim().replace(/\/+$/, '');
  const model   = String(body.model   || props.getProperty('NIM_MODEL_CHAT') || 'z-ai/glm-5.2').trim();
  const apiKey  = String(body.apiKey  || props.getProperty('NIM_API_KEY') || '').trim();

  if (!apiKey)  return json({ success: false, error: 'No API key provided or saved.' });
  if (!baseUrl) return json({ success: false, error: 'No API Base URL provided.' });

  try {
    const res = UrlFetchApp.fetch(baseUrl + '/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify({
        model: model,
        messages: [{ role: 'user', content: 'Reply with exactly: OK' }],
        max_tokens: 5,
        temperature: 0
      }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 200) {
      return json({ success: false, error: 'HTTP ' + code + ': ' + res.getContentText().substring(0, 300) });
    }
    const data = JSON.parse(res.getContentText());
    const reply = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
    return json({ success: true, model: model, reply: reply.trim().substring(0, 80) });
  } catch (e) {
    return json({ success: false, error: e.message });
  }
}

// ═════════════════════════════════════
//  STATE SECURITY REPORT (COS tab)
//  Emails the state's admin a couple of ready-made, scoped Google search
//  links (last hour / last 24h) covering security & public-safety news for
//  their state, pulled from trusted NG gov/news sources. No scraping or
//  external API — just a well-formed search URL, same as the team was
//  already building by hand.
//  NOTE: keep SECURITY_REPORT_KEYWORDS/SOURCES in sync with the client-side
//  copy in index.html (securityReportUrl) if either list is edited.
// ═════════════════════════════════════
const SECURITY_REPORT_KEYWORDS = '(security OR crime OR kidnapping OR robbery OR bandit* OR terrorism OR cultism OR protest OR riot OR "communal clash" OR accident OR "road accident" OR fire OR flood OR emergency OR "building collapse" OR explosion OR "missing person" OR "wanted person" OR "traffic advisory" OR "road closure" OR police OR FRSC OR NEMA OR NSCDC OR DSS OR "public safety" OR outbreak OR cholera OR Lassa OR "weather warning")';
const SECURITY_REPORT_SOURCES  = '(site:police.gov.ng OR site:frsc.gov.ng OR site:nema.gov.ng OR site:nscdc.gov.ng OR site:gov.ng OR site:channels.tv OR site:thecable.ng OR site:premiumtimesng.com OR site:punchng.com OR site:vanguardngr.com OR site:guardian.ng OR site:dailytrust.com OR site:leadership.ng OR site:tribuneonlineng.com)';
function buildSecurityReportUrl_(state, tbs) {
  const q = '("' + state + '") ' + SECURITY_REPORT_KEYWORDS + ' ' + SECURITY_REPORT_SOURCES;
  return 'https://www.google.com/search?tbs=' + tbs + '&q=' + encodeURIComponent(q);
}
function sendStateSecurityReport(state, token) {
  if (!state) return json({ error: 'No state specified.' });
  const rows = sheetToObjects('2. COS');
  const row = rows.find(r => String(r['State']).trim().toLowerCase() === String(state).trim().toLowerCase());
  if (!row) return json({ error: 'State "' + state + '" not found in COS Database.' });

  const identity = resolveIdentity(token);
  const to = row['Admin Email'] || (identity ? identity.email : '') || '';
  if (!to) return json({ error: 'No admin email on file for ' + state + ', and no email on your own account either.' });

  const hourUrl = buildSecurityReportUrl_(state, 'qdr:h');
  const dayUrl  = buildSecurityReportUrl_(state, 'qdr:d');
  const subject = '[ABION] Regional security scan — ' + state;
  const body =
    `Hi ${row['Admin Name'] || 'there'},\n\n` +
    `Here are live-scoped search links for security/public-safety news in ${state}, pulled from trusted NG gov and news sources:\n\n` +
    `Last hour:\n${hourUrl}\n\n` +
    `Last 24 hours:\n${dayUrl}\n\n` +
    `These open Google's live results directly — nothing is cached or summarized, so what you see is current at the moment you click.\n\n` +
    `— NEXUS Automation · ABION Industries`;
  try {
    const sent = sendViaResend(to, subject, null, body, { replyTo: FOUNDER_EMAIL });
    if (!sent) return json({ error: 'Email send failed — check RESEND_API_KEY / logs.' });
    return json({ success: true, sentTo: to });
  } catch (e) {
    return json({ error: 'Email send error: ' + e.message });
  }
}

// ═════════════════════════════════════
//  MEETINGS
// ═════════════════════════════════════

function appendMeeting(data) {
  getOrCreateSheet('7. Meetings', TAB_HEADERS['7. Meetings']);
  data['Meeting_ID'] = 'MTG-' + Date.now();
  data['Created_At'] = new Date().toISOString();
  const result = appendRow('7. Meetings', data);

  // Send calendar invite only for virtual meetings
  if (isAutomationEnabled('calendar_sync') && data.Date && data.Title && data.Location_Type !== 'Physical') {
    try {
      const startDate = new Date(data.Date + 'T' + (data.Time || '09:00') + ':00');
      const endDate   = new Date(startDate.getTime() + (parseInt(data.Duration_Mins) || 60) * 60000);
      const event     = CalendarApp.getDefaultCalendar().createEvent(
        '[ABION] ' + data.Title, startDate, endDate,
        { description: data.Agenda || '', guests: data.External_Emails || '' }
      );
      // Log the Meet/Zoom link back to the sheet if Meet was used
      Logger.log('Meeting calendar event created: ' + event.getId());
    } catch(e) {
      Logger.log('Calendar invite error: ' + e.message);
    }
  }
  return result;
}

// ═════════════════════════════════════
//  MEETING-NOTETAKER INTEGRATION (Fireflies → n8n → review queue)
// ═════════════════════════════════════
// Fireflies already does its own NLP/date-parsing on the transcript and
// ═════════════════════════════════════
//  POLICY ENGINE — risk-tiered actions (NEXUS-vision.md §7)
// ═════════════════════════════════════
// Every AI-initiated action in NEXUS carries an implicit risk posture today
// (does it write automatically, or does it wait for a human?) — this is
// the first pass at making that explicit rather than leaving it to
// rediscover from scattered per-function comments. Levels match the vision
// doc's table exactly:
//
//   0 — Autonomous          Runs and returns, no human in the loop, no
//                            write to anything a human reviews. Covers the
//                            read-only paths: AI Analyst report/attention/
//                            predictions generation, queryAIInChat,
//                            queryAIAnalystFollowup.
//   1 — Low-risk execution  Writes automatically, but to a low-stakes,
//                            reversible location, logged. Nothing in NEXUS
//                            currently qualifies for this level — every
//                            AI-initiated write today goes through Level 2
//                            instead (see below). That's a deliberately
//                            more conservative posture than this table
//                            technically requires for "task suggestions";
//                            worth loosening only once there's a track
//                            record to justify it, not speculatively ahead
//                            of one.
//   2 — Human approval      Writes to '16. Pending AI Tasks', never
//                            directly to '1. Master Tasks'. Every
//                            AI-sourced task-creation path in NEXUS is
//                            currently at this level — see
//                            AI_ACTION_POLICY below for the registry of
//                            which Source values that covers.
//   3 — Executive approval  Not yet applicable — nothing in NEXUS
//                            currently touches money, contracts, hiring,
//                            or external comms autonomously.
//
// AI_ACTION_POLICY is the registry proper: one entry per Source value
// written to '16. Pending AI Tasks', so the Review Queue and any future
// policy tooling can answer "what produced this row, and what gate did it
// pass through" by lookup instead of re-deriving it from memory. When a
// fourth AI-sourced pathway is added, it should get an entry here as part
// of that build, not as an afterthought.
const AI_ACTION_POLICY = {
  'Fireflies': {
    level: 2,
    label: 'Human approval',
    description: 'Meeting action items extracted by ingestMeetingSummary from a Fireflies transcript via n8n. Never reaches Master Tasks without a leader approving the specific row.',
  },
  'Goal Decomposition': {
    level: 2,
    label: 'Human approval',
    description: 'A leader-described goal broken into 4-10 tasks by decomposeGoal / runGoalDecomposition_. Each task still needs individual approval even though the goal itself was leader-initiated — the leader approved the OBJECTIVE, not the specific task breakdown.',
  },
  'AI Analyst': {
    level: 2,
    label: 'Human approval',
    description: 'A single task extracted from a leader message in the AI Analyst chat via createTaskFromChat (the /task prefix). Same gate as the other two sources.',
  },
};

// Not enforcement — nothing currently blocks a write for using an
// unregistered Source value. Just makes an undocumented one loud in the
// logs (so it gets noticed and given a real policy entry) rather than
// silently slipping through with no risk-level metadata attached. Call
// this right before any appendRow into '16. Pending AI Tasks' (or any
// future policy-gated tab) with the Source value being written.
function assertKnownPolicySource_(source) {
  if (!AI_ACTION_POLICY[source]) {
    Logger.log('Policy Engine warning: unregistered AI action Source "' + source + '" — add an entry to AI_ACTION_POLICY before shipping this pathway.');
  }
}

// Leader-only read of the registry itself — lets a future admin/settings
// view render the risk-level table without hardcoding it a second time in
// the frontend. Not used by the Review Queue directly (see
// getPendingAITasks, which inlines the same lookup per-row instead so the
// queue doesn't need a second round-trip) — this is for anything that
// wants the registry on its own, e.g. a policy-overview screen later.
function getPolicyRegistry(token) {
  const err = requireLeader(token); if (err) return err;
  return json({ success: true, policy: AI_ACTION_POLICY });
}

// hands back structured action items (description + assignee + due date),
// so none of that happens here. Path: Fireflies -> n8n -> this doPost
// action. Reuses the same "n8n as middleware, Apps Script as source of
// truth" pattern as N8N/pingN8N elsewhere — this is just the inbound
// direction of that same pattern, not a second integration mechanism.
//
// Landing zone is a REVIEW QUEUE, not a direct write — every extracted
// item lands in '16. Pending AI Tasks' first. Nothing from this pipeline
// ever reaches '1. Master Tasks' without a leader explicitly approving it
// via approvePendingAITask below.

// Inbound calls come from n8n, a server, not a logged-in person — there's
// no user token to check. Gated instead by a shared secret (same pattern
// as RESEND_API_KEY: stored in Script Properties, never in source).
// Set it via Project Settings → Script Properties → MEETING_INGEST_SECRET,
// then give n8n's HTTP node the same value.
function ingestMeetingSummary(body) {
  const expected = PropertiesService.getScriptProperties().getProperty('MEETING_INGEST_SECRET');
  if (!expected || body.secret !== expected) return json({ error: 'Unauthorized.' });

  // Was: always-on the moment the secret was configured, with no
  // leader-facing way to pause AI task ingestion short of deleting the
  // Script Property. Now matches every other automation — a real toggle,
  // visible and switchable from the Automation Controls tab.
  if (!isAutomationEnabled('meeting_notetaker')) {
    return json({ error: 'Meeting notetaker automation is currently disabled.' });
  }

  const items = Array.isArray(body.actionItems) ? body.actionItems : [];
  if (!items.length) return json({ error: 'No action items provided.' });

  const sheet   = getOrCreateSheet('16. Pending AI Tasks', TAB_HEADERS['16. Pending AI Tasks']);
  const headers = TAB_HEADERS['16. Pending AI Tasks'];

  // Exact match only, per spec — no fuzzy-guessing an assignee. A nickname,
  // typo, or mishearing just leaves it unassigned for a leader to sort out.
  const liveTeam    = getAllTeamNames(); // fetched once, reused for every item below
  const rosterCanon = liveTeam.map(n => canonicalizeName(n).trim().toUpperCase());

  let created = 0;
  items.forEach(item => {
    const desc = String(item.description || '').trim();
    if (!desc) return;
    const rawAssignee = String(item.assignee || '').trim();
    const canon        = rawAssignee ? canonicalizeName(rawAssignee).trim().toUpperCase() : '';
    const matchIdx      = canon ? rosterCanon.indexOf(canon) : -1;
    const matchedAssignee = matchIdx >= 0 ? liveTeam[matchIdx] : '';

    const row = {
      'ID'                 : 'PAT-' + Date.now() + '-' + Math.random().toString(36).substr(2,4),
      'Meeting_ID'         : body.meetingId || '',
      'Task Description'   : desc,
      'Suggested Assignee' : rawAssignee,
      'Matched Assignee'   : matchedAssignee,
      'Suggested Deadline' : item.dueDate || '',
      'Priority'           : item.priority || 'MEDIUM',
      'Status'             : 'Pending',
      'Source'             : 'Fireflies',
      'Created_At'         : new Date().toISOString(),
    };
    assertKnownPolicySource_(row['Source']);
    sheet.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
    created++;
  });

  return json({ success: true, created });
}

// Leader-facing review queue — pending items only, oldest first.
function getPendingAITasks(token) {
  const err = requireLeader(token); if (err) return err;
  const items = sheetToObjects('16. Pending AI Tasks').filter(r => (r.Status || 'Pending') === 'Pending');
  // Fold in the policy metadata for each row's Source right here, so the
  // Review Queue can render a risk-level tag without a second call to
  // getPolicyRegistry -- the registry is tiny and rarely changes, so this
  // is cheaper than making the frontend join the two itself.
  items.forEach(function(it) { it._policy = AI_ACTION_POLICY[it.Source] || null; });
  return json({ success: true, items: items, policy: AI_ACTION_POLICY });
}


// NEW (v3.2) — goal decomposition. POST: leader describes a goal in plain
// English; NIM decomposes it into multiple proposed tasks, all landing in
// '16. Pending AI Tasks' tagged with a shared Goal_ID so the review queue
// can group them. A '19. Goals' row is created up front (Status: Drafting)
// so the goal exists as a trackable object even before anything is approved.
function decomposeGoal(body) {
  const identity = resolveIdentity(body.token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'This feature is available to leaders only.' });

  const goalText = String(body.goal || '').trim();
  if (!goalText) return json({ error: 'No goal text provided.' });

  const goalId = 'GOAL-' + Date.now();
  const sheet = getOrCreateSheet('19. Goals', TAB_HEADERS['19. Goals']);
  sheet.appendRow([goalId, goalText, identity.name, new Date().toISOString(), 'Drafting', '', 0, '']);

  const result = runGoalDecomposition_(goalId, goalText);
  if (!result.createdIds.length) {
    return json({ error: 'The AI could not produce a task breakdown for that goal — check NIM_API_KEY / model settings and try again, or rephrase the goal.' });
  }
  return json({ success: true, goalId: goalId, tasksCreated: result.createdIds.length, reasoning: result.reasoning });
}

// Shared by decomposeGoal() and regenerateGoalTasks(): calls NIM, parses
// the JSON array response, and writes each proposed task into
// '16. Pending AI Tasks' tagged with this Goal_ID. Reuses the exact same
// exact-match roster logic as ingestMeetingSummary() above — no fuzzy
// assignee guessing here either.
function runGoalDecomposition_(goalId, goalText) {
  const liveTeam = getAllTeamNames();
  const rosterCanon = liveTeam.map(n => canonicalizeName(n).trim().toUpperCase());
  const rosterBlock = liveTeam.length ? liveTeam.join(', ') : '(no team members on roster yet)';

  const userPrompt = 'GOAL:\n' + goalText + '\n\nTEAM ROSTER (match names EXACTLY or leave blank):\n' + rosterBlock;
  const raw = callNIMApiReport_(AI_GOAL_DECOMPOSE_SYSTEM_PROMPT, userPrompt);

  // Shape is now {reasoning, tasks:[...]} — but fall back to accepting a
  // bare array (old shape / a model that ignores the object wrapper) so a
  // prompt-following slip doesn't just silently produce zero tasks.
  let parsed = [];
  let reasoning = '';
  try {
    const cleaned = String(raw || '').trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
    const obj = JSON.parse(cleaned);
    if (Array.isArray(obj)) {
      parsed = obj;
    } else if (obj && Array.isArray(obj.tasks)) {
      parsed = obj.tasks;
      reasoning = String(obj.reasoning || '').trim();
    }
  } catch (e) {
    Logger.log('runGoalDecomposition_ parse error: ' + e.message + ' | raw: ' + String(raw).substring(0, 300));
    parsed = [];
  }

  const sheet = getOrCreateSheet('16. Pending AI Tasks', TAB_HEADERS['16. Pending AI Tasks']);
  const headers = TAB_HEADERS['16. Pending AI Tasks'];
  const createdIds = [];

  // Exact-match a raw AI-suggested name against the live roster — shared
  // by suggestedAssignee/suggestedSupervisor/suggestedSupportPerson below,
  // all three of which are people-fields following the same "never guess"
  // rule from the prompt.
  const matchRoster = (raw) => {
    const rawName = String(raw || '').trim();
    if (!rawName) return { raw: '', matched: '' };
    const canon = canonicalizeName(rawName).trim().toUpperCase();
    const idx = rosterCanon.indexOf(canon);
    return { raw: rawName, matched: idx >= 0 ? liveTeam[idx] : '' };
  };

  parsed.slice(0, 12).forEach((item, idx) => { // hard cap of 12, even if the model ignores the 4-10 guidance
    const desc = String(item.description || '').trim();
    if (!desc) return;
    const assignee   = matchRoster(item.suggestedAssignee);
    const supervisor = matchRoster(item.suggestedSupervisor);
    const support    = matchRoster(item.suggestedSupportPerson);
    const dept = DEPARTMENT_OPTIONS.indexOf(String(item.suggestedDepartment || '').trim()) >= 0
      ? String(item.suggestedDepartment).trim() : '';
    const id = 'PAT-' + Date.now() + '-' + idx + '-' + Math.random().toString(36).substr(2, 4);

    const row = {
      'ID': id,
      'Meeting_ID': '',
      'Task Description': desc,
      'Suggested Assignee': assignee.raw,
      'Matched Assignee': assignee.matched,
      'Suggested Deadline': item.suggestedDeadline || '',
      'Priority': item.priority || 'MEDIUM',
      'Status': 'Pending',
      'Source': 'Goal Decomposition',
      'Created_At': new Date().toISOString(),
      'Reviewed_By': '',
      'Reviewed_At': '',
      // Repurposing this existing column (originally "meeting summary text")
      // to hold the goal's own description, so the review-queue UI can show
      // a group header without a second round-trip to '19. Goals'.
      'Meeting_Summary': goalText,
      'Goal_ID': goalId,
      'Depends_On_Index': (item.dependsOnIndex === null || item.dependsOnIndex === undefined || item.dependsOnIndex === '') ? '' : item.dependsOnIndex,
      'Suggested Supervisor': supervisor.raw,
      'Matched Supervisor': supervisor.matched,
      'Suggested Support Person': support.raw,
      'Matched Support Person': support.matched,
      'Suggested Department': dept,
    };
    assertKnownPolicySource_(row['Source']);
    sheet.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
    createdIds.push(id);
  });

  // Stamp the goal's PendingCount so getGoals() doesn't have to re-scan
  // Pending AI Tasks just to show a number in the list.
  setGoalPendingCount_(goalId, createdIds.length);
  if (reasoning) setGoalReasoning_(goalId, reasoning);
  return { createdIds: createdIds, reasoning: reasoning };
}

function setGoalPendingCount_(goalId, count) {
  const sheet = SS.getSheetByName('19. Goals');
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idIdx = headers.indexOf('ID');
  const pendIdx = headers.indexOf('PendingCount');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(goalId)) {
      sheet.getRange(i + 1, pendIdx + 1).setValue(count);
      return;
    }
  }
}

// Stamps (or overwrites, on regenerate) the goal's stored Reasoning text —
// see AI_GOAL_DECOMPOSE_SYSTEM_PROMPT's "reasoning" field. Adds the column
// live if this sheet predates it, same pattern as updateTeamPermission's
// live-column-add for Trusted_* grants, so an already-deployed sheet
// doesn't need a manual migration for this to start working.
function setGoalReasoning_(goalId, reasoning) {
  const sheet = SS.getSheetByName('19. Goals');
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  let headers = values[0];
  let idx = headers.indexOf('Reasoning');
  if (idx === -1) {
    idx = headers.length;
    sheet.getRange(1, idx + 1).setValue('Reasoning');
    headers = headers.concat(['Reasoning']);
  }
  const idIdx = headers.indexOf('ID');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(goalId)) {
      sheet.getRange(i + 1, idx + 1).setValue(reasoning);
      return;
    }
  }
}

// GET — every goal with a computed rollup (done/total task counts) and its
// child tasks' current status, so the frontend renders progress without a
// second call per goal.
function getGoals(token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'Leader only.' });

  const goals = sheetToObjects('19. Goals');
  const tasksById = buildTasksById_();

  const result = goals.map(g => {
    const taskIds = String(g.TaskIDs || '').split(',').map(s => s.trim()).filter(Boolean);
    const children = taskIds.map(id => tasksById[id]).filter(Boolean).map(t => ({
      ID: t['ID'], 'Task Description': t['Task Description'], Status: t['Status'],
    }));
    const done = children.filter(c => c.Status === 'Done' || c.Status === 'Cancelled').length;
    return {
      ID: g.ID, Description: g.Description, CreatedBy: g.CreatedBy, CreatedAt: g.CreatedAt,
      Status: g.Status, PendingCount: Number(g.PendingCount) || 0, Notes: g.Notes,
      Reasoning: g.Reasoning || '',
      total: children.length, done: done, children: children,
    };
  }).sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));

  return json({ success: true, goals: result });
}

function getGoalDetail(id, token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!identity.isLeader) return json({ error: 'Leader only.' });

  const goal = sheetToObjects('19. Goals').find(g => String(g.ID) === String(id));
  if (!goal) return json({ error: 'Goal not found.' });

  const tasksById = buildTasksById_();
  const taskIds = String(goal.TaskIDs || '').split(',').map(s => s.trim()).filter(Boolean);
  const children = taskIds.map(tid => tasksById[tid]).filter(Boolean);
  const pending = sheetToObjects('16. Pending AI Tasks').filter(p => p.Goal_ID === id && (p.Status || 'Pending') === 'Pending');

  return json({ success: true, goal: goal, children: children, pending: pending });
}

// Leader-only manual close. Does NOT touch any Master Tasks rows — those
// keep whatever status they're in. This only marks the goal itself closed
// so it stops showing as Active/Drafting in the Goals list.
function archiveGoal(id, token) {
  const callerName = getLeaderIdentity(token);
  if (!callerName) return json({ error: 'Leader login required.' });
  const sheet = SS.getSheetByName('19. Goals');
  if (!sheet) return json({ error: 'Goals sheet not found.' });
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idIdx = headers.indexOf('ID');
  const statusIdx = headers.indexOf('Status');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(id)) {
      sheet.getRange(i + 1, statusIdx + 1).setValue('Archived');
      return json({ success: true });
    }
  }
  return json({ error: 'Goal not found: ' + id });
}

// Leader-only. Discards any still-Pending review-queue items tied to this
// goal (already-approved/rejected ones are left alone — this never touches
// real Master Tasks rows) and re-runs decomposition from the goal's
// original stored description. Simple "start over" semantics, not
// iterative editing.
function regenerateGoalTasks(id, token) {
  const callerName = getLeaderIdentity(token);
  if (!callerName) return json({ error: 'Leader login required.' });

  const goal = sheetToObjects('19. Goals').find(g => String(g.ID) === String(id));
  if (!goal) return json({ error: 'Goal not found.' });

  const sheet = SS.getSheetByName('16. Pending AI Tasks');
  if (sheet) {
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const goalIdx = headers.indexOf('Goal_ID');
    const statusIdx = headers.indexOf('Status');
    // Delete bottom-up so row indices don't shift under us mid-loop.
    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][goalIdx] === id && (values[i][statusIdx] || 'Pending') === 'Pending') {
        sheet.deleteRow(i + 1);
      }
    }
  }

  const result = runGoalDecomposition_(id, goal.Description);
  return json({ success: true, tasksCreated: result.createdIds.length, reasoning: result.reasoning });
}


// Turns a (possibly leader-edited) pending item into a real Master Tasks
// row, then marks the pending item Approved. `data` carries whatever the
// leader confirmed in the review UI — description/assignee/deadline/
// priority — which may differ from what Fireflies originally suggested.
// v3.2 — now also links a newly-created Master Tasks row back onto its
// parent goal (if it has one, i.e. it came from decomposeGoal rather than
// Fireflies) and resolves Depends_On_Index into a real "Depends On" value
// once the target task has itself already been approved.
function approvePendingAITask(id, data, token) {
  const err = requireLeader(token); if (err) return err;
  const identity = resolveIdentity(token);

  const sheet   = SS.getSheetByName('16. Pending AI Tasks');
  if (!sheet) return json({ error: "'16. Pending AI Tasks' not found." });
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idIdx   = headers.indexOf('ID');
  const values  = sheet.getDataRange().getValues();
  let rowNum = -1;
  for (let i = 1; i < values.length; i++) { if (String(values[i][idIdx]) === String(id)) { rowNum = i + 1; break; } }
  if (rowNum === -1) return json({ error: 'Pending item not found: ' + id });

  if (!data || !data.taskDescription || !data.assignedTo) {
    return json({ error: 'Task description and assignee are required to approve.' });
  }

  const newTaskId = '1-' + Date.now();
  const newTaskRow = {
    'ID'               : newTaskId,
    'Task Description' : data.taskDescription,
    'Assigned To'      : data.assignedTo,
    'Deadline'         : data.deadline || '',
    'Priority'         : data.priority || 'MEDIUM',
    'Status'           : 'Not Started',
  };
  // v3.3 — carries whatever the leader confirmed (or left as the AI's
  // suggestion) for Supervisor / Support Person / Department through to
  // the real Master Tasks row. All optional — appendRow only writes
  // columns present here, so an empty string just leaves the cell blank.
  if (data.supervisor)    newTaskRow['Supervisor']      = data.supervisor;
  if (data.supportPerson) newTaskRow['Support Person']  = data.supportPerson;
  if (data.department)    newTaskRow['Business / Dept'] = data.department;

  // Resolve Depends_On_Index (goal-decomposition rows only) into a real
  // "Depends On" value, but ONLY if the task it depends on has ALREADY been
  // approved into Master Tasks (i.e. its ID is already on the goal's
  // TaskIDs list). If the dependency hasn't been approved yet, this
  // silently leaves Depends On blank rather than guessing — approving out
  // of order is allowed, it just doesn't auto-wire a dependency that
  // doesn't exist as a real task yet.
  const goalIdCol = headers.indexOf('Goal_ID');
  const depIdxCol = headers.indexOf('Depends_On_Index');
  const goalIdForThisItem = goalIdCol >= 0 ? values[rowNum - 1][goalIdCol] : '';
  const dependsOnIndex = depIdxCol >= 0 ? values[rowNum - 1][depIdxCol] : '';

  if (goalIdForThisItem && dependsOnIndex !== '' && dependsOnIndex !== null && dependsOnIndex !== undefined) {
    const siblingRows = sheetToObjects('16. Pending AI Tasks').filter(p => p.Goal_ID === goalIdForThisItem);
    // Reconstruct the same array order NIM produced by Created_At — all
    // sibling rows for one goal are written in one appendRow loop, so
    // Created_At order matches dependsOnIndex order exactly.
    siblingRows.sort((a, b) => new Date(a.Created_At) - new Date(b.Created_At));
    const depSibling = siblingRows[Number(dependsOnIndex)];
    if (depSibling && depSibling.Status === 'Approved') {
      // The dependency's own approval would have appended its new Master
      // Tasks ID onto the goal — find it via the goal's TaskIDs, matching
      // this sibling's description (best-effort; description is the only
      // link we have back from a pending row to its resulting task ID).
      const goalRow = sheetToObjects('19. Goals').find(g => g.ID === goalIdForThisItem);
      if (goalRow) {
        const candidateIds = String(goalRow.TaskIDs || '').split(',').map(s => s.trim()).filter(Boolean);
        const tasksById = buildTasksById_();
        const match = candidateIds.map(tid => tasksById[tid]).find(t => t && t['Task Description'] === depSibling['Task Description']);
        if (match) newTaskRow['Depends On'] = match['ID'];
      }
    }
  }

  appendRow('1. Master Tasks', newTaskRow);

  const statusIdx = headers.indexOf('Status');
  const revByIdx  = headers.indexOf('Reviewed_By');
  const revAtIdx  = headers.indexOf('Reviewed_At');
  if (statusIdx >= 0) sheet.getRange(rowNum, statusIdx+1).setValue('Approved');
  if (revByIdx  >= 0) sheet.getRange(rowNum, revByIdx+1).setValue((identity && identity.name) || '');
  if (revAtIdx  >= 0) sheet.getRange(rowNum, revAtIdx+1).setValue(new Date().toISOString());

  // Link this task onto its parent goal, if it has one.
  if (goalIdForThisItem) {
    appendTaskToGoal_(goalIdForThisItem, newTaskId);
  }

  return json({ success: true, taskId: newTaskId });
}

// Appends a newly-approved task's ID onto its goal's TaskIDs, flips the
// goal from Drafting to Active on its first approval, and refreshes
// PendingCount. Called from approvePendingAITask above.
function appendTaskToGoal_(goalId, taskId) {
  const sheet = SS.getSheetByName('19. Goals');
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idIdx = headers.indexOf('ID');
  const taskIdsIdx = headers.indexOf('TaskIDs');
  const statusIdx = headers.indexOf('Status');
  const pendIdx = headers.indexOf('PendingCount');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(goalId)) {
      const existing = String(values[i][taskIdsIdx] || '').split(',').map(s => s.trim()).filter(Boolean);
      if (existing.indexOf(taskId) === -1) existing.push(taskId);
      sheet.getRange(i + 1, taskIdsIdx + 1).setValue(existing.join(','));
      if (String(values[i][statusIdx]) === 'Drafting') {
        sheet.getRange(i + 1, statusIdx + 1).setValue('Active');
      }
      const remainingPending = sheetToObjects('16. Pending AI Tasks').filter(p => p.Goal_ID === goalId && (p.Status || 'Pending') === 'Pending').length;
      sheet.getRange(i + 1, pendIdx + 1).setValue(remainingPending);
      return;
    }
  }
}

// Recomputes and writes a goal's completion status. A goal is Completed
// once every one of its linked tasks is Done or Cancelled AND at least one
// of them is actually Done (an all-Cancelled goal doesn't silently read as
// "completed" — that's closer to Archived, left for a leader to decide).
// Called from updateRow() and updateTaskStatus() on every Master Tasks
// status write — cheap no-op unless the task actually belongs to a goal.
function recomputeGoalCompletion_(taskId) {
  const sheet = SS.getSheetByName('19. Goals');
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const taskIdsIdx = headers.indexOf('TaskIDs');
  const statusIdx = headers.indexOf('Status');
  const tasksById = buildTasksById_();
  for (let i = 1; i < values.length; i++) {
    const ids = String(values[i][taskIdsIdx] || '').split(',').map(s => s.trim()).filter(Boolean);
    if (ids.indexOf(String(taskId)) === -1) continue;
    if (values[i][statusIdx] === 'Archived') return; // manually closed — leave it alone
    const linked = ids.map(tid => tasksById[tid]).filter(Boolean);
    const allClosed = linked.length > 0 && linked.every(t => t.Status === 'Done' || t.Status === 'Cancelled');
    const anyDone = linked.some(t => t.Status === 'Done');
    if (allClosed && anyDone) {
      sheet.getRange(i + 1, statusIdx + 1).setValue('Completed');
    }
    return;
  }
}

function rejectPendingAITask(id, token) {
  const err = requireLeader(token); if (err) return err;
  const identity = resolveIdentity(token);

  const sheet   = SS.getSheetByName('16. Pending AI Tasks');
  if (!sheet) return json({ error: "'16. Pending AI Tasks' not found." });
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idIdx   = headers.indexOf('ID');
  const values  = sheet.getDataRange().getValues();
  let rowNum = -1;
  for (let i = 1; i < values.length; i++) { if (String(values[i][idIdx]) === String(id)) { rowNum = i + 1; break; } }
  if (rowNum === -1) return json({ error: 'Pending item not found: ' + id });

  const statusIdx = headers.indexOf('Status');
  const revByIdx  = headers.indexOf('Reviewed_By');
  const revAtIdx  = headers.indexOf('Reviewed_At');
  if (statusIdx >= 0) sheet.getRange(rowNum, statusIdx+1).setValue('Rejected');
  if (revByIdx  >= 0) sheet.getRange(rowNum, revByIdx+1).setValue((identity && identity.name) || '');
  if (revAtIdx  >= 0) sheet.getRange(rowNum, revAtIdx+1).setValue(new Date().toISOString());

  return json({ success: true });
}

function resendAssignmentEmail(taskId, token) {
  const callerName = getLeaderIdentity(token);
  if (!callerName) return json({ error: 'Leader login required.' });
  const tasks = sheetToObjects('1. Master Tasks');
  const task  = tasks.find(t => String(t['ID']) === String(taskId));
  if (!task)  return json({ error: 'Task not found: ' + taskId });
  if (!task['Assigned_To_Email']) return json({ error: 'No email on this task — add Assignee Email first.' });
  const delivered = sendAssignmentEmail(task);
  if (!delivered) {
    return json({ error: 'Resend did not confirm delivery — check the RESEND_API_KEY / Resend dashboard and try again. (Calendar entry, if any, was still updated.)' });
  }
  return json({ success: true });
}

// ═════════════════════════════════════
//  DRIVE SHARE LINKS (no attachments, no email quota used)
// ═════════════════════════════════════

function getShareLink(input) {
  const fileId = extractDriveFileId(input);
  if (!fileId) {
    // Not a Drive/Docs/Sheets/Slides link — there's no sharing permission
    // for this app to change, so just pass it through as-is rather than
    // erroring. (Google Docs/Sheets/Slides URLs already match the same
    // "/d/FILE_ID/" pattern as Drive files, so those still get fixed up below.)
    const url = String(input || '').trim();
    if (!url) return json({ error: 'No link or ID provided.' });
    return json({ success: true, url, external: true });
  }
  try {
    const file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return json({ success: true, url: file.getUrl(), name: file.getName() });
  } catch(e) {
    return json({ error: "Couldn't access that file — check the link, or make sure it's shared with the NEXUS Google account. (" + e.message + ")" });
  }
}

function extractDriveFileId(input) {
  if (!input) return null;
  input = String(input).trim();
  if (/^[a-zA-Z0-9_-]{20,}$/.test(input)) return input; // already a bare file ID
  const patterns = [/\/d\/([a-zA-Z0-9_-]{20,})/, /[?&]id=([a-zA-Z0-9_-]{20,})/, /\/folders\/([a-zA-Z0-9_-]{20,})/];
  for (const p of patterns) { const m = input.match(p); if (m) return m[1]; }
  return null;
}

// ═════════════════════════════════════
//  CHAT — team channel + per-task discussion threads
// ═════════════════════════════════════
// Channel is either the literal string 'team' (general team chat) or 'task'
// (a per-task discussion thread, with the task ID carried separately in TaskID).

function sendChatMessage(data, token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!data || !data.Message || !String(data.Message).trim()) return json({ error: 'Message cannot be empty.' });
  const channel = data.Channel === 'task' ? 'task' : 'team';
  const sheet   = getOrCreateSheet('10. Chat Messages', TAB_HEADERS['10. Chat Messages']);
  const id      = 'MSG-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
  const row     = {
    ID: id,
    Channel: channel,
    TaskID: channel === 'task' ? String(data.TaskID || '') : '',
    // Author/AuthorEmail come from the verified session, never the client —
    // this is what stops anyone from posting chat messages as someone else.
    Author: identity.name,
    AuthorEmail: identity.email || '',
    Message: String(data.Message).trim().substring(0, 2000),
    Timestamp: new Date().toISOString(),
  };
  const headers = TAB_HEADERS['10. Chat Messages'];
  sheet.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
  if (channel === 'task') {
    try { CacheService.getScriptCache().remove(CHAT_THREADS_CACHE_KEY); } catch(ce) {}
  }

  // Fire-and-forget push + in-app notification.
  // 'team' messages really are for everyone, so those still broadcast.
  // 'task' messages should NOT go to the whole team — they should reach
  // the people who actually need them: the task's assignee (who may be
  // the one waiting on clarity from a leader) plus anyone who has already
  // posted in that thread. The title also now says "Task Reply" with the
  // task's own description, not just "Task <ID>", so a notification is
  // never ambiguous about whether it's general chat or a specific task.
  try {
    const preview = row.Message.length > 120 ? row.Message.substring(0, 117) + '…' : row.Message;
    if (channel === 'team') {
      // Broadcast to the whole roster except the author — same audience
      // as before, but routed through notifyInApp so it lands in-app for
      // anyone who wants that, and deep-links straight to General chat
      // instead of the homepage.
      const allNames = sheetToObjects('9. Team Members').map(m => canonicalizeName(String(m.Name || '').trim())).filter(Boolean);
      const targets = allNames.filter(n => n !== canonicalizeName(row.Author));
      notifyInApp(targets, {
        type: 'chat_team',
        title: 'NEXUS · General Chat · ' + row.Author,
        body: preview,
        url: notifUrl_({ chat: 'team' }),
        tag: 'nexus-chat-team',
      });
    } else {
      // Was doing a fresh full-sheet scan of both '1. Master Tasks' and
      // '10. Chat Messages' on every single send just to figure out who to
      // notify — that's what was making task-chat replies feel slow as
      // those sheets grow. Now reads from the same short-lived cache the
      // thread list uses (see getChatThreadsBaseCached), which is usually
      // already warm, and only falls back to a real scan on a cold cache.
      const base = getChatThreadsBaseCached();
      const entry = base[String(row.TaskID)];
      const taskTitle = entry ? String(entry.taskTitle || '').substring(0, 60) : ('Task ' + row.TaskID);

      const recipients = {};
      if (entry && entry.assignedTo) recipients[canonicalizeName(String(entry.assignedTo).trim())] = true;
      (entry ? entry.authors : []).forEach(a => { recipients[canonicalizeName(String(a).trim())] = true; });
      delete recipients[canonicalizeName(String(row.Author).trim())];

      const names = Object.keys(recipients);
      if (names.length) {
        notifyInApp(names, {
          type: 'chat_task',
          title: 'NEXUS · Task Reply · ' + taskTitle,
          body: row.Author + ': ' + preview,
          url: notifUrl_({ task: row.TaskID, chat: 1 }),
          tag: 'nexus-chat-task-' + row.TaskID,
          taskId: row.TaskID,
        });
      }
    }
  } catch(pe) { Logger.log('notifyInApp (chat) error: ' + pe.message); }

  return json({ success: true, id: id });
}

// Previously required no login at all — anyone with the web app URL and a
// channel/taskId could pull chat history. Now: 'team' requires any verified
// session; 'task' additionally requires task-thread access (see
// hasTaskThreadAccess below).
function getChatMessages(token, channel, taskId, sinceId) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });

  channel = channel === 'task' ? 'task' : 'team';

  if (channel === 'task' && !hasTaskThreadAccess(identity, taskId)) {
    return json({ error: 'Access denied.' });
  }

  let rows = sheetToObjects('10. Chat Messages').filter(r => r.Channel === channel);
  if (channel === 'task') rows = rows.filter(r => String(r.TaskID) === String(taskId || ''));
  rows.sort((a, b) => new Date(a.Timestamp) - new Date(b.Timestamp));
  if (sinceId) {
    const idx = rows.findIndex(r => r.ID === sinceId);
    if (idx >= 0) rows = rows.slice(idx + 1);
  }
  return json({ messages: rows.slice(-200) });
}

// Task-thread chat access: a Leader, anyone with a Master-Tasks trust grant
// (isTrustedForTab, reused from RBAC), the task's current Assigned To /
// Support Person, OR anyone who has EVER posted in this specific thread —
// reusing the same "authors" list already built for push notifications
// (see buildChatThreadsBase). Access is ongoing, not frozen: once you've
// posted you keep full access (past and future messages) until a leader
// explicitly removes you, matching how authors already behave for pushes —
// this is deliberately the same access model, not a second one.
function hasTaskThreadAccess(identity, taskId) {
  if (identity.isLeader) return true;
  if (isTrustedForTab(identity.name, '1. Master Tasks')) return true;

  const me = canonicalizeName(identity.name).trim().toUpperCase();

  const task = sheetToObjects('1. Master Tasks').find(t => String(t['ID']) === String(taskId || ''));
  // Was: canonicalized the whole "Assigned To" string then compared with
  // strict equality — worked for a single assignee but never matched either
  // person on a multi-assignee task ("Ayo T. Joshua, Yewande"), and Support
  // Person had the exact same bug (compared as one un-split blob). Both are
  // genuinely multi-value now, so this goes through the same shared,
  // split-then-canonicalize-each check used everywhere else — which also
  // brings Supervisor into thread access for free.
  if (task && isPersonOnTask(me, task)) return true;

  const base  = getChatThreadsBaseCached();
  const entry = base[String(taskId)];
  if (entry) {
    const authors = entry.authors.map(a => canonicalizeName(String(a)).trim().toUpperCase());
    if (authors.indexOf(me) !== -1) return true;
  }

  return false;
}

// Same access model as hasTaskThreadAccess (leader / trusted-for-Master-Tasks
// / assigned-support-supervisor / prior author in that specific thread), but
// takes a pre-fetched tasks-by-ID map + threads base instead of re-reading
// the whole Master Tasks sheet once per task-id. getChatThreads/getChatSummary
// iterate over every task-with-messages in the org, so calling
// hasTaskThreadAccess (which does its own sheetToObjects call) inside that
// loop would mean one full sheet read per thread — this is the same check,
// just batched for list/summary endpoints instead of single-thread fetches.
function threadVisibleTo_(identity, taskId, tasksById, base) {
  if (identity.isLeader) return true;
  if (isTrustedForTab(identity.name, '1. Master Tasks')) return true;
  const me = canonicalizeName(identity.name).trim().toUpperCase();
  const task = tasksById[String(taskId)];
  if (task && isPersonOnTask(me, task)) return true;
  const entry = base[String(taskId)];
  if (entry) {
    const authors = entry.authors.map(a => canonicalizeName(String(a)).trim().toUpperCase());
    if (authors.indexOf(me) !== -1) return true;
  }
  return false;
}
function buildTasksById_() {
  const map = {};
  sheetToObjects('1. Master Tasks').forEach(t => { if (t['ID']) map[String(t['ID'])] = t; });
  return map;
}

// Marks a channel as read for a user (channel = 'team' or a task ID string).
// `name` is now derived from the caller's verified session token — never
// trusted from the client — so nobody can mark someone else's chat as read
// (or, worse, post as them; see sendChatMessage).
function markChatRead(token, channel) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!channel) return json({ error: 'channel required.' });
  const name  = identity.name;
  const sheet = getOrCreateSheet('12. Chat Read State', TAB_HEADERS['12. Chat Read State']);
  const data  = sheet.getDataRange().getValues();
  const now   = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(name).trim() && String(data[i][1]).trim() === String(channel).trim()) {
      sheet.getRange(i + 1, 3).setValue(now);
      return json({ success: true });
    }
  }
  sheet.appendRow([name, channel, now]);
  return json({ success: true });
}

// Unread counts for a given user: team channel + every task thread that has messages.
function getChatSummary(token) {
  const identity = resolveIdentity(token);
  const name = identity ? identity.name : null;
  if (!name) return json({ team: 0, tasks: {}, totalTasks: 0 });
  const reads = sheetToObjects('12. Chat Read State').filter(r => String(r.Name).trim() === String(name).trim());
  const lastReadFor = (channel) => {
    const r = reads.find(x => String(x.Channel).trim() === String(channel).trim());
    return r ? new Date(r.LastReadTimestamp) : new Date(0);
  };
  const all = sheetToObjects('10. Chat Messages');

  const teamMsgs   = all.filter(r => r.Channel === 'team');
  const teamLast   = lastReadFor('team');
  const teamUnread = teamMsgs.filter(r => new Date(r.Timestamp) > teamLast && r.Author !== name).length;

  const taskMsgs = all.filter(r => r.Channel === 'task' && r.TaskID);
  const byTask = {};
  taskMsgs.forEach(r => { (byTask[r.TaskID] = byTask[r.TaskID] || []).push(r); });
  // Skip tasks marked Done — a finished task's chat shouldn't keep
  // contributing to the "you have unread messages" badge. Also skip any
  // task thread this caller doesn't actually have access to (see
  // threadVisibleTo_) — previously this counted EVERY task thread in the
  // org for every logged-in member, which both leaked task info to people
  // not on those tasks and made badge counts show numbers unrelated to the
  // person actually logged in.
  const threadsBase = getChatThreadsBaseCached();
  const tasksById = buildTasksById_();
  const tasks = {};
  let totalTasks = 0;
  Object.keys(byTask).forEach(taskId => {
    const status = threadsBase[taskId] ? threadsBase[taskId].status : '';
    if (status === 'Done') return;
    if (!threadVisibleTo_(identity, taskId, tasksById, threadsBase)) return;
    const last = lastReadFor(taskId);
    const count = byTask[taskId].filter(r => new Date(r.Timestamp) > last && r.Author !== name).length;
    if (count > 0) { tasks[taskId] = count; totalTasks += count; }
  });

  return json({ team: teamUnread, tasks: tasks, totalTasks: totalTasks });
}

// ─── Chat threads cache ────────────────────────────────────────
// getChatThreads() used to do 2-3 full getDataRange().getValues() reads
// (Master Tasks + all of Chat Messages) on EVERY call — every time someone
// opened "My Task Chats"/"All Task Chats", and every ~12s while they sat on
// that tab. As the Chat Messages sheet grows that gets slower and slower,
// which is what was making task-chat filtering feel sluggish. The heavy,
// identity-independent part (which tasks have threads, who's posted,
// what the last message was) is now built once and cached for a few
// seconds; only the cheap per-user part (their own read receipts) still
// hits the sheet on every request.
const CHAT_THREADS_CACHE_KEY = 'nexus_chat_threads_base_v1';
const CHAT_THREADS_CACHE_TTL_SEC = 20;

// Builds { taskId: { taskTitle, assignedTo, status, authors:[...], last:{message,author,timestamp}, entries:[{t,a}...] } }
// `entries` deliberately drops message text (keeps only timestamp+author)
// so this stays small enough to fit in Apps Script's cache (100KB/key) —
// only the single most recent message keeps its full text, for previews.
function buildChatThreadsBase() {
  const taskById = {};
  sheetToObjects('1. Master Tasks').forEach(t => { taskById[String(t['ID'])] = t; });

  const byTask = {};
  sheetToObjects('10. Chat Messages')
    .filter(r => r.Channel === 'task' && r.TaskID)
    .forEach(r => { (byTask[r.TaskID] = byTask[r.TaskID] || []).push(r); });

  const base = {};
  Object.keys(byTask).forEach(taskId => {
    const msgs = byTask[taskId].sort((a, b) => new Date(a.Timestamp) - new Date(b.Timestamp));
    const last = msgs[msgs.length - 1];
    const task = taskById[taskId];
    base[taskId] = {
      taskTitle: task ? task['Task Description'] : ('Task ' + taskId),
      assignedTo: task ? task['Assigned To'] : '',
      status: task ? task['Status'] : '',
      authors: Array.from(new Set(msgs.map(r => r.Author))),
      last: { message: last.Message, author: last.Author, timestamp: last.Timestamp },
      entries: msgs.map(r => ({ t: r.Timestamp, a: r.Author })),
    };
  });
  return base;
}

function getChatThreadsBaseCached() {
  try {
    const cache = CacheService.getScriptCache();
    const hit = cache.get(CHAT_THREADS_CACHE_KEY);
    if (hit) return JSON.parse(hit);
    const base = buildChatThreadsBase();
    try { cache.put(CHAT_THREADS_CACHE_KEY, JSON.stringify(base), CHAT_THREADS_CACHE_TTL_SEC); } catch(pe) { /* over 100KB — skip caching, still return fresh data */ }
    return base;
  } catch (e) {
    // CacheService itself unavailable for some reason — fall back to a
    // direct (slower) read rather than failing the whole request.
    return buildChatThreadsBase();
  }
}

// Returns every task-chat thread the caller has access to (see
// threadVisibleTo_ — leader, trusted-for-Master-Tasks, assigned/support/
// supervisor on the task, or a prior author in that thread) that has at
// least one message, with task title/assignee/unread count/last message,
// so the frontend can render a "browse all task chats" list (scoped to what
// this person can actually see) and a "My Task Chats" filter (threads where
// I'm assigned/support/supervisor, or I've already posted in that thread).
// Previously this returned EVERY task thread in the org to anyone logged
// in — an info leak (titles/assignees/previews for tasks you're not on)
// and the source of badge/list counts showing numbers unrelated to you.
function getChatThreads(token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  const me = canonicalizeName(identity.name).trim().toUpperCase();

  const reads = sheetToObjects('12. Chat Read State').filter(r => String(r.Name).trim() === String(identity.name).trim());
  const lastReadFor = (channel) => {
    const r = reads.find(x => String(x.Channel).trim() === String(channel).trim());
    return r ? new Date(r.LastReadTimestamp) : new Date(0);
  };

  const base = getChatThreadsBaseCached();
  const tasksById = buildTasksById_();

  const threads = Object.keys(base)
    .filter(taskId => threadVisibleTo_(identity, taskId, tasksById, base))
    .map(taskId => {
      const entry = base[taskId];
      const cutoff = lastReadFor(taskId);
      const unread = entry.entries.filter(e => new Date(e.t) > cutoff && e.a !== identity.name).length;
      const participants = entry.authors.map(a => canonicalizeName(String(a)).trim().toUpperCase());
      const assignee = canonicalizeName(String(entry.assignedTo || '')).trim().toUpperCase();
      const task = tasksById[taskId];
      const onTask = task ? isPersonOnTask(me, task) : false;
      return {
        taskId: taskId,
        taskTitle: entry.taskTitle,
        assignedTo: entry.assignedTo,
        status: entry.status,
        unread: unread,
        lastMessage: entry.last.message,
        lastAuthor: entry.last.author,
        lastTimestamp: entry.last.timestamp,
        isMine: onTask || (!!assignee && assignee === me) || participants.indexOf(me) !== -1,
      };
    }).sort((a, b) => new Date(b.lastTimestamp) - new Date(a.lastTimestamp));

  return json({ threads: threads });
}

// ═════════════════════════════════════
//  WEB PUSH — subscriptions + sending (via Vercel serverless function)
// ═════════════════════════════════════
// Apps Script cannot do the ECDH/AES-GCM payload encryption Web Push requires,
// so it just stores subscriptions and asks the Vercel endpoint (which holds
// the VAPID private key + the `web-push` npm package) to do the actual send.

function savePushSubscription(token, subscription) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!subscription || !subscription.endpoint) return json({ error: 'subscription required.' });
  const name  = identity.name;
  const sheet = getOrCreateSheet('11. Push Subscriptions', TAB_HEADERS['11. Push Subscriptions']);
  const data  = sheet.getDataRange().getValues();
  const now   = new Date().toISOString();
  const keysJson = JSON.stringify(subscription.keys || {});
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(name).trim()) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([[name, subscription.endpoint, keysJson, now]]);
      return json({ success: true });
    }
  }
  sheet.appendRow([name, subscription.endpoint, keysJson, now]);
  return json({ success: true });
}

function removePushSubscription(token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  const name  = identity.name;
  const sheet = getOrCreateSheet('11. Push Subscriptions', TAB_HEADERS['11. Push Subscriptions']);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(name).trim()) { sheet.deleteRow(i + 1); break; }
  }
  return json({ success: true });
}

// Sends `payload` (title/body/url/tag) to every subscribed user except `excludeName`.
function broadcastPush(excludeName, payload) {
  if (!VERCEL_PUSH_URL || VERCEL_PUSH_URL.indexOf('PASTE_') === 0) return;
  const subs = sheetToObjects('11. Push Subscriptions');
  subs.forEach(sub => {
    if (!sub.Endpoint) return;
    if (excludeName && String(sub.Name).trim() === String(excludeName).trim()) return;
    let keys = {};
    try { keys = JSON.parse(sub.Keys_JSON || '{}'); } catch(e) {}
    try {
      UrlFetchApp.fetch(VERCEL_PUSH_URL, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify({
          subscription: { endpoint: sub.Endpoint, keys: keys },
          payload: payload,
        }),
      });
    } catch(e) {
      Logger.log('Push send failed for ' + sub.Name + ': ' + e.message);
    }
  });
}

// Like broadcastPush, but only sends to the specific people in `names`
// (matched by canonicalized display name) instead of "everyone except X".
// Used for task-chat replies, which should reach the assignee/participants
// of that thread, not the entire team.
function broadcastPushTo(names, payload) {
  if (!VERCEL_PUSH_URL || VERCEL_PUSH_URL.indexOf('PASTE_') === 0) return;
  const wanted = (names || []).map(n => canonicalizeName(String(n)).trim().toUpperCase());
  if (!wanted.length) return;
  const subs = sheetToObjects('11. Push Subscriptions');
  subs.forEach(sub => {
    if (!sub.Endpoint) return;
    if (wanted.indexOf(canonicalizeName(String(sub.Name)).trim().toUpperCase()) === -1) return;
    let keys = {};
    try { keys = JSON.parse(sub.Keys_JSON || '{}'); } catch(e) {}
    try {
      UrlFetchApp.fetch(VERCEL_PUSH_URL, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify({
          subscription: { endpoint: sub.Endpoint, keys: keys },
          payload: payload,
        }),
      });
    } catch(e) {
      Logger.log('Push send failed for ' + sub.Name + ': ' + e.message);
    }
  });
}

// ═════════════════════════════════════
//  IN-APP NOTIFICATIONS + PER-PERSON PREFERENCE
//  ─────────────────────────────────────
//  Historically the only way NEXUS told anyone about anything was email —
//  assignments, escalations, deadline reminders, EOD nudges, chat, partner
//  follow-ups, all of it landed in an inbox. That doesn't scale once a
//  leader is responsible for 100+ people: every straggler's escalation
//  chain adds more email. This gives everyone a Notify_Pref on the Team
//  Members roster ('email' | 'inapp' | 'both', default 'both') and a
//  parallel in-app feed (this sheet) that the bell icon in the frontend
//  reads. Email call sites elsewhere in this file now check wantsEmail()
//  before sending; anything that should also/instead show up in-app calls
//  notifyInApp() here, which respects wantsInApp() per recipient and
//  reuses the existing push-subscription infra for the OS-level nudge.
// ═════════════════════════════════════

const NOTIFY_PREF_COL     = 'Notify_Pref';  // column appended to '9. Team Members'
const NOTIFY_PREF_DEFAULT = 'both';         // 'email' | 'inapp' | 'both'

// Finds (or lazily adds) the Notify_Pref column on '9. Team Members' and
// returns its 1-based column index. The sheet predates this feature, so
// TAB_HEADERS alone won't add the column — this backfills it exactly once.
function ensureNotifyPrefColumn_() {
  const sheet = getOrCreateSheet('9. Team Members', TAB_HEADERS['9. Team Members']);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = headers.indexOf(NOTIFY_PREF_COL);
  if (idx === -1) {
    idx = headers.length;
    sheet.getRange(1, idx + 1).setValue(NOTIFY_PREF_COL).setFontWeight('bold');
  }
  return idx + 1; // 1-based for Range APIs
}

function getNotifyPref(name) {
  const canon = canonicalizeName(String(name || '').trim());
  if (!canon) return NOTIFY_PREF_DEFAULT;
  const row = sheetToObjects('9. Team Members').find(m => canonicalizeName(String(m.Name || '').trim()) === canon);
  const pref = row && row[NOTIFY_PREF_COL] ? String(row[NOTIFY_PREF_COL]).trim().toLowerCase() : '';
  return ['email', 'inapp', 'both'].includes(pref) ? pref : NOTIFY_PREF_DEFAULT;
}

function wantsEmail(name) { const p = getNotifyPref(name); return p === 'email' || p === 'both'; }
function wantsInApp(name) { const p = getNotifyPref(name); return p === 'inapp' || p === 'both'; }

// ── Task-assignment notifications: forced Both, leader-only override ──
// Every team member used to be able to dial assignment notifications down
// to "in-app only" or "email only" via their own Notify_Pref — which meant
// someone could quietly opt out of being told they'd been handed a task.
// Assignment notifications now default to BOTH email + in-app for
// everyone, and a team member can no longer change that themselves; only a
// leader can flip a specific person back to "follow their own preference"
// (Assign_Notify_Mode = 'follow_pref'), from the Team Directory card, not
// the person's own profile. Missing/blank/anything-else = forced Both.
const ASSIGN_NOTIFY_MODE_COL = 'Assign_Notify_Mode'; // '' / 'both' (forced) | 'follow_pref' (leader override)

function ensureAssignNotifyModeColumn_() {
  const sheet = getOrCreateSheet('9. Team Members', TAB_HEADERS['9. Team Members']);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = headers.indexOf(ASSIGN_NOTIFY_MODE_COL);
  if (idx === -1) {
    idx = headers.length;
    sheet.getRange(1, idx + 1).setValue(ASSIGN_NOTIFY_MODE_COL).setFontWeight('bold');
  }
  return idx + 1; // 1-based for Range APIs
}

function assignmentForcedBoth_(name) {
  const canon = canonicalizeName(String(name || '').trim());
  if (!canon) return true;
  const row = sheetToObjects('9. Team Members').find(m => canonicalizeName(String(m.Name || '').trim()) === canon);
  const mode = row && row[ASSIGN_NOTIFY_MODE_COL] ? String(row[ASSIGN_NOTIFY_MODE_COL]).trim().toLowerCase() : '';
  return mode !== 'follow_pref';
}
function assignmentWantsEmail(name) { return assignmentForcedBoth_(name) || wantsEmail(name); }
function assignmentWantsInApp(name) { return assignmentForcedBoth_(name) || wantsInApp(name); }

// Leader-only: flips whether a specific team member's task-ASSIGNMENT
// notifications are forced to Both (default) or follow their own personal
// Notify_Pref instead. Mirrors updateTeamPermission's pattern exactly.
function setAssignNotifyMode(token, targetName, mode) {
  const callerName = getLeaderIdentity(token);
  if (!callerName) return json({ error: 'Leader login required to change this.' });
  mode = String(mode || '').trim().toLowerCase();
  if (!['both', 'follow_pref'].includes(mode)) return json({ error: 'Invalid mode — use both or follow_pref.' });
  targetName = String(targetName || '').trim();
  if (!targetName) return json({ error: 'No team member specified.' });

  const col   = ensureAssignNotifyModeColumn_();
  const sheet = SS.getSheetByName('9. Team Members');
  const data  = sheet.getDataRange().getValues();
  const targetCanon = canonicalizeName(targetName).toUpperCase();
  for (let i = 1; i < data.length; i++) {
    if (canonicalizeName(String(data[i][0]).trim()).toUpperCase() === targetCanon) {
      sheet.getRange(i + 1, col).setValue(mode === 'follow_pref' ? 'follow_pref' : 'both');
      Logger.log('Assign_Notify_Mode change: ' + callerName + ' set ' + targetName + ' = ' + mode);
      return json({ success: true, name: targetName, mode: mode });
    }
  }
  return json({ error: 'Team member "' + targetName + '" not found on the roster.' });
}

// ── Digest emails (Monday / daily) — separate preference from everything
// else. Previously digests shared Notify_Pref with escalations/assignments/
// chat, so "no escalation emails" also silently meant "no digest emails"
// and vice versa. This gives digests their own independent toggle.
const DIGEST_PREF_COL     = 'Digest_Notify_Pref'; // 'email' | 'inapp' | 'both' | 'none'
const DIGEST_PREF_DEFAULT = 'both';

function ensureDigestPrefColumn_() {
  const sheet = getOrCreateSheet('9. Team Members', TAB_HEADERS['9. Team Members']);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = headers.indexOf(DIGEST_PREF_COL);
  if (idx === -1) {
    idx = headers.length;
    sheet.getRange(1, idx + 1).setValue(DIGEST_PREF_COL).setFontWeight('bold');
  }
  return idx + 1;
}

function getDigestNotifyPref(name) {
  const canon = canonicalizeName(String(name || '').trim());
  if (!canon) return DIGEST_PREF_DEFAULT;
  const row = sheetToObjects('9. Team Members').find(m => canonicalizeName(String(m.Name || '').trim()) === canon);
  const pref = row && row[DIGEST_PREF_COL] ? String(row[DIGEST_PREF_COL]).trim().toLowerCase() : '';
  return ['email', 'inapp', 'both', 'none'].includes(pref) ? pref : DIGEST_PREF_DEFAULT;
}
function wantsDigestEmail(name) { const p = getDigestNotifyPref(name); return p === 'email' || p === 'both'; }
function wantsDigestInApp(name) { const p = getDigestNotifyPref(name); return p === 'inapp' || p === 'both'; }

function emailIfWantedDigest_(name, email, subject, htmlBody, plainBody, options) {
  if (!email || !wantsDigestEmail(name)) return false;
  return sendViaResend(email, subject, htmlBody, plainBody, options);
}

function saveDigestNotifyPref(token, pref) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  pref = String(pref || '').trim().toLowerCase();
  if (!['email', 'inapp', 'both', 'none'].includes(pref)) {
    return json({ error: 'Invalid preference — use email, inapp, both, or none.' });
  }
  const col   = ensureDigestPrefColumn_();
  const sheet = SS.getSheetByName('9. Team Members');
  const data  = sheet.getDataRange().getValues();
  const canon = canonicalizeName(identity.name);
  for (let i = 1; i < data.length; i++) {
    if (canonicalizeName(String(data[i][0]).trim()) === canon) {
      sheet.getRange(i + 1, col).setValue(pref);
      return json({ success: true, pref: pref });
    }
  }
  return json({ error: 'Could not find your row on the Team Members roster.' });
}

// Convenience for call sites that currently do `sendViaResend(email, ...)`
// straight to a named team member — swap that for this and it silently
// no-ops when that person has opted out of email. Returns the same
// true/false "did it actually send" signal sendViaResend does.
function emailIfWanted_(name, email, subject, htmlBody, plainBody, options) {
  if (!email || !wantsEmail(name)) return false;
  return sendViaResend(email, subject, htmlBody, plainBody, options);
}

function saveNotifyPref(token, pref) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  pref = String(pref || '').trim().toLowerCase();
  if (!['email', 'inapp', 'both'].includes(pref)) {
    return json({ error: 'Invalid preference — use email, inapp, or both.' });
  }
  const col   = ensureNotifyPrefColumn_();
  const sheet = SS.getSheetByName('9. Team Members');
  const data  = sheet.getDataRange().getValues();
  const canon = canonicalizeName(identity.name);
  for (let i = 1; i < data.length; i++) {
    if (canonicalizeName(String(data[i][0]).trim()) === canon) {
      sheet.getRange(i + 1, col).setValue(pref);
      return json({ success: true, pref: pref });
    }
  }
  return json({ error: 'Could not find your row on the Team Members roster.' });
}

// One notification row per recipient. taskId (when present) is what lets
// the frontend deep-link straight back into the relevant task instead of
// the generic homepage.
function logInAppNotification_(name, n) {
  const sheet = getOrCreateSheet('18. Notifications', TAB_HEADERS['18. Notifications']);
  const id = 'NTF-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
  sheet.appendRow([
    id, canonicalizeName(String(name).trim()), n.type || 'general',
    n.title || 'NEXUS', n.body || '', n.url || APP_LINK, n.taskId || '',
    n.tag || '', new Date().toISOString(), 'FALSE',
  ]);
  return id;
}

// Central fan-out: for each name in `names`, writes an in-app row if they
// want in-app notifications, then pushes to whichever of those people also
// have a push subscription. Doesn't touch email at all — callers handle
// email separately via wantsEmail()/emailIfWanted_(), since email bodies
// are longer and already hand-built per call site.
// `opts.force` — names that bypass the gate entirely (used by assignment
// notifications, which are forced to Both by default; see
// assignmentForcedBoth_ above). `opts.gate` — a wantsInApp-shaped function
// to use instead of the general per-person Notify_Pref (used by digests,
// which have their own independent Digest_Notify_Pref).
function notifyInApp(names, n, opts) {
  opts = opts || {};
  const gate = opts.gate || wantsInApp;
  const forceSet = new Set((opts.force || []).map(x => canonicalizeName(String(x || '').trim())));
  const uniqueNames = Array.from(new Set(
    (names || []).map(x => canonicalizeName(String(x || '').trim())).filter(Boolean)
  ));
  const pushable = [];
  uniqueNames.forEach(name => {
    if (!forceSet.has(name) && !gate(name)) return;
    try { logInAppNotification_(name, n); } catch (e) { Logger.log('logInAppNotification_ failed for ' + name + ': ' + e.message); }
    pushable.push(name);
  });
  if (pushable.length) {
    try {
      broadcastPushTo(pushable, {
        title: n.title, body: n.body,
        url: n.url || APP_LINK, tag: n.tag || 'nexus-notify',
      });
    } catch (e) { Logger.log('notifyInApp push failed: ' + e.message); }
  }
}

// Builds a deep-link URL back into the app. `task` opens the task modal
// directly (openDeepLinkedTaskIfAny already handles ?task=); `chat=1`
// alongside a task opens that task's discussion thread instead of just
// the task modal; `chat=team` opens the General chat panel.
function notifUrl_(params) {
  const base = APP_LINK.replace(/\/$/, '') + '/';
  const qs = Object.keys(params || {})
    .filter(k => params[k] !== undefined && params[k] !== null && params[k] !== '')
    .map(k => k + '=' + encodeURIComponent(params[k]))
    .join('&');
  return qs ? base + '?' + qs : base;
}

function getNotifications(token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  const canon = canonicalizeName(identity.name);
  let rows = sheetToObjects('18. Notifications').filter(r => canonicalizeName(String(r.Recipient || '').trim()) === canon);
  rows.sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
  rows = rows.slice(0, 60);
  const unread = rows.filter(r => r.Read !== 'TRUE' && r.Read !== true).length;
  return json({ notifications: rows, unread: unread });
}

function markNotificationRead(token, id) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  const sheet  = getOrCreateSheet('18. Notifications', TAB_HEADERS['18. Notifications']);
  const data   = sheet.getDataRange().getValues();
  const canon  = canonicalizeName(identity.name);
  const idIdx  = TAB_HEADERS['18. Notifications'].indexOf('ID');
  const recIdx = TAB_HEADERS['18. Notifications'].indexOf('Recipient');
  const rdIdx  = TAB_HEADERS['18. Notifications'].indexOf('Read');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(id) && canonicalizeName(String(data[i][recIdx]).trim()) === canon) {
      sheet.getRange(i + 1, rdIdx + 1).setValue('TRUE');
      return json({ success: true });
    }
  }
  return json({ error: 'Notification not found.' });
}

function markAllNotificationsRead(token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  const sheet  = getOrCreateSheet('18. Notifications', TAB_HEADERS['18. Notifications']);
  const data   = sheet.getDataRange().getValues();
  const canon  = canonicalizeName(identity.name);
  const recIdx = TAB_HEADERS['18. Notifications'].indexOf('Recipient');
  const rdIdx  = TAB_HEADERS['18. Notifications'].indexOf('Read');
  let changed = 0;
  for (let i = 1; i < data.length; i++) {
    if (canonicalizeName(String(data[i][recIdx]).trim()) === canon && data[i][rdIdx] !== 'TRUE') {
      sheet.getRange(i + 1, rdIdx + 1).setValue('TRUE');
      changed++;
    }
  }
  return json({ success: true, changed: changed });
}

// ═════════════════════════════════════
//  TEAM ROSTER — Add Member (leader-only)
// ═════════════════════════════════════

function addTeamMember(data, token) {
  const callerName = getLeaderIdentity(token);
  if (!callerName) return json({ error: 'Leader login required to add a team member.' });

  const name  = String((data && data['Name'])  || '').trim();
  const email = String((data && data['Email']) || '').trim();
  if (!name)  return json({ error: 'Name is required.' });
  if (!email) return json({ error: 'Email is required.' });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json({ error: 'That email address doesn\'t look valid.' });

  const existing = sheetToObjects('9. Team Members');
  if (existing.some(m => String(m.Name || '').trim().toLowerCase() === name.toLowerCase())) {
    return json({ error: 'A team member named "' + name + '" already exists.' });
  }

  getOrCreateSheet('9. Team Members', TAB_HEADERS['9. Team Members']);
  const sheet = SS.getSheetByName('9. Team Members');
  sheet.appendRow([
    name, email,
    String((data && data['Role'])       || '').trim(),
    String((data && data['Department']) || '').trim(),
    String((data && data['EOD_Days'])   || '').trim(),
    String((data && data['Notes'])      || '').trim(),
  ]);
  // New rows start Pending — googleLogin refuses to log them in until this
  // flips to Active via confirmInvite (invitee clicks the emailed link) or
  // approveTeamMember (leader override). appendRow above leaves Status blank
  // (it's past the columns explicitly listed), so set it explicitly rather
  // than relying on that blank being read as Pending — a blank Status means
  // Active everywhere else (pre-migration rosters), so Pending must be
  // written, never assumed.
  setTeamMemberStatus(name, 'Pending');

  const inviteSent = sendTeamInviteEmail(name, email, callerName);

  Logger.log('Team member added (Pending): ' + name + ' (' + email + ') by ' + callerName + (inviteSent ? ' — invite sent' : ' — invite email FAILED, approve manually'));
  return json({ success: true, name, email, status: 'Pending', inviteSent });
}

// Generates a one-time confirmation token for `name`/`email`, stores it
// (7-day TTL — see TEAM_INVITE_TTL_MS), and emails the invitee a link back
// into this same web app's doGet (action=confirmInvite) that flips their
// roster row to Active. Returns false (never throws) on email failure so a
// bad send doesn't undo the roster row itself — the row just stays Pending
// until a leader calls approveTeamMember instead.
function sendTeamInviteEmail(name, email, invitedBy) {
  try {
    const token = Utilities.getUuid();
    const store = JSON.parse(PropertiesService.getScriptProperties().getProperty(TEAM_INVITE_PROP) || '{}');
    Object.keys(store).forEach(t => { if (store[t].expires < Date.now()) delete store[t]; });
    store[token] = { name: name, email: email, expires: Date.now() + TEAM_INVITE_TTL_MS };
    PropertiesService.getScriptProperties().setProperty(TEAM_INVITE_PROP, JSON.stringify(store));

    const confirmUrl = ScriptApp.getService().getUrl() + '?action=confirmInvite&token=' + encodeURIComponent(token);
    const subject = 'You\'ve been invited to NEXUS · ABION Industries';
    const html =
      '<p>Hi ' + name + ',</p>' +
      '<p>' + invitedBy + ' added you to NEXUS, ABION\'s internal operations platform. ' +
      'Confirm your account to finish setting up access:</p>' +
      '<p><a href="' + confirmUrl + '">Confirm my account</a></p>' +
      '<p>This link expires in 7 days. If you weren\'t expecting this, you can ignore it.</p>';
    const plain = 'Hi ' + name + ',\n\n' + invitedBy + ' added you to NEXUS, ABION\'s internal operations platform.\n' +
      'Confirm your account here: ' + confirmUrl + '\n\nThis link expires in 7 days.';
    return sendViaResend([email], subject, html, plain, {});
  } catch (e) {
    Logger.log('sendTeamInviteEmail error: ' + e.message);
    return false;
  }
}

// Public (no-auth) endpoint the invite email link hits. Deliberately not
// behind resolveIdentity()/requireAnyAuth() — the invitee has no session yet;
// the one-time token IS the credential. Consumes the token on success (or on
// a stale/missing roster match) so it can't be replayed. Returns a plain
// HTML page since this is opened directly in a browser, not called from the
// frontend's fetch()-based API client.
function confirmInvite(token) {
  const store = JSON.parse(PropertiesService.getScriptProperties().getProperty(TEAM_INVITE_PROP) || '{}');
  const invite = store[token];
  const page = (title, message) => HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;max-width:480px;margin:60px auto;text-align:center;padding:0 20px">' +
    '<h2>' + title + '</h2><p>' + message + '</p>' +
    '<p><a href="' + APP_LINK + '">Go to NEXUS</a></p></div>'
  ).setTitle(title);

  if (!invite || invite.expires < Date.now()) {
    if (invite) { delete store[token]; PropertiesService.getScriptProperties().setProperty(TEAM_INVITE_PROP, JSON.stringify(store)); }
    return page('Invite not found', 'This confirmation link is invalid or has expired. Ask a leader to resend your invite or approve you directly.');
  }

  delete store[token];
  PropertiesService.getScriptProperties().setProperty(TEAM_INVITE_PROP, JSON.stringify(store));
  const updated = setTeamMemberStatus(invite.name, 'Active');
  if (!updated) return page('Account not found', 'We couldn\'t find your roster row anymore — ask a leader to re-add you.');

  Logger.log('Invite confirmed: ' + invite.name + ' (' + invite.email + ')');
  return page('You\'re confirmed!', 'Your NEXUS account is now active. You can sign in with Google.');
}

// Leader override for a Pending roster row — covers an invite email that
// bounced or was never opened. Sets Status straight to Active without
// requiring the token flow.
function approveTeamMember(name, token) {
  const callerName = getLeaderIdentity(token);
  if (!callerName) return json({ error: 'Leader login required to approve a team member.' });
  name = String(name || '').trim();
  if (!name) return json({ error: 'No team member specified.' });
  const updated = setTeamMemberStatus(name, 'Active');
  if (!updated) return json({ error: 'Team member "' + name + '" not found on the roster.' });
  Logger.log('Team member approved: ' + name + ' by ' + callerName);
  return json({ success: true, name, status: 'Active' });
}

// Leader-only: grant or revoke a team member's trusted-editor access for one
// tab (Master Tasks / COS / Partnerships locked fields + row-ownership
// bypass on that tab). Writes the flag straight to their roster row on
// "9. Team Members" so it takes effect immediately — no code edit or
// redeploy — and logs who changed it, for whom, and the before/after value
// to "15. Permission Log" for the audit trail.
function updateTeamPermission(targetName, tabName, grant, token) {
  const callerName = getLeaderIdentity(token);
  if (!callerName) return json({ error: 'Leader login required to change permissions.' });

  const col = TAB_TRUST_COLUMN[tabName];
  if (!col) return json({ error: 'Unknown or unsupported tab for permissions: ' + tabName });
  targetName = String(targetName || '').trim();
  if (!targetName) return json({ error: 'No team member specified.' });

  const sheet   = getOrCreateSheet('9. Team Members', TAB_HEADERS['9. Team Members']);
  const values  = sheet.getDataRange().getValues();
  let headers   = values[0] || [];

  // Sheet predates this permission column (created before this feature
  // shipped) — add it live rather than failing every grant attempt.
  let colIdx = headers.indexOf(col);
  if (colIdx === -1) {
    colIdx = headers.length;
    sheet.getRange(1, colIdx + 1).setValue(col);
    headers = headers.concat([col]);
  }

  const nameIdx = headers.indexOf('Name');
  let rowIdx = -1;
  const targetCanon = canonicalizeName(targetName).toUpperCase();
  for (let i = 1; i < values.length; i++) {
    if (canonicalizeName(String(values[i][nameIdx] || '')).toUpperCase() === targetCanon) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return json({ error: 'Team member "' + targetName + '" not found on the roster.' });

  const oldValue = String((values[rowIdx][colIdx] !== undefined ? values[rowIdx][colIdx] : '') || 'N').trim() || 'N';
  const newValue = grant ? 'Y' : 'N';
  sheet.getRange(rowIdx + 1, colIdx + 1).setValue(newValue);

  appendRow('15. Permission Log', {
    Timestamp : new Date().toISOString(),
    GrantedBy : callerName,
    GrantedTo : targetName,
    Tab       : tabName,
    OldValue  : oldValue,
    NewValue  : newValue,
  });

  Logger.log('Permission change: ' + callerName + ' set ' + targetName + ' / ' + tabName + ' = ' + newValue);
  return json({ success: true, name: targetName, tabName: tabName, value: newValue });
}

// ═════════════════════════════════════
//  LEADER AUTH
//  Default passphrases live in LEADER_PASSPHRASES below.
//  If a leader changes their passphrase, the new value is stored in
//  PropertiesService (persists across deployments) and overrides the default.
//  Tokens stored in PropertiesService (survives across requests).
//  Token TTL: 24 hours.
// ═════════════════════════════════════

// Returns the currently-active passphrase for a leader key (e.g. 'AYO'),
// preferring any password they've changed to over the hardcoded default.
function getLeaderPassphrase(nameKey) {
  const override = PropertiesService.getScriptProperties().getProperty('nexus_pw_' + nameKey);
  return (override !== null && override !== undefined) ? override : LEADER_PASSPHRASES[nameKey];
}

function setLeaderPassphrase(nameKey, newPass) {
  PropertiesService.getScriptProperties().setProperty('nexus_pw_' + nameKey, newPass);
}

function leaderLogin(name, passphrase) {
  if (!name || !passphrase) return json({ error: 'Name and passphrase required.' });
  Logger.log('leaderLogin attempt — name: ' + name);
  const nameKey  = name.toUpperCase();
  const expected = getLeaderPassphrase(nameKey);
  if (!expected || expected !== passphrase.trim()) {
    Logger.log('leaderLogin FAILED — wrong passphrase for: ' + name);
    return json({ error: 'Invalid name or passphrase.' });
  }
  // Store the canonical display name ("Ayo T. Joshua"), not the raw key
  // ("AYO") — this is what shows up as the chat Author, so it has to match
  // what Google login stores for the same person.
  const displayName = LEADER_DISPLAY_NAMES[nameKey] || name;
  Logger.log('leaderLogin SUCCESS — ' + displayName);
  // Generate token, store with expiry
  const token   = Utilities.getUuid();
  const expires = Date.now() + 24 * 60 * 60 * 1000; // 24h
  const store   = JSON.parse(PropertiesService.getScriptProperties().getProperty(LEADER_TOKEN_PROP) || '{}');
  // Clean expired tokens
  Object.keys(store).forEach(t => { if (store[t].expires < Date.now()) delete store[t]; });
  store[token] = { name: displayName, expires };
  PropertiesService.getScriptProperties().setProperty(LEADER_TOKEN_PROP, JSON.stringify(store));
  Logger.log('Leader login: ' + displayName);
  // Passphrase login has no Google picture of its own — pull whatever was
  // last saved to the roster (e.g. from a prior Google login on another
  // device) so the avatar still follows this leader here too.
  const rosterRow = sheetToObjects('9. Team Members').find(m => String(m.Name||'').trim() === displayName);
  const picture = (rosterRow && rosterRow.Picture) || '';
  return json({ success: true, token, name: displayName, picture });
}

function leaderLogout(token) {
  if (!token) return json({ success: true });
  try {
    const store = JSON.parse(PropertiesService.getScriptProperties().getProperty(LEADER_TOKEN_PROP) || '{}');
    delete store[token];
    PropertiesService.getScriptProperties().setProperty(LEADER_TOKEN_PROP, JSON.stringify(store));
  } catch(e) {}
  return json({ success: true });
}

// Lets a logged-in leader change their own passphrase.
// Requires a valid session token AND the correct current passphrase.
function changeLeaderPassword(token, currentPassphrase, newPassphrase) {
  const name = validateLeaderToken(token);
  if (!name) return json({ error: 'Leader login required.' });
  if (!currentPassphrase || !newPassphrase) return json({ error: 'Current and new passphrase are required.' });
  if (newPassphrase.trim().length < 6) return json({ error: 'New passphrase must be at least 6 characters.' });

  const nameKey  = name.toUpperCase();
  const expected = getLeaderPassphrase(nameKey);
  if (!expected || expected !== currentPassphrase.trim()) return json({ error: 'Current passphrase is incorrect.' });

  setLeaderPassphrase(nameKey, newPassphrase.trim());
  Logger.log('Leader passphrase changed: ' + name);
  return json({ success: true, message: 'Passphrase updated.' });
}

// ═════════════════════════════════════
//  TEAM AUTH — Google Sign-In
//  Anyone can open the Google popup, but a session is only issued if the
//  verified Google account's email matches someone in "9. Team Members".
//  This is what chat, push subscriptions, and status updates now check
//  identity against — never a client-supplied name string.
//  Token TTL: 30 days.
// ═════════════════════════════════════

// Verifies a Google ID token server-side via Google's public tokeninfo
// endpoint (no extra library/cost). Confirms it was issued for THIS app
// (aud === GOOGLE_CLIENT_ID) and that the email is verified.
function verifyGoogleIdToken(idToken) {
  if (!idToken) return null;
  if (GOOGLE_CLIENT_ID.indexOf('PASTE_') === 0) {
    Logger.log('verifyGoogleIdToken: GOOGLE_CLIENT_ID not configured yet.');
    return null;
  }
  try {
    const resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) return null;
    const info = JSON.parse(resp.getContentText());
    const verifiedFlag = info.email_verified === 'true' || info.email_verified === true;
    if (!info.email || !verifiedFlag) return null;
    if (info.aud !== GOOGLE_CLIENT_ID) { Logger.log('ID token aud mismatch'); return null; }
    return { email: String(info.email).toLowerCase().trim(), name: info.name || '', picture: info.picture || '' };
  } catch (e) { Logger.log('verifyGoogleIdToken error: ' + e.message); return null; }
}

// Persists (or refreshes) a team member's Google profile picture URL on
// their "9. Team Members" roster row. Called on every successful Google
// login so the avatar stays current. Because it's stored on the roster
// (not just client-side), a leader who authenticates via passphrase on a
// different device still sees their last-known Google photo — the roster
// row is the single source of truth, regardless of which login path
// (Google OAuth or passphrase) was used to sign in this session.
// No-ops if a picture URL isn't available (never overwrite a saved photo
// with blank just because this particular login didn't return one).
function saveTeamMemberPicture(name, pictureUrl) {
  if (!name || !pictureUrl) return;
  try {
    const sheet  = getOrCreateSheet('9. Team Members', TAB_HEADERS['9. Team Members']);
    const values = sheet.getDataRange().getValues();
    let headers  = values[0] || [];

    // Sheet predates this column (created before this feature shipped) —
    // add it live rather than silently failing to save every photo.
    let colIdx = headers.indexOf('Picture');
    if (colIdx === -1) {
      colIdx = headers.length;
      sheet.getRange(1, colIdx + 1).setValue('Picture');
      headers = headers.concat(['Picture']);
    }

    const nameIdx = headers.indexOf('Name');
    const canon = canonicalizeName(name).toUpperCase();
    for (let i = 1; i < values.length; i++) {
      if (canonicalizeName(String(values[i][nameIdx] || '')).toUpperCase() === canon) {
        if (String(values[i][colIdx] || '') !== pictureUrl) {
          sheet.getRange(i + 1, colIdx + 1).setValue(pictureUrl);
        }
        return;
      }
    }
  } catch (e) { Logger.log('saveTeamMemberPicture error: ' + e.message); }
}

// Sets `name`'s Status cell on "9. Team Members" (adding the column live if
// this roster predates migrateAddTeamStatusColumn, same defensive pattern as
// saveTeamMemberPicture/updateTeamPermission above). Returns true if a
// matching row was found and updated, false otherwise — callers decide what
// "member not found" means for them.
function setTeamMemberStatus(name, status) {
  const sheet  = getOrCreateSheet('9. Team Members', TAB_HEADERS['9. Team Members']);
  const values = sheet.getDataRange().getValues();
  let headers  = values[0] || [];

  let colIdx = headers.indexOf('Status');
  if (colIdx === -1) {
    colIdx = headers.length;
    sheet.getRange(1, colIdx + 1).setValue('Status');
    headers = headers.concat(['Status']);
  }

  const nameIdx = headers.indexOf('Name');
  const canon = canonicalizeName(name).toUpperCase();
  for (let i = 1; i < values.length; i++) {
    if (canonicalizeName(String(values[i][nameIdx] || '')).toUpperCase() === canon) {
      sheet.getRange(i + 1, colIdx + 1).setValue(status);
      return true;
    }
  }
  return false;
}

function googleLogin(idToken) {
  const verified = verifyGoogleIdToken(idToken);
  if (!verified) return json({ error: 'Could not verify Google sign-in. Please try again.' });

  const roster = getTeamEmailMap(); // { Name -> Email }
  let matchedName = null;
  Object.keys(roster).forEach(name => {
    if (String(roster[name] || '').toLowerCase().trim() !== verified.email) return;
    // If this email matches more than one roster name (e.g. a leader's seed
    // entry plus a later "who actually owns this inbox" row added under
    // their own first name), don't let whichever one happened to load last
    // silently win. Prefer whichever match resolves to a designated leader,
    // so a leader's account never gets quietly demoted to team-member access.
    const isNewLeader  = isLeaderName(canonicalizeName(name));
    const isCurrLeader = matchedName && isLeaderName(canonicalizeName(matchedName));
    if (!matchedName || (isNewLeader && !isCurrLeader)) matchedName = name;
  });
  if (!matchedName) {
    return json({
      error: 'no_match',
      email: verified.email,
      message: 'That Google account (' + verified.email + ') isn\'t on the NEXUS roster yet. Ask a leader to add you under Team Members with this exact email.',
    });
  }
  matchedName = canonicalizeName(matchedName);

  // Invite gate: a roster row stays Pending until the invitee confirms via
  // the emailed link (or a leader approves them directly) — see
  // TEAM_INVITE_PROP comment. Blank/missing Status is treated as Active so
  // rosters from before this migration keep working unchanged; only an
  // EXPLICIT "Pending" blocks the login. This runs before saveTeamMemberPicture
  // deliberately — a Pending account shouldn't have anything about it touched
  // on a login attempt beyond this check.
  const rosterRowsForGate = sheetToObjects('9. Team Members');
  const gateRow = rosterRowsForGate.find(m => canonicalizeName(String(m.Name || '')).toUpperCase() === matchedName.toUpperCase());
  const status = String((gateRow && gateRow.Status) || '').trim().toLowerCase();
  if (status === 'pending') {
    return json({
      error: 'invite_pending',
      message: 'Your invite to NEXUS hasn\'t been confirmed yet. Check ' + verified.email + ' for the confirmation email, or ask a leader to approve you.',
    });
  }

  // Was defined but never actually called — every Google login carries a
  // fresh photo URL (verified.picture) but it was silently dropped instead
  // of being saved to the roster. Also no-ops safely if this particular
  // login didn't return a picture (see saveTeamMemberPicture's own guard).
  saveTeamMemberPicture(matchedName, verified.picture);

  const token   = Utilities.getUuid();
  const expires = Date.now() + 14 * 24 * 60 * 60 * 1000; // 14 days
  const store   = JSON.parse(PropertiesService.getScriptProperties().getProperty(TEAM_TOKEN_PROP) || '{}');
  Object.keys(store).forEach(t => { if (store[t].expires < Date.now()) delete store[t]; });
  store[token] = { name: matchedName, email: verified.email, expires };
  PropertiesService.getScriptProperties().setProperty(TEAM_TOKEN_PROP, JSON.stringify(store));
  const asLeader = isLeaderName(matchedName);
  // Prefer the roster's stored photo (falls back correctly even on a login
  // that didn't return a fresh one) over verified.picture directly, so the
  // client always gets whatever photo is actually on file.
  const rosterRows = sheetToObjects('9. Team Members');
  const rosterRow = rosterRows.find(m => canonicalizeName(String(m.Name || '')).toUpperCase() === matchedName.toUpperCase());
  const picture = (rosterRow && rosterRow.Picture) || verified.picture || '';
  Logger.log('Team Google login: ' + matchedName + ' (' + verified.email + ')' + (asLeader ? ' [leader]' : ''));
  return json({ success: true, token, name: matchedName, email: verified.email, isLeader: asLeader, picture });
}

function teamLogout(token) {
  if (!token) return json({ success: true });
  try {
    const store = JSON.parse(PropertiesService.getScriptProperties().getProperty(TEAM_TOKEN_PROP) || '{}');
    delete store[token];
    PropertiesService.getScriptProperties().setProperty(TEAM_TOKEN_PROP, JSON.stringify(store));
  } catch (e) {}
  return json({ success: true });
}

function validateTeamToken(token) {
  if (!token) return null;
  try {
    const store = JSON.parse(PropertiesService.getScriptProperties().getProperty(TEAM_TOKEN_PROP) || '{}');
    const entry = store[token];
    if (!entry || entry.expires < Date.now()) return null;
    return entry; // { name, email, expires }
  } catch (e) { return null; }
}

// Resolves "who is making this request" from EITHER a leader session or a
// Google-verified team session — never from a client-supplied name field.
// Returns null if the token doesn't match a live session of either kind.
// Guard helpers for the router: return null (falsy) when authorized, so
// callers can write `return requireLeader(token) || doTheThing();` — an
// error response short-circuits, otherwise the real handler runs.
function requireLeader(token) {
  return getLeaderIdentity(token) ? null : json({ error: 'Leader login required.' });
}
function requireAnyAuth(token) {
  return resolveIdentity(token) ? null : json({ error: 'Login required.' });
}
// Gates row CREATION on the tabs that respect the "+ NEW" button — leaders
// always pass; everyone else must carry the per-tab Trusted_* grant. This
// was previously missing entirely: enforceLeaderOnlyFields() only strips
// specific leader-only fields off the payload, it never blocked the append
// itself, so a non-trusted, non-leader caller could still POST a new row
// (with those fields silently dropped) even though the UI hides the button.
function requireNewButtonAccess(token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (identity.isLeader) return null;
  if (isTrustedForTab(identity.name, 'NewButton')) return null;
  return json({ error: 'You do not have permission to add new rows.' });
}

function resolveIdentity(token) {
  const leaderName = validateLeaderToken(token);
  if (leaderName) return { name: leaderName, email: getTeamEmailMap()[leaderName] || '', isLeader: true };
  const team = validateTeamToken(token);
  // Was hardcoded to false — meant a leader (e.g. Ayo) who signed in with
  // Google, the preferred/correct path, was never recognized as a leader
  // even though their identity resolved correctly against the roster.
  if (team) return { name: canonicalizeName(team.name), email: team.email, isLeader: isLeaderName(canonicalizeName(team.name)) };
  return null;
}

// Returns the caller's display name if authorized as a leader — via the
// legacy passphrase login OR a Google account whose roster name is a
// designated leader — otherwise null. Use this (not validateLeaderToken
// directly) anywhere that gates a leader-only action, so a Google-signed-in
// leader gets the same access as a legacy passphrase login.
function getLeaderIdentity(token) {
  const legacyName = validateLeaderToken(token);
  if (legacyName) return legacyName;
  const team = validateTeamToken(token);
  // Was checking isLeaderName(team.name) directly — but team.name is
  // whatever was stored in the session at login, which can be a
  // pre-alias/raw roster name (e.g. "Othniel") if the token predates a
  // NAME_ALIASES entry or was issued by an older deploy. resolveIdentity()
  // already canonicalizes before this same check; this now matches it, so
  // "am I logged in as a leader" and "does this leader-gated action pass"
  // never disagree with each other again.
  const name = team && canonicalizeName(team.name);
  if (name && isLeaderName(name)) return name;
  return null;
}

function deleteRowAction(tabName, id, token) {
  // Leader-only
  const callerName = getLeaderIdentity(token);
  if (!callerName) return json({ error: 'Leader login required to delete rows.' });

  const sheet = SS.getSheetByName(tabName);
  if (!sheet) return json({ error: 'Tab not found: ' + tabName });

  const allValues  = sheet.getDataRange().getValues();
  const headers    = allValues[0];
  const idStrategy = TAB_ID_COL[tabName] || 'ID';
  let targetRowIndex = -1;

  if (idStrategy === '_rowIndex') {
    const rowNum = parseInt(id, 10);
    if (!isNaN(rowNum) && rowNum >= 1 && rowNum < allValues.length) targetRowIndex = rowNum;
  } else {
    const idColIdx = headers.indexOf(idStrategy);
    for (let i = 1; i < allValues.length; i++) {
      if (String(allValues[i][idColIdx]).trim() === String(id).trim()) { targetRowIndex = i; break; }
    }
  }

  if (targetRowIndex === -1) return json({ error: 'Row not found: ' + id });

  sheet.deleteRow(targetRowIndex + 1); // +1 because sheet rows are 1-indexed
  Logger.log('Row deleted by ' + callerName + ' — tab: ' + tabName + ', id: ' + id);
  return json({ success: true });
}

function validateLeaderToken(token) {
  if (!token) return null;
  try {
    const store = JSON.parse(PropertiesService.getScriptProperties().getProperty(LEADER_TOKEN_PROP) || '{}');
    const entry = store[token];
    if (!entry || entry.expires < Date.now()) return null;
    return entry.name;
  } catch(e) { return null; }
}

// Shared RBAC guard for leader-only fields, per tab. Strips those fields
// from `data` in place if the caller isn't a logged-in leader or a trusted
// editor. Used on both row creation and row edits, for every tab listed here.
//
// Master Tasks: scheduling/assignment fields — who's on the hook and when.
// COS: the fields most worth forging — headcount and status, i.e. the
// numbers used to report real-world progress on the ground.
// Partnerships: reassigning who owns a partner, and Document Status (so
// nobody can mark a contract "Signed" that isn't).
const TAB_LEADER_ONLY_FIELDS = {
  '1. Master Tasks': ['Deadline','Deadline_Time','Depends On','Supervisor','Supervisor_Email','Assigned To','Assigned_To_Email','Priority','Support Person','Support_Person_Email','Business / Dept','Follow_Up_Hours'],
  '2. COS'         : ['Members Count','WhatsApp Group Created','Status'],
  '3. Partnerships': ['Assigned To','Document Status'],
};
// Back-compat: any call site that doesn't pass a tabName (or passes one not
// in the map above) falls back to the original Master Tasks field list, so
// existing callers keep working exactly as before.
const LEADER_ONLY_FIELDS = TAB_LEADER_ONLY_FIELDS['1. Master Tasks'];

function enforceLeaderOnlyFields(data, token, tabName) {
  const fieldList = TAB_LEADER_ONLY_FIELDS[tabName] || LEADER_ONLY_FIELDS;
  const needsLeaderAccess = fieldList.some(f => data[f] !== undefined);
  if (!needsLeaderAccess) return;
  // Was: resolved the caller via getLeaderIdentity(token), which only ever
  // returns a name for people who are ALREADY designated leaders — so a
  // trusted-but-non-leader editor (e.g. Amara via TRUSTED_EDITORS) could
  // never actually pass this check; the bypass was silently dead. Now
  // resolves identity from ANY valid session (team member or leader) so
  // both leaders and per-tab trusted grants work as intended.
  const identity = resolveIdentity(token);
  if (identity && identity.isLeader) return; // full leaders always pass
  const isTrusted = identity && isTrustedForTab(identity.name, tabName);
  if (!isTrusted) {
    fieldList.forEach(f => { delete data[f]; });
  }
}

// ═════════════════════════════════════
//  READ
// ═════════════════════════════════════

// ─── Auto-tracked activity (not self-reported) ─────────────────
// Called from updateRow/updateTaskStatus whenever a real, verifiable
// change lands on a task — never from anything the team member typed
// into a text field. `name` is the actual authenticated caller, resolved
// server-side from their token, so this can't be spoofed by editing the
// EOD form.
function logAutoActivity(name, taskId, taskDesc, field, oldVal, newVal) {
  try {
    if (String(oldVal||'') === String(newVal||'')) return; // no real change, don't log noise
    const sheet = getOrCreateSheet('14. Auto Activity', TAB_HEADERS['14. Auto Activity']);
    sheet.appendRow([
      new Date().toISOString(), name || 'Unknown', taskId || '', taskDesc || '',
      field, (oldVal===undefined||oldVal===null)?'':oldVal, (newVal===undefined||newVal===null)?'':newVal
    ]);
  } catch(e) { Logger.log('logAutoActivity failed: ' + e.message); }
}

// Cross-checks manual EOD ("5. KPI Log") submissions against the auto-
// tracked activity for the same person/day. Flags any report that claims
// a day's work ('Report Submitted (Y/N)' = 'Y') but has zero matching
// auto-tracked task changes that day — i.e. nothing the system can verify
// actually happened. This never edits or blocks the EOD report itself,
// it just surfaces the mismatch to leaders.
function computeActivityFlags() {
  const kpi  = sheetToObjects('5. KPI Log').slice(-90);
  const auto = sheetToObjects('14. Auto Activity');
  const autoCounts = {};
  auto.forEach(a => {
    const d    = String(a.Timestamp || '').substring(0, 10);
    const name = String(a['Team Member'] || '').trim().toUpperCase();
    const key  = d + '|' + name;
    autoCounts[key] = (autoCounts[key] || 0) + 1;
  });
  return kpi.map(k => {
    const d       = String(k.Date || '').substring(0, 10);
    const name    = String(k['Team Member'] || '').trim().toUpperCase();
    const key     = d + '|' + name;
    const autoCount = autoCounts[key] || 0;
    const submitted = String(k['Report Submitted (Y/N)'] || '').trim().toUpperCase() === 'Y';
    return {
      Date: k.Date, 'Team Member': k['Team Member'], autoCount,
      flagged: submitted && autoCount === 0,
    };
  }).filter(f => f.flagged);
}

function readAll(token) {
  const identity      = resolveIdentity(token);
  const isLeaderCaller = !!(identity && identity.isLeader);
  // Create the tab on first leader load rather than waiting for the first
  // real task edit — so it's visible in the sheet immediately, not just
  // after someone changes a task's status/progress/proof link.
  if (isLeaderCaller) getOrCreateSheet('14. Auto Activity', TAB_HEADERS['14. Auto Activity']);
  return json({
    tasks      : sheetToObjects('1. Master Tasks').slice(0, 300),
    cos        : sheetToObjects('2. COS'),
    partners   : sheetToObjects('3. Partnerships'),
    orgs       : sheetToObjects('4. Org Database'),
    kpi        : sheetToObjects('5. KPI Log').slice(-60),
    marketing  : sheetToObjects('6. Marketing'),
    meetings   : sheetToObjects('7. Meetings'),
    checklists : sheetToObjects('17. Task Checklists'),
    automations: getAutomationsData(),
    // Toggle audit trail — who flipped which automation switch and when.
    // Leader-only: same reasoning as autoActivity/activityFlags above.
    automationLog: isLeaderCaller ? sheetToObjects('13. Automation Log').slice(-100).reverse() : [],
    // IsLeader is computed here, not stored on the sheet — LEADER_DISPLAY_NAMES
    // (above) is the one place leader status is defined. Exposing it per-row
    // means the frontend never has to keep its own duplicate leader list in
    // sync with this one.
    team       : sheetToObjects('9. Team Members').map(function(m) {
      m.IsLeader = isLeaderName(canonicalizeName(m.Name));
      return m;
    }),
    // Leader-only: raw auto-tracked activity feed and the mismatch flags
    // derived from it. Non-leaders never receive this in the payload at
    // all (not just hidden in the UI), same as other leader-gated data.
    autoActivity : isLeaderCaller ? sheetToObjects('14. Auto Activity').slice(-150).reverse() : [],
    activityFlags: isLeaderCaller ? computeActivityFlags() : [],
    meta       : { generatedAt: new Date().toISOString(), version: '3.2' },
  });
}

function readTab(tabName) {
  return json({ rows: sheetToObjects(tabName) });
}

function readFiltered(tabName, params) {
  let rows = sheetToObjects(tabName || '1. Master Tasks');
  if (params.status)   rows = rows.filter(r => r['Status'] === params.status || r['Current Status'] === params.status);
  if (params.owner)    rows = rows.filter(r => (r['Assigned To'] || '').toLowerCase().includes(params.owner.toLowerCase()));
  if (params.priority) rows = rows.filter(r => r['Priority'] === params.priority);
  if (params.dateFrom) {
    const from = new Date(params.dateFrom);
    rows = rows.filter(r => { const d = r['Deadline'] || r['Due Date'] || r['Follow-Up Date']; return d && new Date(d) >= from; });
  }
  if (params.dateTo) {
    const to = new Date(params.dateTo);
    rows = rows.filter(r => { const d = r['Deadline'] || r['Due Date'] || r['Follow-Up Date']; return d && new Date(d) <= to; });
  }
  return json({ rows, count: rows.length });
}

function getAutomationsData() {
  const sheet = getOrCreateSheet('8. Automation Controls', TAB_HEADERS['8. Automation Controls']);
  const defaults = [
    ['email_assignments',  'Task Assignment Emails',    'Email alert when a task is assigned to someone',                  'TRUE','',''],
    ['overdue_reminders',  'Overdue Task Reminders',    'Direct Gmail reminders at day 2, 5, 7 since last update',         'TRUE','',''],
    ['eod_nudges',         'EOD Report Nudges',         "Gmail summary of who hasn't filed EOD by 6 PM",                   'TRUE','',''],
    ['morning_digest',     'Daily Morning Digest',      'Personal task summary emailed to each member at 8am (Mon=full)',  'TRUE','',''],
    ['cob_reminders',      'COB EOD Reminders',         'Personal reminder to file EOD sent to unfiled members at 4pm',   'TRUE','',''],
    ['n8n_webhooks',       'n8n Webhook Triggers',      'Fire webhooks to n8n (optional — system works without n8n)',      'TRUE','',''],
    ['deadline_reminders', 'Deadline Email Reminders',  'Email assignees 3 days and 1 day before task deadline',           'TRUE','',''],
    ['calendar_sync',      'Google Calendar Sync',      'Auto-create calendar events when deadlines are set on tasks',     'TRUE','',''],
    ['dependency_checks',  'Dependency Status Checks',  'Auto-set tasks to Blocked when dependencies are incomplete',      'TRUE','',''],
    ['recurring_tasks',    'Recurring Task Generation', 'Auto-generate new task instances from recurring templates',       'TRUE','',''],
    ['meeting_notetaker',  'Meeting Notetaker → Tasks', 'Accept AI-extracted action items from meetings into the review queue', 'TRUE','',''],
  ];

  let rows = sheetToObjects('8. Automation Controls');

  // Was: only ever seeded this list when the sheet had ZERO rows total —
  // meaning any automation key added to the code AFTER the sheet already
  // had some rows from an earlier version never got backfilled. It kept
  // running (isAutomationEnabled defaults to true for an unknown key) but
  // with no row, no toggle, no visibility in this tab at all — which is
  // exactly the bug this fixes. Now backfills whichever specific keys are
  // actually missing, every time this loads, instead of an all-or-nothing
  // one-time seed.
  const existingKeys = new Set(rows.map(r => r['Key']));
  const missing = defaults.filter(d => !existingKeys.has(d[0]));
  if (missing.length) {
    missing.forEach(r => sheet.appendRow(r));
    rows = sheetToObjects('8. Automation Controls'); // re-read so the response includes the newly-added rows
  }

  return rows.map(r => ({ ...r, Enabled: r['Enabled'] === 'TRUE' || r['Enabled'] === true }));
}


// ═════════════════════════════════════
//  ASSIGNMENT EMAIL + CALENDAR MARKER
//  Called by appendRow when a task has an assignee email.
//  Calendar: all-day event on deadline date, no reminders = appears
//  as a date marker, not a blocking meeting event.
// ═════════════════════════════════════

// Formats a task's Deadline for display, including the time whenever
// Deadline_Time ("HH:MM", 24h) is set. If no time was set, says so
// explicitly instead of leaving a bare date — since addDeadlineToCalendar()
// silently defaults untimed deadlines to 9:00 AM, the email should say
// that too rather than let the assignee guess or assume end-of-day.
function formatDeadline(deadlineStr, deadlineTime) {
  if (!deadlineStr) return 'No deadline set';
  const dateLabel = new Date(deadlineStr).toDateString();
  if (deadlineTime && /^\d{1,2}:\d{2}/.test(String(deadlineTime))) {
    const [hStr, mStr] = String(deadlineTime).split(':');
    let h = parseInt(hStr, 10);
    const m = parseInt(mStr, 10);
    const ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12 || 12;
    return dateLabel + ' at ' + h + ':' + String(m).padStart(2, '0') + ' ' + ampm;
  }
  return dateLabel + ' (no time set — calendar defaults to 9:00 AM)';
}

// Returns true/false reflecting whether Resend actually accepted the email —
// callers that report back to a human (resendAssignmentEmail) need this;
// callers that fire-and-forget (appendRow/updateRow) can ignore it. Used to
// silently return undefined either way, which is indistinguishable from
// success to a caller doing `if (result)` — the bug behind resends that
// silently failed but still told the leader it went out.
function sendAssignmentEmail(task) {
  if (!task) { Logger.log('sendAssignmentEmail SKIPPED — task is undefined'); return false; }
  // 'Assigned To' / 'Assigned_To_Email' and 'Support Person' / 'Support_Person_Email'
  // can each now hold more than one person (comma-joined by the frontend's
  // multi-select). Everything below works off these arrays instead of a
  // single name/email so the notification actually fans out to everyone
  // on the task, not just the first person listed.
  const assigneeEmails = splitEmailList(task['Assigned_To_Email']);
  if (!assigneeEmails.length) {
    Logger.log('sendAssignmentEmail SKIPPED — no email on task: ' + (task['ID'] || '?'));
    return false;
  }
  Logger.log('sendAssignmentEmail — to: ' + assigneeEmails.join(', ') + ' task: ' + (task['ID'] || '?'));

  const assigneeNames = splitPeopleNames(task['Assigned To']);
  const assigneeDisplay = formatNameList(assigneeNames.length ? assigneeNames : ['Team member']);
  const taskDesc     = task['Task Description'] || '(no description)';
  const taskId       = task['ID']               || '—';
  const deadline     = formatDeadline(task['Deadline'], task['Deadline_Time']);
  const priority     = task['Priority']         || 'Normal';
  const dept         = task['Business / Dept']  || '—';
  const supportNames = splitPeopleNames(task['Support Person']);
  const supportDisplay = formatNameList(supportNames) || '—';
  const notes        = task['Notes']            || '';
  // Deep-link straight to this task instead of the generic homepage.
  const taskLink     = APP_LINK.replace(/\/$/, '') + '/?task=' + encodeURIComponent(taskId);
  // CC emails: prefer the explicit, editable "Support_Person_Email" field (set/overridden
  // in the app) — only fall back to the team map if it was left blank. Anyone
  // already in the "to" list (e.g. also on support) is dropped from CC so no
  // one gets the same email twice.
  const teamMap = getTeamEmailMap();
  const supportEmailsRaw = splitEmailList(task['Support_Person_Email']).length
    ? splitEmailList(task['Support_Person_Email'])
    : supportNames.map(n => teamMap[n]).filter(Boolean);
  const ccSupport = supportEmailsRaw.filter(e => e && !assigneeEmails.includes(e));
  // Supervisor is an oversight role, not a "doer" — CC'd for visibility like
  // Support Person, but not named in the greeting/intro (which is about who's
  // actually doing the work). Now genuinely multi-value, same as Assigned
  // To / Support Person: a task can have more than one supervisor.
  const supervisorNames = splitPeopleNames(task['Supervisor']);
  const supervisorDisplay = formatNameList(supervisorNames);
  const supervisorEmailsRaw = splitEmailList(task['Supervisor_Email']).length
    ? splitEmailList(task['Supervisor_Email'])
    : supervisorNames.map(n => teamMap[n]).filter(Boolean);
  const ccSupervisor = supervisorEmailsRaw.filter(e => e && !assigneeEmails.includes(e) && !ccSupport.includes(e));
  const ccAll = ccSupport.concat(ccSupervisor);
  const ccNoteParts = [];
  if (ccSupport.length) ccNoteParts.push(supportDisplay + ' is CC\'d as support');
  if (ccSupervisor.length) ccNoteParts.push(supervisorDisplay + ' is CC\'d as supervisor');
  const ccNotePlain = ccNoteParts.join('; ') + (ccNoteParts.length ? '.' : '');
  const ccNoteHtml  = ccNoteParts.length ? 'Note: ' + ccNoteParts.join('; ') + '.' : '';
  const greetingName = ccSupport.length ? (assigneeDisplay + ' &amp; ' + supportDisplay) : assigneeDisplay;
  const introLine = ccSupport.length
    ? 'A new task has been assigned to <strong>' + assigneeDisplay + '</strong>, with <strong>' + supportDisplay + '</strong> supporting.'
    : 'A new task has been assigned to <strong>' + assigneeDisplay + '</strong>.';
  const introLinePlain = ccSupport.length
    ? 'A new task has been assigned to ' + assigneeDisplay + ', with ' + supportDisplay + ' supporting.'
    : 'A new task has been assigned to ' + assigneeDisplay + '.';

  const subject = `[ABION] New Task Assigned: "${taskDesc}"`;

  var prioColor  = priority === 'URGENT' ? '#f05060' : priority === 'HIGH' ? '#f0a030' : priority === 'MEDIUM' ? '#f0c040' : '#8bc34a';
  var notesBorder = notes ? 'border-bottom:1px solid #f0f0f8;' : '';
  var notesRow    = notes ? '<tr><td style="padding:10px 14px;color:#666;font-size:12px">Notes</td><td style="padding:10px 14px">' + notes + '</td></tr>' : '';

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a2e;max-width:560px">' +
    '<div style="background:#5b52f0;padding:18px 24px;border-radius:8px 8px 0 0">' +
    '<span style="font-size:20px;font-weight:800;letter-spacing:4px;color:#fff">NEX<span style="color:#00d4a0">US</span></span>' +
    '<span style="font-size:12px;font-weight:700;color:rgba(255,255,255,0.85);margin-left:10px;font-family:monospace">ABION Industries</span>' +
    '</div>' +
    '<div style="background:#f8f8ff;border:1px solid #e0e0f0;border-top:none;border-radius:0 0 8px 8px;padding:24px">' +
    '<p style="margin:0 0 16px">Hi <strong>' + greetingName + '</strong>,</p>' +
    '<p style="margin:0 0 16px">' + introLine + '</p>' +
    '<table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #e0e0f0;border-radius:6px;margin-bottom:20px">' +
    '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px;width:110px">Task ID</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;font-weight:700;font-family:monospace">' + taskId + '</td></tr>' +
    '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px">Task</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;font-weight:700">' + taskDesc + '</td></tr>' +
    '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px">Priority</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8"><span style="background:' + prioColor + ';color:#fff;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700">' + priority + '</span></td></tr>' +
    '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px">Department</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8">' + dept + '</td></tr>' +
    '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px">Support Person</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8">' + supportDisplay + '</td></tr>' +
    (supervisorDisplay ? '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px">Supervisor</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8">' + supervisorDisplay + '</td></tr>' : '') +
    '<tr><td style="padding:10px 14px;' + notesBorder + 'color:#666;font-size:12px">Deadline</td><td style="padding:10px 14px;' + notesBorder + 'font-weight:700;color:#f05060">' + deadline + '</td></tr>' +
    notesRow +
    '</table>' +
    '<p style="margin:0 0 20px;font-size:13px;color:#555">Your deadline has been added to your Google Calendar.</p>' +
    '<p style="margin:0 0 8px;font-size:13px;color:#555">Open <a href="' + taskLink + '" style="color:#5b52f0;font-weight:700;text-decoration:none">this task in NEXUS</a> to view and update it.</p>' +
    '<p style="margin:0 0 12px;font-size:13px;color:#555">Please acknowledge by updating the task status to <strong>In Progress</strong> once you begin.</p>' +
    '<p style="margin:0 0 20px;font-size:12px;color:#888">Need clarity on scope, timing, or anything else? Reply in the NEXUS Chat thread for this task rather than over email.</p>' +
    '<p style="margin:0 0 20px;font-size:11px;color:#999;font-style:italic">' + ccNoteHtml + '</p>' +
    '<hr style="border:none;border-top:1px solid #e0e0f0;margin:20px 0">' +
    '<p style="margin:0;font-size:11px;color:#999;font-family:monospace">— NEXUS Automation · ABION Industries · Task ' + taskId + '</p>' +
    '</div></div>';

  var plainBody =
    'Hi ' + greetingName.replace('&amp;', '&') + ',\n\n' + introLinePlain + '\n\n' +
    'Task ID:       ' + taskId + '\nTask:          ' + taskDesc + '\nPriority:      ' + priority + '\n' +
    'Department:    ' + dept + '\nSupport Person:' + supportDisplay + '\n' +
    (supervisorDisplay ? 'Supervisor:    ' + supervisorDisplay + '\n' : '') +
    'Deadline:      ' + deadline + '\n' +
    (notes ? 'Notes:         ' + notes + '\n' : '') +
    '\nOpen this task in NEXUS: ' + taskLink +
    '\nAcknowledge by moving the task to "In Progress" once you begin.' +
    '\nQuestions? Reply in NEXUS Chat for this task rather than over email.' +
    (ccNotePlain ? '\n\n(' + ccNotePlain + ')' : '') +
    '\n\n— NEXUS Automation · ABION Industries · Task ' + taskId;

  // Assignment notifications are forced to Both (email + in-app) by
  // default — a team member can no longer opt themselves out of being told
  // they've been assigned something via their general Notify_Pref. Only a
  // leader can flip a specific person back to "follow their own
  // preference" from the Team Directory (assignmentWantsEmail /
  // assignmentForcedBoth_ above). assigneeEmails/assigneeNames are
  // parallel arrays, same convention as everywhere else in this function.
  const emailTargets = assigneeEmails.filter((email, i) => assignmentWantsEmail(assigneeNames[i] || assigneeNames[0]));
  const ccFiltered = ccAll.filter(email => {
    const supportIdx = supportEmailsRaw.indexOf(email);
    if (supportIdx !== -1) return assignmentWantsEmail(supportNames[supportIdx] || supportNames[0]);
    const supervisorIdx = supervisorEmailsRaw.indexOf(email);
    if (supervisorIdx !== -1) return assignmentWantsEmail(supervisorNames[supervisorIdx] || supervisorNames[0]);
    return true;
  });
  const delivered = emailTargets.length
    ? sendViaResend(emailTargets, subject, htmlBody, plainBody, { replyTo: FOUNDER_EMAIL, cc: ccFiltered })
    : false;

  // In-app notification — everyone with a stake in this task (assignees,
  // support, supervisors) gets a bell-icon entry deep-linked to it,
  // independent of whether they also got the email above. Anyone still
  // forced-Both (the default) bypasses their personal in-app gate too.
  const assignmentStakeholders = assigneeNames.concat(supportNames, supervisorNames);
  const forcedBothNames = assignmentStakeholders.filter(n => assignmentForcedBoth_(n));
  notifyInApp(assignmentStakeholders, {
    type: 'assignment',
    title: 'NEXUS · New Task Assigned',
    body: taskDesc.substring(0, 100),
    url: taskLink,
    tag: 'nexus-assignment-' + taskId,
    taskId: taskId,
  }, { force: forcedBothNames });

  // guests accepts a comma-separated string of every assignee's email, so
  // everyone assigned gets the calendar invite, not just the first one.
  // Calendar entry is independent of email delivery — still create/update it
  // even if Resend failed, since the deadline itself is still real.
  if (task['Deadline']) addDeadlineToCalendar(task, assigneeEmails);

  return delivered;
}

// Writes a calendar event ID back onto the task's row in '1. Master Tasks'
// (Calendar_Event_ID column), so the NEXT call for this same task can find
// and replace it instead of creating a duplicate. Best-effort — a failure
// here just means the next edit creates a fresh event instead of reusing
// this one, never a hard error for the caller.
function saveCalendarEventIdForTask(taskId, eventId) {
  try {
    const sheet = SS.getSheetByName('1. Master Tasks');
    if (!sheet) return;
    const values  = sheet.getDataRange().getValues();
    const headers = values[0];
    const idCol   = headers.indexOf('ID');
    const eventCol = headers.indexOf('Calendar_Event_ID');
    if (idCol === -1 || eventCol === -1) { Logger.log('saveCalendarEventIdForTask — missing ID or Calendar_Event_ID column'); return; }
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol]).trim() === String(taskId).trim()) {
        sheet.getRange(i + 1, eventCol + 1).setValue(eventId || '');
        return;
      }
    }
    Logger.log('saveCalendarEventIdForTask — task row not found for id ' + taskId);
  } catch(e) { Logger.log('saveCalendarEventIdForTask error: ' + e.message); }
}

function addDeadlineToCalendar(task, assigneeEmails) {
  // Creates a 30-min event on the deadline day, at the time specified on the
  // task (Deadline_Time, "HH:MM") — or 9:00 AM by default if that field was
  // left untouched.
  // Colour: RED (Tomato) so it stands out clearly on the calendar.
  // Reminders: morning-of at event start (0 min offset) + day-before at 9am.
  //
  // Created on the script owner's calendar (this is a container-bound
  // script — CalendarApp.getDefaultCalendar() is always the account running
  // the deployment), with every assignee added as a guest so the event lands
  // on each of their own calendars too — same as before. What changed:
  // sendInvites is now false. Google Calendar's own guest-invite email is
  // generic (not NEXUS-branded) and links to the calendar event / the raw
  // sheet, not the task in NEXUS — so it's no longer relied on for
  // notification at all. sendDeadlineEmail() below sends our own branded
  // email with a deep link instead; CalendarApp here is purely for the
  // calendar entry itself.
  // Caveat worth knowing: without sendInvites, guests aren't emailed by
  // Calendar, but they're still added as guests on the event, so it still
  // shows up on their calendar the normal way a shared invite does (subject
  // to their own account's "automatically add invitations" setting) — this
  // isn't a way to silently write into someone else's calendar without them
  // being a guest on the event at all, which isn't something a container-
  // bound script can do without domain-wide delegation.
  //
  // Reassignment / repeated-edit fix: this used to call createEvent() every
  // single time — reassigning a task, or just editing its deadline again,
  // stacked a brand-new event on top of the old one forever. The old
  // assignee was never removed as a guest and the calendar accumulated
  // duplicates. Now the task's prior event (tracked via Calendar_Event_ID)
  // is deleted first, so there's always exactly one live calendar event per
  // task, with the guest list matching whoever is actually assigned now.
  const taskDesc = task['Task Description'] || '(no description)';
  const taskId   = task['ID'] || '—';
  const taskLink = APP_LINK.replace(/\/$/, '') + '/?task=' + encodeURIComponent(taskId);
  const calendar = CalendarApp.getDefaultCalendar();

  // Compute new time values once — used by both the update and create paths.
  const d = new Date(task['Deadline']);
  let hh = 9, mm = 0;
  const deadlineTime = task['Deadline_Time'];
  if (deadlineTime && /^\d{1,2}:\d{2}/.test(String(deadlineTime))) {
    const parts = String(deadlineTime).split(':');
    hh = parseInt(parts[0], 10); mm = parseInt(parts[1], 10);
  }
  d.setHours(hh, mm, 0, 0);
  const end   = new Date(d.getTime() + 30 * 60000);
  const title = `🔴 [ABION] DEADLINE: ${taskDesc}`;
  const desc  = `Task ID: ${taskId}\nThis task is due today.\nOpen it in NEXUS: ${taskLink}`;

  // If a prior event exists, edit it in place rather than delete + recreate.
  // This preserves the event identity (guests keep their RSVP state, the
  // calendar entry doesn't flicker away and back).  Event ID doesn't change
  // so there's nothing new to persist — just return.
  const priorEventId = task['Calendar_Event_ID'];
  if (priorEventId) {
    try {
      const priorEvent = calendar.getEventById(priorEventId);
      if (priorEvent) {
        priorEvent.setTitle(title);
        priorEvent.setTime(d, end);
        priorEvent.setDescription(desc);
        // Sync guests to the current assignee list — this used to only touch
        // title/time/description, so a task reassigned to different people
        // (without the deadline itself also changing) left the event's guest
        // list frozen on whoever was assigned when the event was first
        // created: new assignees never got added, old ones never removed.
        const desiredGuests = (assigneeEmails || []).filter(Boolean);
        const currentGuests = priorEvent.getGuestList().map(g => g.getEmail());
        currentGuests.filter(e => !desiredGuests.includes(e)).forEach(e => {
          try { priorEvent.removeGuest(e); } catch(e2) { Logger.log('Could not remove guest ' + e + ' from event ' + priorEventId + ': ' + e2.message); }
        });
        desiredGuests.filter(e => !currentGuests.includes(e)).forEach(e => {
          try { priorEvent.addGuest(e); } catch(e2) { Logger.log('Could not add guest ' + e + ' to event ' + priorEventId + ': ' + e2.message); }
        });
        Logger.log('Updated calendar event ' + priorEventId + ' for task ' + taskId);
        return; // ID unchanged — no need to write back to the sheet
      }
    } catch(e) { Logger.log('Could not update prior calendar event for task ' + taskId + ' — will create fresh: ' + e.message); }
  }

  // No prior event (first deadline set, or prior event was externally deleted)
  // — create a fresh one and persist its ID.
  try {
    const event = calendar.createEvent(title, d, end, {
      description: desc,
      guests: (assigneeEmails || []).join(','),
      sendInvites: false,
    });
    // Red/Tomato colour so it pops on the calendar
    event.setColor(CalendarApp.EventColor.RED);
    // Remind at event time (0 min) and 1 day before (1440 min)
    event.addPopupReminder(0);
    event.addPopupReminder(1440);
    event.addEmailReminder(1440); // also an email the day before
    saveCalendarEventIdForTask(taskId, event.getId());
    Logger.log('Deadline calendar event created for ' + (assigneeEmails || []).join(', '));
  } catch(e) { Logger.log('Calendar event error: ' + e.message); }
}

// Branded, task-linked replacement for Calendar's own invite email — reuses
// the same NEXUS header/footer look and sendViaResend plumbing as
// sendAssignmentEmail, just with deadline-specific copy. CCs support person
// and supervisor the same way sendAssignmentEmail does, since anyone kept in
// the loop on the assignment should be kept in the loop on its deadline too.
function sendDeadlineEmail(task, assigneeEmails) {
  assigneeEmails = (assigneeEmails && assigneeEmails.length) ? assigneeEmails : splitEmailList(task['Assigned_To_Email']);
  if (!assigneeEmails.length) { Logger.log('sendDeadlineEmail SKIPPED — no assignee email on task: ' + (task['ID']||'?')); return; }

  const assigneeNames = splitPeopleNames(task['Assigned To']);
  const assigneeDisplay = formatNameList(assigneeNames.length ? assigneeNames : ['Team member']);
  const taskDesc  = task['Task Description'] || '(no description)';
  const taskId    = task['ID'] || '—';
  const priority  = task['Priority'] || 'Normal';
  const deadline  = formatDeadline(task['Deadline'], task['Deadline_Time']);
  const taskLink  = APP_LINK.replace(/\/$/, '') + '/?task=' + encodeURIComponent(taskId);

  const teamMap = getTeamEmailMap();
  const supportNames = splitPeopleNames(task['Support Person']);
  const supportEmailsRaw = splitEmailList(task['Support_Person_Email']).length
    ? splitEmailList(task['Support_Person_Email'])
    : supportNames.map(n => teamMap[n]).filter(Boolean);
  const ccSupport = supportEmailsRaw.filter(e => e && !assigneeEmails.includes(e));
  const supervisorName  = String(task['Supervisor'] || '').trim();
  const supervisorEmail = task['Supervisor_Email'] || (supervisorName ? teamMap[supervisorName] : '') || '';
  const ccSupervisor = (supervisorEmail && !assigneeEmails.includes(supervisorEmail) && !ccSupport.includes(supervisorEmail)) ? [supervisorEmail] : [];
  const ccAll = ccSupport.concat(ccSupervisor);

  const subject = `[ABION] Deadline Reminder: "${taskDesc}" — ${deadline}`;
  var prioColor = priority === 'URGENT' ? '#f05060' : priority === 'HIGH' ? '#f0a030' : priority === 'MEDIUM' ? '#f0c040' : '#8bc34a';

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a2e;max-width:560px">' +
    '<div style="background:#5b52f0;padding:18px 24px;border-radius:8px 8px 0 0">' +
    '<span style="font-size:20px;font-weight:800;letter-spacing:4px;color:#fff">NEX<span style="color:#00d4a0">US</span></span>' +
    '<span style="font-size:12px;font-weight:700;color:rgba(255,255,255,0.85);margin-left:10px;font-family:monospace">ABION Industries</span>' +
    '</div>' +
    '<div style="background:#f8f8ff;border:1px solid #e0e0f0;border-top:none;border-radius:0 0 8px 8px;padding:24px">' +
    '<p style="margin:0 0 16px">Hi <strong>' + assigneeDisplay + '</strong>,</p>' +
    '<p style="margin:0 0 16px">A deadline has been set for your task, due <strong>' + deadline + '</strong>. It has also been added to your Google Calendar.</p>' +
    '<table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #e0e0f0;border-radius:6px;margin-bottom:20px">' +
    '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px;width:110px">Task ID</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;font-weight:700;font-family:monospace">' + taskId + '</td></tr>' +
    '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px">Task</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;font-weight:700">' + taskDesc + '</td></tr>' +
    '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8;color:#666;font-size:12px">Priority</td><td style="padding:10px 14px;border-bottom:1px solid #f0f0f8"><span style="background:' + prioColor + ';color:#fff;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700">' + priority + '</span></td></tr>' +
    '<tr><td style="padding:10px 14px;color:#666;font-size:12px">Deadline</td><td style="padding:10px 14px;font-weight:700;color:#f05060">' + deadline + '</td></tr>' +
    '</table>' +
    '<p style="margin:0 0 8px;font-size:13px;color:#555">Open <a href="' + taskLink + '" style="color:#5b52f0;font-weight:700;text-decoration:none">this task in NEXUS</a> to view and update it.</p>' +
    '<p style="margin:0 0 20px;font-size:12px;color:#888">Need more time, or already done? Reply in the NEXUS Chat thread for this task rather than over email.</p>' +
    (ccAll.length ? '<p style="margin:0 0 20px;font-size:11px;color:#999;font-style:italic">Note: ' +
      [ccSupport.length ? supportNames.join(', ') + ' (support)' : '', ccSupervisor.length ? supervisorName + ' (supervisor)' : '']
        .filter(Boolean).join('; ') + ' also CC\'d on this email.</p>' : '') +
    '<hr style="border:none;border-top:1px solid #e0e0f0;margin:20px 0">' +
    '<p style="margin:0;font-size:11px;color:#999;font-family:monospace">— NEXUS Automation · ABION Industries · Task ' + taskId + '</p>' +
    '</div></div>';

  var plainBody =
    'Hi ' + assigneeDisplay + ',\n\n' +
    'A deadline has been set for your task, due ' + deadline + '. It has also been added to your Google Calendar.\n\n' +
    'Task ID:  ' + taskId + '\nTask:     ' + taskDesc + '\nPriority: ' + priority + '\nDeadline: ' + deadline + '\n\n' +
    'Open this task in NEXUS: ' + taskLink +
    '\nNeed more time, or already done? Reply in NEXUS Chat for this task rather than over email.' +
    '\n\n— NEXUS Automation · ABION Industries · Task ' + taskId;

  const emailTargets = assigneeEmails.filter((email, i) => wantsEmail(assigneeNames[i] || assigneeNames[0]));
  if (emailTargets.length) {
    sendViaResend(emailTargets, subject, htmlBody, plainBody, {
      replyTo: FOUNDER_EMAIL,
      cc: ccAll,
    });
  }
  Logger.log('sendDeadlineEmail — to: ' + assigneeEmails.join(', ') + (ccAll.length ? ' | cc: ' + ccAll.join(', ') : '') + ' task: ' + taskId);

  notifyInApp(assigneeNames, {
    type: 'deadline_reminder',
    title: 'NEXUS · Deadline Reminder',
    body: taskDesc.substring(0, 90) + ' — due ' + deadline,
    url: taskLink,
    tag: 'nexus-deadline-' + taskId,
    taskId: taskId,
  });
}

// ═════════════════════════════════════
//  MONDAY DIGEST
//  Each assignee gets all their active tasks — title, status, deadline, notes.
// ═════════════════════════════════════

function sendMondayDigest() {
  if (!isAutomationEnabled('overdue_reminders')) return;
  const today = new Date();
  if (today.getDay() !== 1) return; // Mondays only

  const tasks = sheetToObjects('1. Master Tasks');
  const skip  = ['Done','Cancelled'];
  const byAssignee = {};

  tasks.forEach(task => {
    if (skip.includes(task['Status'])) return;
    if (!task['Assigned_To_Email'])     return;
    const email = task['Assigned_To_Email'];
    if (!byAssignee[email]) byAssignee[email] = { name: task['Assigned To'] || email, tasks: [] };
    byAssignee[email].tasks.push(task);
  });

  Object.entries(byAssignee).forEach(([email, { name, tasks: myTasks }]) => {
    const overdue  = myTasks.filter(t => t['Deadline'] && new Date(t['Deadline']) < today);
    const urgent   = myTasks.filter(t => ['URGENT','HIGH'].includes(t['Priority']));
    const upcoming = myTasks.filter(t => {
      if (!t['Deadline']) return false;
      const diff = Math.ceil((new Date(t['Deadline']) - today) / 86400000);
      return diff >= 0 && diff <= 7;
    });

    const taskLines = myTasks.map((t, i) => {
      const isOverdue = t['Deadline'] && new Date(t['Deadline']) < today;
      const dl = t['Deadline'] ? new Date(t['Deadline']).toDateString() : 'No deadline';
      return `${i+1}. ${t['Task Description']}\n` +
             `   Status: ${t['Status']} | Progress: ${t['% Progress']||0}%\n` +
             `   Deadline: ${dl}${isOverdue ? ' ⚠ OVERDUE' : ''}\n` +
             (t['Notes'] ? `   Notes: ${t['Notes']}\n` : '');
    }).join('\n');

    const subject = `[ABION] Your Week Ahead — ${myTasks.length} Active Task(s) · ${today.toDateString()}`;
    const body =
      `Good morning ${name},\n\n` +
      `Here's your task summary for the week:\n\n` +
      `📋 ACTIVE TASKS (${myTasks.length})\n` +
      `━━━━━━━━━━━━━━━━━━━━\n` +
      taskLines +
      `━━━━━━━━━━━━━━━━━━━━\n\n` +
      (overdue.length  ? `⚠️  OVERDUE: ${overdue.length} task(s) past deadline — prioritise these.\n` : '') +
      (urgent.length   ? `🔴 URGENT/HIGH: ${urgent.length} task(s) need your attention.\n` : '') +
      (upcoming.length ? `🔔 DUE THIS WEEK: ${upcoming.length} task(s) due within 7 days.\n` : '') +
      `\nUpdate your tasks in NEXUS:\n${APP_LINK}\n\n` +
      `Have a productive week!\n\n— NEXUS Automation · ABION Industries`;

    try {
      emailIfWantedDigest_(name, email, subject, null, body, { replyTo: FOUNDER_EMAIL });
      Logger.log('Monday digest — to: ' + email + (wantsDigestEmail(name) ? ' (sent)' : ' (skipped per Digest_Notify_Pref)'));
    } catch(e) { Logger.log('Monday digest error for ' + email + ': ' + e.message); }
    notifyInApp([name], {
      type: 'digest',
      title: 'NEXUS · Your Week Ahead',
      body: myTasks.length + ' active task(s)' + (overdue.length ? ', ' + overdue.length + ' overdue' : ''),
      url: notifUrl_({ notif: 1 }),
      tag: 'nexus-monday-digest-' + localDateISO(),
    }, { gate: wantsDigestInApp });
  });
}

// ═════════════════════════════════════
//  DAILY MORNING DIGEST
//  Runs Mon–Fri at 8am. Each assignee gets their active tasks.
//  Monday also includes the fuller weekly summary (sendMondayDigest handles that).
// ═════════════════════════════════════

function sendDailyMorningDigest() {
  if (!isAutomationEnabled('morning_digest')) return;
  const today = new Date();
  const day   = today.getDay();
  if (day === 0 || day === 6) return; // Skip weekends
  if (day === 1) { sendMondayDigest(); return; } // Monday → use the richer Monday digest

  const dayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const tasks  = sheetToObjects('1. Master Tasks');
  const skip   = ['Done','Cancelled'];
  const byAssignee = {};

  tasks.forEach(task => {
    if (skip.includes(task['Status'])) return;
    const email = task['Assigned_To_Email'] || getTeamEmailMap()[task['Assigned To']] || '';
    if (!email) return;
    if (!byAssignee[email]) byAssignee[email] = { name: task['Assigned To'] || email, tasks: [] };
    byAssignee[email].tasks.push(task);
  });

  Object.entries(byAssignee).forEach(([email, { name, tasks: myTasks }]) => {
    const overdue  = myTasks.filter(t => t['Deadline'] && new Date(t['Deadline']) < today);
    const urgent   = myTasks.filter(t => ['URGENT','HIGH'].includes(t['Priority']));

    const taskLines = myTasks.map((t, i) => {
      const dl = t['Deadline'] ? new Date(t['Deadline']).toDateString() : 'No deadline';
      const flag = t['Deadline'] && new Date(t['Deadline']) < today ? ' ⚠ OVERDUE' : '';
      return `${i+1}. ${t['Task Description']}\n` +
             `   Status: ${t['Status']} | ${t['% Progress']||0}% | Due: ${dl}${flag}`;
    }).join('\n');

    const subject = `[ABION] Good ${dayNames[day]} ${name.split(' ')[0]} — ${myTasks.length} active task(s)`;
    const body =
      `Good morning ${name},\n\n` +
      `Here are your active tasks for today (${today.toDateString()}):\n\n` +
      `━━━━━━━━━━━━━━━━━━━━\n` +
      taskLines + '\n' +
      `━━━━━━━━━━━━━━━━━━━━\n\n` +
      (overdue.length ? `⚠️  ${overdue.length} task(s) OVERDUE — please update or escalate.\n` : '') +
      (urgent.length  ? `🔴 ${urgent.length} URGENT/HIGH priority task(s) need focus today.\n` : '') +
      `\nUpdate your tasks: ${APP_LINK}\n\n— NEXUS Automation · ABION Industries`;

    try {
      emailIfWantedDigest_(name, email, subject, null, body, { replyTo: FOUNDER_EMAIL });
      Logger.log('Morning digest — to: ' + email + (wantsDigestEmail(name) ? ' (sent)' : ' (skipped per Digest_Notify_Pref)'));
    } catch(e) { Logger.log('Morning digest error for ' + email + ': ' + e.message); }
    notifyInApp([name], {
      type: 'digest',
      title: 'NEXUS · Good ' + dayNames[day],
      body: myTasks.length + ' active task(s)' + (overdue.length ? ', ' + overdue.length + ' overdue' : ''),
      url: notifUrl_({ notif: 1 }),
      tag: 'nexus-morning-digest-' + localDateISO(),
    }, { gate: wantsDigestInApp });
  });
}

// ═════════════════════════════════════
//  COB / EOD PERSONAL REMINDER
//  Fires at 4:30 PM (set trigger to hour=16, then filter for after :30 is not possible
//  in Apps Script time-based triggers — so set to 4 PM and accept ±30 min window).
//  Sends each team member a personal reminder to submit their EOD before close of day.
//  Only fires if they haven't filed yet today.
// ═════════════════════════════════════

function sendCOBReminders() {
  if (!isAutomationEnabled('cob_reminders')) return;
  const today = new Date();
  const day   = today.getDay();
  if (day === 0 || day === 6) return; // Skip weekends

  const todayStr = localDateISO();
  const kpi      = sheetToObjects('5. KPI Log');
  const filed    = new Set(
    kpi.filter(k => k['Date'] === todayStr).map(k => k['Team Member'])
  );

  const expected = getMembersExpectedToday();
  const pending  = expected.filter(m => !filed.has(m));
  if (!pending.length) return;

  pending.forEach(memberName => {
    const email = getTeamEmailMap()[memberName];
    if (!email) return;
    const subject = `[ABION] ⏰ Don't forget your EOD report — ${today.toDateString()}`;
    const body =
      `Hi ${memberName.split(' ')[0]},\n\n` +
      `Just a reminder to submit your End-of-Day report before close of business today.\n\n` +
      `It takes 2 minutes and helps the whole team stay aligned.\n\n` +
      `➜ Submit here: ${APP_LINK}\n\n` +
      `If you're done for the day, hit the EOD tab in NEXUS and fill in:\n` +
      `  • What you completed\n` +
      `  • Key win / achievement\n` +
      `  • Any blockers\n` +
      `  • Plan for tomorrow\n\n` +
      `— NEXUS Automation · ABION Industries`;
    try {
      emailIfWanted_(memberName, email, subject, null, body, { replyTo: FOUNDER_EMAIL });
      Logger.log('COB reminder — to: ' + email + (wantsEmail(memberName) ? ' (sent)' : ' (skipped — in-app only)'));
    } catch(e) { Logger.log('COB reminder error for ' + email + ': ' + e.message); }
    notifyInApp([memberName], {
      type: 'eod_nudge',
      title: 'NEXUS · EOD Report Due',
      body: "Don't forget to submit today's End-of-Day report.",
      url: notifUrl_({ notif: 1 }),
      tag: 'nexus-eod-' + todayStr,
    });
  });
}

// ═════════════════════════════════════
//  HOURLY FOLLOW-UP FRAMEWORK
//  Sheet column: 'Follow_Up_Hours' on Master Tasks (e.g. value = 2 means ping every 2h).
//  This trigger fires every hour and checks tasks flagged for hourly follow-up.
//  Installed automatically by installTriggers() — no separate manual step needed.
//  HOW TO USE:
//    1. Add column 'Follow_Up_Hours' to Master Tasks sheet (a number, e.g. 2) — already
//       present in this form's schema, just set a value on the relevant task.
//    2. The function below checks if NOW mod interval ≈ 0 before sending.
//  NOTE: Apps Script hourly triggers fire at random minutes within the hour.
//        The check uses a ±15 min window around the interval boundary.
// ═════════════════════════════════════

function runHourlyFollowUps() {
  if (!isAutomationEnabled('overdue_reminders')) return;
  const now   = new Date();
  const day   = now.getDay();
  if (day === 0 || day === 6) return; // weekends off
  const h     = now.getHours();
  if (h < 8 || h > 19) return; // only during working hours 8am–7pm

  const tasks = sheetToObjects('1. Master Tasks');
  const skip  = ['Done','Cancelled'];

  tasks.forEach(task => {
    if (skip.includes(task['Status'])) return;

    // Check for Follow_Up_Hours flag
    const intervalHours = parseFloat(task['Follow_Up_Hours'] || '');
    if (!intervalHours || isNaN(intervalHours) || intervalHours <= 0) return;

    const email = task['Assigned_To_Email'] || getTeamEmailMap()[task['Assigned To']] || '';
    if (!email) return;

    const lastUpdated = task['Last_Updated'] ? new Date(task['Last_Updated']) : null;
    if (!lastUpdated) return;

    const hoursSince = (now - lastUpdated) / 3600000;
    // Check if we're within a ±0.4h (±24min) window of a follow-up interval boundary
    const remainder = hoursSince % intervalHours;
    const nearBoundary = remainder < 0.4 || remainder > (intervalHours - 0.4);
    if (!nearBoundary) return;

    const subject = `[ABION] ⏱ ${Math.round(hoursSince)}h check-in: "${task['Task Description']}"`;
    const body =
      `Hi ${task['Assigned To'] || 'Team member'},\n\n` +
      `This task is flagged for ${intervalHours}h follow-up intervals.\n\n` +
      `━━━━━━━━━━━━━━━━━━━━\n` +
      `Task:     ${task['Task Description']}\n` +
      `Status:   ${task['Status']}\n` +
      `Progress: ${task['% Progress']||0}%\n` +
      `Deadline: ${task['Deadline'] || 'No deadline'}\n` +
      `Last update: ${lastUpdated.toLocaleString()}\n` +
      `━━━━━━━━━━━━━━━━━━━━\n\n` +
      `Please update the status or progress now:\n${APP_LINK}\n\n` +
      `— NEXUS Automation · ABION Industries`;
    try {
      sendViaResend(email, subject, null, body, {
        replyTo: FOUNDER_EMAIL,
        cc: [FOUNDER_EMAIL].filter(e => e && e !== email),
      });
      Logger.log('Hourly follow-up sent to ' + email + ' for task ' + task['ID']);
    } catch(e) { Logger.log('Hourly follow-up error: ' + e.message); }
  });
}

// Standalone installer for runHourlyFollowUps, kept as a manual fallback —
// installTriggers() now creates this trigger automatically as part of its
// normal set, so you shouldn't need to run this separately anymore. Only
// useful if you want the hourly trigger without re-running the full
// installTriggers() (which resets every other trigger too).
function installHourlyTrigger() {
  // Remove any existing hourly follow-up triggers first
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runHourlyFollowUps') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runHourlyFollowUps').timeBased().everyHours(1).create();
  Logger.log('Hourly follow-up trigger installed — runs runHourlyFollowUps() every 60 min.');
  Logger.log('Add "Follow_Up_Hours" column to Master Tasks sheet, set to 2 (or any number) on relevant tasks.');
}



function appendRow(tabName, data) {
  const sheet   = getOrCreateSheet(tabName, TAB_HEADERS[tabName] || []);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Reject a manually-typed ID (e.g. the Task ID / Content ID field) that's
  // already used by another row in this tab. IDs are how tasks reference
  // each other (Depends On, chat discussion threads, calendar sync) — a
  // silent duplicate would mean two unrelated tasks answer to the same ID.
  // This only guards CREATING a new row; editing an existing task (moving
  // it through Not Started → In Progress → Done, etc.) goes through
  // updateRow, not here, so that's unaffected.
  const idColCheck = headers.indexOf('ID');
  const manualId   = idColCheck >= 0 ? String(data['ID'] || '').trim() : '';
  if (idColCheck >= 0 && manualId && sheet.getLastRow() > 1) {
    const existingIds = sheet.getRange(2, idColCheck + 1, sheet.getLastRow() - 1, 1)
      .getValues().map(r => String(r[0]).trim());
    if (existingIds.includes(manualId)) {
      return json({ error: 'ID "' + manualId + '" is already in use by another row — pick a different ID, or open that existing item to edit it instead.' });
    }
  }

  if (headers.includes('Last_Updated'))                         data['Last_Updated'] = new Date().toISOString();
  if (headers.includes('Date Added') && !data['Date Added'])    data['Date Added']   = new Date().toISOString().split('T')[0];
  // % Progress is a pure, read-only function of Status — ignore any
  // client-sent value outright (the frontend no longer has an editable
  // input for it, but don't trust that blindly) and compute it fresh from
  // whichever field drives this tab's ladder. FROZEN statuses (no ladder
  // entry) default new rows to 0 rather than leaving % Progress blank.
  const progField = PROGRESS_STATUS_FIELD[tabName];
  if (progField && headers.includes('% Progress')) {
    const computed = data[progField] ? statusToProgress(tabName, data[progField]) : null;
    data['% Progress'] = computed !== null ? computed : 0;
  }

  const row   = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');
  const idCol = headers.indexOf('ID');
  if (idCol >= 0 && !row[idCol]) row[idCol] = tabName.split('.')[0].trim() + '-' + Date.now();

  sheet.appendRow(row);

  // Send assignment email + calendar marker for new tasks only
  // If Assigned_To_Email missing, attempt server-side lookup via getTeamEmailMap()
  if (tabName === '1. Master Tasks' && isAutomationEnabled('email_assignments')) {
    Logger.log('appendRow email check — Assigned To: "' + (data['Assigned To']||'') + '" | Email: "' + (data['Assigned_To_Email']||'') + '"');
    if (!data['Assigned_To_Email'] && data['Assigned To']) {
      data['Assigned_To_Email'] = getTeamEmailMap()[data['Assigned To']] || '';
      Logger.log('Email resolved from map: "' + data['Assigned_To_Email'] + '"');
    }
    if (data['Assigned_To_Email']) {
      // Re-attach the auto-generated ID so the email shows the correct task ID
      if (!data['ID'] && row[idCol >= 0 ? idCol : 0]) data['ID'] = row[idCol >= 0 ? idCol : 0];
      Logger.log('Calling sendAssignmentEmail to: ' + data['Assigned_To_Email']);
      sendAssignmentEmail(data);
    } else {
      Logger.log('sendAssignmentEmail SKIPPED — email still empty after map lookup for: "' + (data['Assigned To']||'') + '"');
      // Notify founder so they can manually resend once email is set
      try {
        sendViaResend(FOUNDER_EMAIL,
          '[NEXUS] Task assigned but no email found: ' + (data['Task Description']||''),
          null,
          'A new task was created but no email could be resolved for "' + (data['Assigned To']||'unknown') + '".\n\n' +
          'Task: ' + (data['Task Description']||'--') + '\nID: ' + (data['ID']||'--') + '\n\n' +
          'Add the email to the task in the sheet, then use Resend Assignment from NEXUS.\n\n' + SHEET_LINK
        );
      } catch(ge) { Logger.log('Founder notify error: ' + ge.message); }
    }
  } else {
    Logger.log('appendRow email SKIPPED — tab: ' + tabName + ' | automation enabled: ' + isAutomationEnabled('email_assignments'));
  }

  // Auto-post into the task's own chat thread the moment it's created with
  // an assignee — so even if the assignment email/push gets missed, opening
  // the chat itself shows a "Task Assigned" message, with Acknowledge / In
  // Progress buttons right on it (frontend renders these off MsgType).
  if (tabName === '1. Master Tasks' && data['Assigned To']) {
    const newId = row[idCol >= 0 ? idCol : 0];
    const shortDesc = String(data['Task Description'] || '').substring(0, 60);
    postSystemTaskMessage(
      newId, shortDesc,
      'Task assigned to ' + data['Assigned To'] + ': ' + (data['Task Description'] || ''),
      [data['Assigned To']],
      'task_assigned'
    );
  }

  return json({ success: true, id: row[idCol >= 0 ? idCol : 0] });
}

// ═════════════════════════════════════
//  WRITE — UPDATE ROW
// ═════════════════════════════════════

// System-generated message into a task's chat thread — used for automated
// pings (e.g. acknowledgment) where an actual person didn't type anything.
// Deliberately does NOT use GmailApp — this only writes to the Chat
// Messages sheet and fires a push notification (UrlFetchApp to Vercel),
// so it costs zero Gmail send quota. `recipientNames` controls who gets
// pushed; the message itself is visible to anyone who opens the thread.
function postSystemTaskMessage(taskId, taskTitle, message, recipientNames, msgType) {
  try {
    const sheet = getOrCreateSheet('10. Chat Messages', TAB_HEADERS['10. Chat Messages']);
    const id = 'MSG-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
    const row = {
      ID: id, Channel: 'task', TaskID: String(taskId || ''),
      Author: 'NEXUS Automation', AuthorEmail: '',
      Message: String(message).substring(0, 2000),
      Timestamp: new Date().toISOString(),
      MsgType: msgType || '',
    };
    const headers = TAB_HEADERS['10. Chat Messages'];
    sheet.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
    try { CacheService.getScriptCache().remove(CHAT_THREADS_CACHE_KEY); } catch(ce) {}
    if (recipientNames && recipientNames.length) {
      broadcastPushTo(recipientNames, {
        title: 'NEXUS · ' + taskTitle,
        body: message,
        url: APP_LINK,
        tag: 'nexus-chat-task-' + taskId,
      });
    }
  } catch(e) { Logger.log('postSystemTaskMessage error: ' + e.message); }
}

function updateRow(tabName, id, data, token) {
  const sheet = SS.getSheetByName(tabName);
  if (!sheet) return json({ error: 'Tab not found: ' + tabName });

  const allValues  = sheet.getDataRange().getValues();
  const headers    = allValues[0];
  const idStrategy = TAB_ID_COL[tabName] || 'ID';
  let targetRowIndex = -1;

  if (idStrategy === '_rowIndex') {
    const rowNum = parseInt(id, 10);
    if (!isNaN(rowNum) && rowNum >= 1 && rowNum < allValues.length) targetRowIndex = rowNum;
  } else {
    const idColIdx = headers.indexOf(idStrategy);
    if (idColIdx === -1) return json({ error: 'ID column "' + idStrategy + '" not found in ' + tabName });
    for (let i = 1; i < allValues.length; i++) {
      if (String(allValues[i][idColIdx]).trim() === String(id).trim()) { targetRowIndex = i; break; }
    }
  }

  if (targetRowIndex === -1) return json({ error: 'Row not found. id=' + id + ' tab=' + tabName });

  const luIdx = headers.indexOf('Last_Updated');
  if (luIdx >= 0) data['Last_Updated'] = new Date().toISOString();

  const existingRow = allValues[targetRowIndex];

  // RBAC: on tabs with a per-row owner (Assigned To), anyone who isn't a
  // leader/trusted editor may only edit rows assigned to them — not
  // teammates' tasks or partners. COS has no per-row owner (it tracks
  // state records, not individual assignments) so it's exempt here — it's
  // protected purely by field-level locks below instead.
  const OWNERSHIP_TABS = ['1. Master Tasks', '3. Partnerships'];
  if (OWNERSHIP_TABS.includes(tabName)) {
    const callerIdentity = resolveIdentity(token);
    const isLeaderCaller  = !!(callerIdentity && callerIdentity.isLeader);
    const isTrusted = callerIdentity && isTrustedForTab(callerIdentity.name, tabName);
    if (!isLeaderCaller && !isTrusted) {
      const assignedIdx = headers.indexOf('Assigned To');
      const ownerRaw     = assignedIdx >= 0 ? String(existingRow[assignedIdx] || '') : '';
      // Was a single strict-equality check against the whole raw string —
      // failed for multi-assignee tasks ("Ayo T. Joshua, Yewande") and for
      // any name that needed canonicalizing (aliases, missing periods,
      // etc.), even when the caller genuinely was one of the assignees.
      const owners = ownerRaw.split(/[,\/&+]+/)
        .map(n => canonicalizeName(n.trim()).toUpperCase())
        .filter(Boolean);
      const callerUpper = callerIdentity ? canonicalizeName(callerIdentity.name).toUpperCase() : '';
      const label = tabName === '3. Partnerships' ? 'partners' : 'tasks';
      if (!callerIdentity || !owners.length || owners.indexOf(callerUpper) === -1) {
        return json({ error: 'You can only edit ' + label + ' assigned to you.' });
      }
    }
  }

  // RBAC: leader-only fields require a valid leader token OR trusted editor status
  enforceLeaderOnlyFields(data, token, tabName);

  const newRow      = headers.map((h, i) => (data[h] !== undefined) ? data[h] : existingRow[i]);

  // % Progress is a pure, read-only function of Status (and, for Master
  // Tasks with a checklist attached, that checklist's completion within
  // the current status's range — see computeInterpolatedProgress) — ignore
  // any client-sent value outright (defensive; the frontend input is gone)
  // and recompute instead. FROZEN statuses (no ladder entry, e.g.
  // Blocked/Cancelled/Declined) leave % Progress exactly as it already was.
  const progField = PROGRESS_STATUS_FIELD[tabName];
  const progIdx   = headers.indexOf('% Progress');
  if (progIdx >= 0) {
    newRow[progIdx] = existingRow[progIdx]; // ignore any client-sent value; start from current
    if (progField && data[progField] !== undefined) {
      const computed = computeInterpolatedProgress(tabName, data[progField], id);
      if (computed !== null) newRow[progIdx] = computed;
    }
  }

  sheet.getRange(targetRowIndex + 1, 1, 1, newRow.length).setValues([newRow]);

  // NEW (v3.2) — cheap no-op unless this row actually belongs to a goal.
  if (tabName === '1. Master Tasks' && data['Status'] !== undefined) {
    recomputeGoalCompletion_(id);
  }

  // Auto-tracked activity — logged off the real write above, not off
  // anything typed into a text field, so it can't be fabricated. Only
  // fields that genuinely evidence progress are tracked (status, %
  // progress, proof link); free-text fields like Notes are skipped since
  // they're self-reported, not verifiable.
  if (tabName === '1. Master Tasks') {
    const actor = resolveIdentity(token);
    const actorName = (actor && actor.name) || 'Unknown';
    const descIdx = headers.indexOf('Task Description');
    const idIdx   = headers.indexOf('ID');
    const taskDesc = descIdx >= 0 ? newRow[descIdx] : '';
    const taskId   = idIdx >= 0 ? newRow[idIdx] : id;
    ['Status', '% Progress', 'Proof Link'].forEach(field => {
      const fIdx = headers.indexOf(field);
      if (fIdx < 0) return;
      logAutoActivity(actorName, taskId, taskDesc, field, existingRow[fIdx], newRow[fIdx]);
    });
  }

  // In-app acknowledgment ping — fires once, the first time a task moves
  // into the "Acknowledged" status, straight into that task's chat thread.
  // (Previously bound to Not Started -> In Progress, back when there was
  // no dedicated Acknowledged tier — rebound here now that one exists,
  // since that's the actual semantic moment being pinged about.)
  // Deliberately in-app only (chat + push), no email — see
  // postSystemTaskMessage for why. LEADERS get the ping since they're the
  // only ones who can assign tasks in the first place (RBAC), so they're
  // the ones actually waiting to know it landed.
  if (tabName === '1. Master Tasks' && data['Status'] === 'Acknowledged') {
    const statusIdx = headers.indexOf('Status');
    const prevStatus = statusIdx >= 0 ? String(existingRow[statusIdx] || '') : '';
    if (prevStatus !== 'Acknowledged') {
      const actor = resolveIdentity(token);
      const actorName = (actor && actor.name) || 'Someone';
      const descIdx = headers.indexOf('Task Description');
      const idIdx   = headers.indexOf('ID');
      const taskDesc = descIdx >= 0 ? newRow[descIdx] : '';
      const taskId   = idIdx >= 0 ? newRow[idIdx] : id;
      const leaderNames = Object.values(LEADER_DISPLAY_NAMES).filter(n => n !== actorName);
      postSystemTaskMessage(taskId, String(taskDesc).substring(0,60), actorName + ' acknowledged this task.', leaderNames);
    }
  }

  // Fire assignment email if Assigned_To_Email is being set (new assignment via edit)
  const assignedToEmailIdx = headers.indexOf('Assigned_To_Email');
  const prevAssignedEmail  = assignedToEmailIdx >= 0 ? String(existingRow[assignedToEmailIdx] || '') : '';
  const assignmentEmailFired = tabName === '1. Master Tasks' && data['Assigned_To_Email'] && isAutomationEnabled('email_assignments') && data['Assigned_To_Email'] !== prevAssignedEmail;
  if (assignmentEmailFired) {
    // Build full task object from newRow for the email
    const taskObj = {};
    headers.forEach((h, i) => { taskObj[h] = newRow[i]; });
    Logger.log('updateRow — assignment email triggered, email changed from "' + prevAssignedEmail + '" to "' + data['Assigned_To_Email'] + '"');
    sendAssignmentEmail(taskObj);
  }

  // Deadline set/changed independently of a (re)assignment — e.g. a leader
  // adds or moves the date/time on a task that's already assigned. This is
  // exactly the case sendDeadlineEmail/addDeadlineToCalendar's "hooked up to
  // something that changes a deadline independently of assignment" comment
  // was waiting on: previously nothing at all happened here (no calendar
  // event, no email) unless the task was brand new. Skipped when this same
  // edit ALSO just fired the assignment email above — sendAssignmentEmail
  // already creates the calendar event and folds the deadline into its own
  // branded email, so running this too would double-book the event and
  // double-email the assignee(s).
  if (tabName === '1. Master Tasks' && !assignmentEmailFired && data['Deadline'] !== undefined && isAutomationEnabled('email_assignments')) {
    const deadlineIdx  = headers.indexOf('Deadline');
    const prevDeadline = deadlineIdx >= 0 ? String(existingRow[deadlineIdx] || '') : '';
    const newDeadline  = String(data['Deadline'] || '');
    if (newDeadline && newDeadline !== prevDeadline) {
      const taskObj = {};
      headers.forEach((h, i) => { taskObj[h] = newRow[i]; });
      const assigneeEmails = splitEmailList(taskObj['Assigned_To_Email']);
      if (assigneeEmails.length) {
        Logger.log('updateRow — deadline changed independently, from "' + prevDeadline + '" to "' + newDeadline + '"');
        addDeadlineToCalendar(taskObj, assigneeEmails);
        sendDeadlineEmail(taskObj, assigneeEmails);
      } else {
        Logger.log('updateRow — deadline changed but no assignee email on task ' + (taskObj['ID']||id) + '; skipped calendar + email.');
      }
    }
  }

  // Same "Task Assigned" auto-message as new-task creation (see appendRow),
  // but for reassignment — Assigned To changing to someone new on an
  // existing task. Compares against the pre-write value so this only fires
  // on an actual change, not every edit that happens to include this field.
  if (tabName === '1. Master Tasks' && data['Assigned To'] !== undefined) {
    const assignedIdxRA = headers.indexOf('Assigned To');
    const prevAssigned   = assignedIdxRA >= 0 ? String(existingRow[assignedIdxRA] || '').trim() : '';
    const newAssigned    = String(data['Assigned To'] || '').trim();
    if (newAssigned && newAssigned !== prevAssigned) {
      const descIdxRA = headers.indexOf('Task Description');
      const idIdxRA   = headers.indexOf('ID');
      const taskDescRA = descIdxRA >= 0 ? newRow[descIdxRA] : '';
      const taskIdRA   = idIdxRA >= 0 ? newRow[idIdxRA] : id;
      postSystemTaskMessage(
        taskIdRA, String(taskDescRA).substring(0,60),
        'Task assigned to ' + newAssigned + ': ' + taskDescRA,
        [newAssigned],
        'task_assigned'
      );
    }
  }

  return json({ success: true, id });
}

// ═════════════════════════════════════
//  WRITE — BATCH APPEND
// ═════════════════════════════════════

// Tabs sensitive enough that batch-appending into them at all requires a
// leader session (not just stripping a few fields) — team roster and
// automation toggles shouldn't be writable by an anonymous caller who
// just points batchAppend at a different tabName than the UI intends.
const LEADER_ONLY_TABS = ['9. Team Members', '8. Automation Controls'];

function batchAppend(tabName, rows, token) {
  if (LEADER_ONLY_TABS.includes(tabName) && !getLeaderIdentity(token)) {
    return json({ error: 'Leader login required to modify ' + tabName + '.' });
  }
  const sheet    = getOrCreateSheet(tabName, TAB_HEADERS[tabName] || []);
  const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idColIdx = headers.indexOf('ID');

  // Same duplicate-ID guard as appendRow (see the comment there) — batch/
  // bulk-pasted rows can carry explicit IDs too. Check them against both
  // existing rows and each other, and reject the whole batch before writing
  // anything if any collide, rather than a partial write.
  if (idColIdx >= 0) {
    const existingIds = sheet.getLastRow() > 1
      ? new Set(sheet.getRange(2, idColIdx + 1, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()))
      : new Set();
    const seenInBatch = new Set();
    for (const data of rows) {
      const manualId = String(data['ID'] || '').trim();
      if (!manualId) continue;
      if (existingIds.has(manualId) || seenInBatch.has(manualId)) {
        return json({ error: 'ID "' + manualId + '" is already in use — pick a different ID for each row, or edit the existing item instead.' });
      }
      seenInBatch.add(manualId);
    }
  }

  rows.forEach(data => {
    // Was: only called for '1. Master Tasks'. Now always enforced so
    // leader-only fields (Deadline, Priority, etc.) can't be smuggled in
    // via batchAppend against any tab that happens to share those column names.
    enforceLeaderOnlyFields(data, token, tabName);
    if (headers.includes('Last_Updated')) data['Last_Updated'] = new Date().toISOString();
    const row = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');
    if (idColIdx >= 0 && !row[idColIdx]) row[idColIdx] = tabName.split('.')[0].trim() + '-' + Date.now() + Math.random().toString(36).substr(2,4);
    sheet.appendRow(row);
  });
  return json({ success: true, count: rows.length });
}

// ═════════════════════════════════════
//  PROGRESS CALCULATOR (per-tab status ladders)
// ═════════════════════════════════════
// % Progress is now a pure, READ-ONLY function of Status — there is no
// manual override anymore (the frontend no longer exposes an editable
// % Progress input, and this backend ignores any client-sent value — see
// the '% Progress' handling in appendRow/updateRow below). Each tab has
// its own ordered ladder, and each tab has ONE field that drives it (see
// PROGRESS_STATUS_FIELD — 'Status' for Master Tasks/COS, 'Current Status'
// for Partnerships; Document Status and WhatsApp Group Created are
// deliberately NOT folded in, they stay separate tracked fields).
//
// This is no longer a one-way "floor that only nudges up" — since Status
// is the only lever left, setting it directly SETS % Progress to that
// tier's value, including downward moves. That's intentional: e.g. Master
// Tasks' "Revisions Needed" sits below "In Progress" on purpose, so a
// bounced-back task visibly drops instead of freezing wherever it was.
//
// FROZEN statuses are the one exception — any status NOT listed in a
// tab's ladder (Master Tasks: Blocked/Cancelled; Partnerships: Declined)
// doesn't participate at all. % Progress is left exactly as it was the
// moment the item entered that state. See progressToStatus() below for
// how dependency auto-unblocking restores the pre-Blocked value.
const PROGRESS_LADDERS = {
  '1. Master Tasks': {
    'Not Started'      : 0,
    'Acknowledged'     : 10,
    'In Progress'      : 40,
    'Revisions Needed' : 25,   // backward tier — bounced-back work, sits below In Progress
    'In Review'        : 80,
    'Done'             : 100,
    // 'Blocked' and 'Cancelled' intentionally omitted — FROZEN, see comment above.
  },
  '2. COS': {
    'Not Started'      : 0,
    'Acknowledged'     : 10,
    'In Progress'      : 40,
    'Follow-Up Needed' : 20,   // backward tier — admin gone quiet, needs re-engagement
    'Active'           : 100,
  },
  '3. Partnerships': {
    'Not Started'  : 0,
    'In Discussion': 25,
    'Negotiating'  : 55,
    'Cold'         : 15,   // backward tier — went quiet, sits below In Discussion
    'Active'       : 100,
    // 'Declined' intentionally omitted — FROZEN, see comment above.
  },
};

// Which field on each tab drives its progress ladder.
const PROGRESS_STATUS_FIELD = {
  '1. Master Tasks': 'Status',
  '2. COS'         : 'Status',
  '3. Partnerships': 'Current Status',
};

// Returns the % Progress value for a given tab+status, or null if that
// status is FROZEN (not in the ladder) — null means "don't touch % Progress".
function statusToProgress(tabName, status) {
  const ladder = PROGRESS_LADDERS[tabName];
  if (!ladder) return null;
  return Object.prototype.hasOwnProperty.call(ladder, status) ? ladder[status] : null;
}

// Reverse lookup — given a frozen % Progress value, find the ladder status
// it corresponds to (values are unique within a tab's ladder by design).
// Used when un-freezing a status (e.g. dependency auto-unblock) so the
// restored status matches whatever % Progress was actually frozen at,
// instead of blindly jumping back to a fixed status and silently
// resetting real progress. Returns null if no exact match (e.g. legacy
// data with an arbitrary value never in the ladder) — callers should fall
// back to a sensible default status in that case.
function progressToStatus(tabName, pct) {
  const ladder = PROGRESS_LADDERS[tabName];
  if (!ladder) return null;
  for (const key in ladder) {
    if (ladder[key] === pct) return key;
  }
  return null;
}

function progressToStatus(tabName, pct) {
  const ladder = PROGRESS_LADDERS[tabName];
  if (!ladder) return null;
  for (const key in ladder) {
    if (ladder[key] === pct) return key;
  }
  return null;
}

// Ceiling status each Master Tasks status interpolates toward when a task
// has a checklist attached (see computeInterpolatedProgress) — checklist
// completion moves % Progress between the current status's floor and this
// ceiling's value, it never changes which status/range the task is in.
// Checklist-only feature, Master Tasks only. Statuses not listed here
// (Done, and anything FROZEN — Blocked/Cancelled) don't interpolate: Done
// has nothing above it, and frozen tasks don't move regardless of checklist.
const PROGRESS_CEILING = {
  '1. Master Tasks': {
    'Not Started'      : 'Acknowledged',
    'Acknowledged'     : 'In Progress',
    'In Progress'      : 'In Review',
    'Revisions Needed' : 'In Progress',
    'In Review'        : 'Done',
  },
};

// % Progress for a status, folding in checklist completion when the task
// has one. Falls back to the plain status floor (statusToProgress) with no
// interpolation when: the status is FROZEN (returns null, same as before —
// caller leaves % Progress untouched), there's no ceiling defined for this
// status/tab, or the task has no checklist items at all.
function computeInterpolatedProgress(tabName, status, taskId) {
  const floor = statusToProgress(tabName, status);
  if (floor === null) return null;
  const ceilingStatus = (PROGRESS_CEILING[tabName] || {})[status];
  if (!ceilingStatus || !taskId) return floor;
  const ceiling = statusToProgress(tabName, ceilingStatus);
  if (ceiling === null) return floor;

  const items = sheetToObjects('17. Task Checklists').filter(c => String(c['Task ID']) === String(taskId));
  if (!items.length) return floor;
  const doneCount = items.filter(c => String(c.Done).trim().toUpperCase() === 'Y').length;
  return Math.round(floor + (doneCount / items.length) * (ceiling - floor));
}

// Re-derives and writes % Progress for one Master Tasks row from its
// current Status + current checklist completion — used whenever a
// checklist item is added/toggled/deleted (no Status change involved, so
// updateRow's own progress logic never runs for that).
function recomputeTaskProgress(taskId) {
  const sheet = SS.getSheetByName('1. Master Tasks');
  if (!sheet) return;
  const headers   = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idIdx     = headers.indexOf('ID');
  const statusIdx = headers.indexOf('Status');
  const progIdx   = headers.indexOf('% Progress');
  if (idIdx < 0 || statusIdx < 0 || progIdx < 0) return;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(taskId)) {
      const computed = computeInterpolatedProgress('1. Master Tasks', values[i][statusIdx], taskId);
      if (computed !== null) sheet.getRange(i+1, progIdx+1).setValue(computed);
      return;
    }
  }
}

// Same access rule as chat-thread access minus the "ever posted" fallback
// (checklist editing is scoped tighter, per explicit scoping conversation):
// leader, trusted-for-Master-Tasks, current assignee, or current support
// person on that specific task.
function canEditTaskChecklist(identity, taskId) {
  if (identity.isLeader) return true;
  if (isTrustedForTab(identity.name, '1. Master Tasks')) return true;
  const me = canonicalizeName(identity.name).trim().toUpperCase();
  const task = sheetToObjects('1. Master Tasks').find(t => String(t['ID']) === String(taskId || ''));
  if (!task) return false;
  // Was comparing the raw "Assigned To" / "Support Person" strings whole
  // against `me` — never matched on a multi-assignee/multi-support task.
  // Same shared, split-then-canonicalize-each check as everywhere else,
  // which also brings Supervisor into checklist access.
  return isPersonOnTask(me, task);
}

function getTaskChecklist(taskId, token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!canEditTaskChecklist(identity, taskId)) return json({ error: 'No access to this task.' });
  return json({ success: true, items: sheetToObjects('17. Task Checklists').filter(c => String(c['Task ID']) === String(taskId)) });
}

function addChecklistItem(taskId, description, token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!canEditTaskChecklist(identity, taskId)) return json({ error: 'No access to this task.' });
  const desc = String(description || '').trim();
  if (!desc) return json({ error: 'Description required.' });
  const sheet   = getOrCreateSheet('17. Task Checklists', TAB_HEADERS['17. Task Checklists']);
  const headers = TAB_HEADERS['17. Task Checklists'];
  const row = {
    ID: 'CHK-' + Date.now() + '-' + Math.random().toString(36).substr(2,4),
    'Task ID': taskId, Description: desc, Done: 'N',
    Created_At: new Date().toISOString(), Created_By: identity.name,
  };
  sheet.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
  recomputeTaskProgress(taskId);
  return json({ success: true, id: row.ID });
}

function toggleChecklistItem(itemId, taskId, done, token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!canEditTaskChecklist(identity, taskId)) return json({ error: 'No access to this task.' });
  const sheet = SS.getSheetByName('17. Task Checklists');
  if (!sheet) return json({ error: 'Checklist not found.' });
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idIdx = headers.indexOf('ID'), doneIdx = headers.indexOf('Done');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(itemId)) {
      sheet.getRange(i+1, doneIdx+1).setValue(done ? 'Y' : 'N');
      recomputeTaskProgress(taskId);
      return json({ success: true });
    }
  }
  return json({ error: 'Checklist item not found.' });
}

function deleteChecklistItem(itemId, taskId, token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  if (!canEditTaskChecklist(identity, taskId)) return json({ error: 'No access to this task.' });
  const sheet = SS.getSheetByName('17. Task Checklists');
  if (!sheet) return json({ error: 'Checklist not found.' });
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idIdx = headers.indexOf('ID');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(itemId)) {
      sheet.deleteRow(i+1);
      recomputeTaskProgress(taskId);
      return json({ success: true });
    }
  }
  return json({ error: 'Checklist item not found.' });
}

function updateTaskStatus(taskId, newStatus, token) {
  const identity = resolveIdentity(token);
  if (!identity) return json({ error: 'Login required.' });
  const sheet   = SS.getSheetByName('1. Master Tasks');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf('ID');
  const statCol = headers.indexOf('Status');
  const progCol = headers.indexOf('% Progress');
  const updCol  = headers.indexOf('Last_Updated');
  const descCol = headers.indexOf('Task Description');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(taskId)) {
      const prevStatus = data[i][statCol];
      const prevProg   = progCol >= 0 ? data[i][progCol] : undefined;
      sheet.getRange(i+1, statCol+1).setValue(newStatus);
      recomputeGoalCompletion_(taskId); // NEW (v3.2) — cheap no-op unless this task belongs to a goal
      // % Progress is a pure function of Status (plus checklist completion
      // within the current status's range, if this task has one — see
      // computeInterpolatedProgress). Can move down, e.g. into "Revisions
      // Needed". FROZEN statuses (Blocked/Cancelled) return null and are
      // left exactly as they were — see PROGRESS_LADDERS comment.
      let newProg = prevProg;
      if (progCol >= 0) {
        const computed = computeInterpolatedProgress('1. Master Tasks', newStatus, taskId);
        if (computed !== null) {
          newProg = computed;
          sheet.getRange(i+1, progCol+1).setValue(computed);
        }
      }
      if (updCol >= 0) sheet.getRange(i+1, updCol+1).setValue(new Date().toISOString());
      if (isAutomationEnabled('dependency_checks')) updateDependencyStatuses();
      // Auto-tracked activity — see comment on the same block in updateRow().
      const taskDesc = descCol >= 0 ? data[i][descCol] : '';
      logAutoActivity(identity.name, taskId, taskDesc, 'Status', prevStatus, newStatus);
      if (progCol >= 0) logAutoActivity(identity.name, taskId, taskDesc, '% Progress', prevProg, newProg);
      Logger.log('Task status updated by ' + identity.name + ' — ' + taskId + ' → ' + newStatus);
      return json({ success: true, taskId, newStatus });
    }
  }
  return json({ error: 'Task not found: ' + taskId });
}

// ═════════════════════════════════════
//  ✅ FIX: FOLLOW-UP EMAIL AUTOMATION
//  This is what was missing in v3.0.
//  Sends DIRECT Gmail at day 2 (nudge), day 5 (warning), day 7 (escalation).
//  Does NOT depend on n8n being configured.
//  Run via installTriggers() → fires daily at 6 PM.
// ═════════════════════════════════════

// Priority-scaled escalation ladder (Master Tasks only) — these exact
// day-numbers are proposed defaults from a scoping conversation, not
// locked-final; MEDIUM matches the pre-existing 2/5/7 default unchanged.
// Two independent triggers feed the same 3-tier system — staleness
// (days since Last_Updated) and deadline proximity (days until Deadline,
// gated on progress still being low) — whichever fires first wins: an
// imminent, still-low-progress deadline escalates straight to tier 3
// regardless of what the staleness count is.
const ESCALATION_TIERS = {
  'URGENT': { staleness: [1,2,3],   deadlineDays: 2, deadlineProgressMax: 100 },
  'HIGH'  : { staleness: [2,4,6],   deadlineDays: 3, deadlineProgressMax: 100 },
  'MEDIUM': { staleness: [2,5,7],   deadlineDays: 2, deadlineProgressMax: 50  },
  'LOW'   : { staleness: [4,8,12],  deadlineDays: 1, deadlineProgressMax: 50  },
};

function runFollowUpAutomation() {
  if (!isAutomationEnabled('overdue_reminders')) return;

  const tasks    = sheetToObjects('1. Master Tasks');
  const today    = new Date();
  const skipStat = ['Done','Cancelled'];

  tasks.forEach(task => {
    if (skipStat.includes(task['Status']))   return;
    if (!task['Assigned_To_Email'])           return;
    if (!task['Last_Updated'])                return;

    const cfg = ESCALATION_TIERS[task['Priority']] || ESCALATION_TIERS['MEDIUM'];

    const staleDays = Math.floor((today - new Date(task['Last_Updated'])) / 86400000);
    let tier = 0;
    if      (staleDays === cfg.staleness[2]) tier = 3;
    else if (staleDays === cfg.staleness[1]) tier = 2;
    else if (staleDays === cfg.staleness[0]) tier = 1;

    let deadlineTriggered = false;
    if (tier < 3 && task['Deadline']) {
      const progress = parseInt(task['% Progress']) || 0;
      if (progress < cfg.deadlineProgressMax) {
        const deadline = new Date(task['Deadline']); deadline.setHours(0,0,0,0);
        const todayMid = new Date(); todayMid.setHours(0,0,0,0);
        const daysOut  = Math.round((deadline - todayMid) / 86400000);
        if (daysOut === cfg.deadlineDays) { tier = 3; deadlineTriggered = true; }
      }
    }
    if (tier === 0) return;

    const assigneeName  = task['Assigned To']       || 'Team member';
    const assigneeEmail = task['Assigned_To_Email'];
    const taskId        = task['ID']                || '—';
    const taskDesc       = task['Task Description']  || '(no description)';
    const deadline        = formatDeadline(task['Deadline'], task['Deadline_Time']);
    const status           = task['Status']            || '—';
    const progress          = task['% Progress']        || '0';
    const supportPersonName = task['Support Person'] || '';
    const supervisorName    = task['Supervisor']     || '';
    const reason = deadlineTriggered
      ? `its deadline is ${cfg.deadlineDays} day(s) away and progress is still under ${cfg.deadlineProgressMax}%`
      : `it hasn't been updated in ${staleDays} day(s)`;

    const taskBlock =
      `━━━━━━━━━━━━━━━━━━━━\n` +
      `Task ID:  ${taskId}\n` +
      `Task:     ${taskDesc}\n` +
      `Priority: ${task['Priority'] || 'MEDIUM'}\n` +
      `Status:   ${status}\n` +
      `Progress: ${progress}%\n` +
      `Deadline: ${deadline}\n` +
      `━━━━━━━━━━━━━━━━━━━━\n\n`;

    let pushNames = [];

    // Email is now gated per-recipient by Notify_Pref (wantsEmail) — this
    // is the fix for escalation traffic flooding a leader's inbox once the
    // team crosses ~100 people: anyone (the assignee, team lead, COO,
    // deputy, or founder) can set their preference to "In-App" and they'll
    // still see every tier of escalation from the bell icon, just not in
    // email. Nobody is silently dropped — wantsInApp() below still fires
    // for everyone regardless of their email choice.
    if (tier === 1) {
      const subject = `Reminder: "${taskDesc}" — ${reason}`;
      const body =
        `Hi ${assigneeName},\n\n` +
        `Just a nudge — ${reason}.\n\n` +
        taskBlock +
        `Please update the status or progress:\n${APP_LINK}\n\n` +
        `If you're blocked, mark the task as "Blocked" so leadership can help.\n\n` +
        `— NEXUS Automation · ABION Industries`;
      const cc2day = [TEAM_LEAD_NAME].filter(n => wantsEmail(n)).map(() => TEAM_LEAD_EMAIL).filter(e => e && !e.includes('your_')).join(',');
      try { emailIfWanted_(assigneeName, assigneeEmail, subject, null, body, { replyTo: FOUNDER_EMAIL, cc: cc2day || null }); } catch(e) { Logger.log('tier1 email error: '+e.message); }
      pushNames = [assigneeName, TEAM_LEAD_NAME];
    }

    if (tier === 2) {
      const subject = `⚠️ Action Needed: "${taskDesc}" — ${reason}`;
      const body =
        `Hi ${assigneeName},\n\n` +
        `This task needs attention: ${reason}, and may be at risk.\n\n` +
        taskBlock +
        `Please update now:\n${APP_LINK}\n\n` +
        `— NEXUS Automation · ABION Industries`;
      const tier2CcNames = [[TEAM_LEAD_NAME, TEAM_LEAD_EMAIL], [COO_NAME, COO_EMAIL], [DEPUTY_NAME, DEPUTY_EMAIL]];
      const tier2Cc = tier2CcNames.filter(([n, e]) => e && !e.includes('your_') && wantsEmail(n)).map(([, e]) => e);
      try {
        emailIfWanted_(assigneeName, assigneeEmail, subject, null, body, {
          replyTo: FOUNDER_EMAIL,
          cc: tier2Cc,
        });
      } catch(e) { Logger.log(e.message); }
      pushNames = [assigneeName, TEAM_LEAD_NAME, COO_NAME, DEPUTY_NAME];
    }

    if (tier === 3) {
      const subject = `🚨 Escalation: "${taskDesc}" — ${reason}`;
      const body =
        `This task requires your immediate review: ${reason}.\n\n` +
        taskBlock +
        `Assigned To: ${assigneeName} (${assigneeEmail})\n` +
        (supportPersonName ? `Support Person: ${supportPersonName}\n` : '') +
        (supervisorName ? `Supervisor:  ${supervisorName}\n` : '') +
        `\nReview and take action:\n${APP_LINK}\n\n` +
        `— NEXUS Automation · ABION Industries`;
      const tier3CcNames = [[canonicalizeName(assigneeName), assigneeEmail], [COO_NAME, COO_EMAIL], [DEPUTY_NAME, DEPUTY_EMAIL]];
      const tier3Cc = tier3CcNames.filter(([n, e]) => e && !e.includes('your_') && wantsEmail(n)).map(([, e]) => e);
      try {
        emailIfWanted_(FOUNDER_NAME, FOUNDER_EMAIL, subject, null, body, {
          cc: tier3Cc,
        });
      } catch(e) { Logger.log(e.message); }
      pushNames = [FOUNDER_NAME, assigneeName, COO_NAME, DEPUTY_NAME];
    }

    // In-app layer — same widening audience as the email, but every name
    // here gets a persistent row in the bell icon (subject to their own
    // Notify_Pref) regardless of whether they also got an email, and the
    // link deep-links straight to this task instead of the homepage.
    const uniqueNames = Array.from(new Set(pushNames.map(n => canonicalizeName(String(n)).trim())));
    notifyInApp(uniqueNames, {
      type: 'escalation',
      title: 'NEXUS · Task Escalation · ' + taskDesc.substring(0,50),
      body: reason.charAt(0).toUpperCase() + reason.slice(1) + ' (tier ' + tier + ')',
      url: notifUrl_({ task: taskId }),
      tag: 'nexus-escalation-' + taskId,
      taskId: taskId,
    });
  });

  runPartnerFollowUps();
}

function runPartnerFollowUps() {
  const partners = sheetToObjects('3. Partnerships');
  const today    = new Date();
  const skip     = ['Active','Declined'];

  partners.forEach(p => {
    if (skip.includes(p['Current Status'])) return;
    if (!p['Follow-Up Date'])               return;

    const daysLate = Math.floor((today - new Date(p['Follow-Up Date'])) / 86400000);
    if (daysLate <= 0) return;

    if (daysLate === 1 || daysLate === 3 || daysLate === 7) {
      const isCritical = daysLate >= 7;
      const subject    = isCritical
        ? `🚨 Partner Escalation: ${p['Organisation']} — ${daysLate} days overdue`
        : `⏱ Partner Follow-Up Overdue: ${p['Organisation']} (${daysLate}d late)`;
      const body =
        `Partner follow-up is overdue.\n\n` +
        `━━━━━━━━━━━━━━━━━━━━\n` +
        `Organisation: ${p['Organisation']}\n` +
        `Assigned To:  ${p['Assigned To']}\n` +
        `Status:       ${p['Current Status']}\n` +
        `Was Due:      ${new Date(p['Follow-Up Date']).toDateString()}\n` +
        `Days Overdue: ${daysLate}\n` +
        `━━━━━━━━━━━━━━━━━━━━\n\n` +
        `Update in NEXUS:\n${SHEET_LINK}\n\n` +
        `— NEXUS Automation · ABION Industries`;
      try { emailIfWanted_(FOUNDER_NAME, FOUNDER_EMAIL, subject, null, body); } catch(e) { Logger.log(e.message); }
      // In-app layer — reuses the same escalation infra as Master Tasks,
      // ladder/audience otherwise untouched per spec (Founder only), but
      // now goes through notifyInApp so it's visible from the bell icon
      // even for a founder who's opted out of the email for this.
      notifyInApp([canonicalizeName(FOUNDER_NAME)], {
        type: 'partner_escalation',
        title: 'NEXUS · Partner Follow-Up · ' + p['Organisation'],
        body: (isCritical ? '🚨 ' : '⏱ ') + daysLate + ' day(s) overdue',
        url: notifUrl_({ notif: 1 }),
        tag: 'nexus-partner-escalation-' + p['Organisation'],
      });
    }
  });
}

// ═════════════════════════════════════
//  ✅ FIX: EOD KPI CHECK WITH DIRECT GMAIL
//  Was missing in v3.0 — only pinged n8n.
//  Now sends email directly, n8n ping is optional bonus.
// ═════════════════════════════════════

// ─── EOD ATTENDANCE SCHEDULE ──────────────────────────────────
// Femi = Mon–Fri (all days). Ayo = Mon, Tue, Thu, Fri (not Wed).
// Add other names with their days below (0=Sun,1=Mon,...,5=Fri,6=Sat).
const EOD_SCHEDULE = {
  'Femi Balogun'  : [1,2,3,4,5],  // Mon-Fri
  'Ayo T. Joshua' : [1,2,4,5],    // Mon,Tue,Thu,Fri — no Wednesday
};
// Members NOT in EOD_SCHEDULE are expected every Mon-Fri by default.

function getMembersExpectedToday() {
  const today = new Date();
  const day   = today.getDay(); // 0=Sun...6=Sat
  if (day === 0 || day === 6) return []; // weekend
  return getAllTeamNames().filter(m => {
    const schedule = EOD_SCHEDULE[m];
    if (schedule) return schedule.includes(day);
    return true; // default: Mon-Fri
  });
}

function checkMissingKPI() {
  if (!isAutomationEnabled('eod_nudges')) return;

  const today    = new Date();
  const day      = today.getDay();
  if (day === 0 || day === 6) return; // Skip weekends

  const hour = today.getHours();
  if (hour < 18) return; // Only run after 6 PM

  const expected = getMembersExpectedToday();
  if (!expected.length) return;

  const todayStr  = localDateISO();  // ← FIX: compare as 'YYYY-MM-DD', same format stored
  const kpi       = sheetToObjects('5. KPI Log');
  const filed     = new Set(
    kpi.filter(k => k['Date'] === todayStr)
       .map(k => k['Team Member'])
  );
  const missing = expected.filter(m => !filed.has(m));
  if (missing.length === 0) return;

  const subject = `📋 EOD Alert: ${missing.length} member(s) expected today have not filed (${today.toDateString()})`;
  const body    =
    `The following team members were expected today but have not submitted their EOD report:\n\n` +
    missing.map(m => `  → ${m}`).join('\n') +
    `\n\nOnly members scheduled for today are listed.\n` +
    `Review the KPI log:\n${SHEET_LINK}\n\n` +
    `— NEXUS Automation · ABION Industries`;

  try {
    sendViaResend(FOUNDER_EMAIL, subject, null, body);
    if (COO_EMAIL && COO_EMAIL !== 'your_coo@email.com') {
      sendViaResend(COO_EMAIL, subject, null, body);
    }
  } catch(e) { Logger.log(e.message); }

  if (isAutomationEnabled('n8n_webhooks')) {
    pingN8N(N8N.EOD_NUDGE, { missing, expected: expected.length, date: todayStr });
  }
}

// ═════════════════════════════════════
//  DEADLINE TRACKING & EMAIL REMINDERS
// ═════════════════════════════════════

function checkDeadlinesAndNotify(force) {
  if (!force && !isAutomationEnabled('deadline_reminders')) return;
  const today = new Date(); today.setHours(0,0,0,0);
  const tasks = sheetToObjects('1. Master Tasks');
  const sent  = { overdue:0, upcoming:0 };

  tasks.forEach(task => {
    if (['Done','Cancelled'].includes(task['Status'])) return;
    if (!task['Deadline'] || !task['Assigned_To_Email'])  return;

    const deadline = new Date(task['Deadline']); deadline.setHours(0,0,0,0);
    const diffDays = Math.round((deadline - today) / 86400000);

    let subject = '', body = '';

    if (diffDays < 0) {
      subject = `⚠️ OVERDUE: ${task['Task Description']}`;
      body    = `Hi ${task['Assigned To']},\n\nYour task is ${Math.abs(diffDays)} day(s) overdue.\n\n`
              + `Task: ${task['Task Description']}\nDeadline was: ${task['Deadline']}\nStatus: ${task['Status']}\n\n`
              + `Please update the task status or contact your supervisor.\n\nOpen in NEXUS: ${APP_LINK}\n\n— NEXUS Automation`;
      sent.overdue++;
    } else if (diffDays === 3 || diffDays === 1) {
      subject = `🔔 Deadline in ${diffDays} day(s): ${task['Task Description']}`;
      body    = `Hi ${task['Assigned To']},\n\nYour task is due in ${diffDays} day(s).\n\n`
              + `Task: ${task['Task Description']}\nDeadline: ${task['Deadline']}\nStatus: ${task['Status']}\n`
              + `Progress: ${task['% Progress']}%\n\nOpen in NEXUS: ${APP_LINK}\n\n— NEXUS Automation`;
      sent.upcoming++;
    }

    if (body) {
      try { emailIfWanted_(task['Assigned To'], task['Assigned_To_Email'], subject, null, body); } catch(e) { Logger.log(e.message); }
      notifyInApp([task['Assigned To']], {
        type: diffDays < 0 ? 'overdue' : 'deadline_reminder',
        title: 'NEXUS · ' + (diffDays < 0 ? 'Task Overdue' : 'Deadline in ' + diffDays + ' day(s)'),
        body: String(task['Task Description'] || '').substring(0, 90),
        url: notifUrl_({ task: task['ID'] }),
        tag: 'nexus-deadline-check-' + task['ID'],
        taskId: task['ID'],
      });
    }
  });

  if (isAutomationEnabled('n8n_webhooks')) triggerOverdueCheck();
  return json({ success:true, emailsSent:sent });
}

// ═════════════════════════════════════
//  DEPENDENCY MANAGEMENT
// ═════════════════════════════════════

function updateDependencyStatuses() {
  if (!isAutomationEnabled('dependency_checks')) return json({ success:true, skipped:true });

  const sheet     = SS.getSheetByName('1. Master Tasks');
  const allValues = sheet.getDataRange().getValues();
  const headers   = allValues[0];
  const idIdx     = headers.indexOf('ID');
  const statusIdx = headers.indexOf('Status');
  const depIdx    = headers.indexOf('Depends On');
  const updIdx    = headers.indexOf('Last_Updated');

  if (depIdx === -1) return json({ success:true, note:'"Depends On" column not found.' });

  const statusMap = {};
  for (let i = 1; i < allValues.length; i++) {
    if (allValues[i][idIdx]) statusMap[String(allValues[i][idIdx]).trim()] = String(allValues[i][statusIdx]).trim();
  }

  let changed = 0;
  for (let i = 1; i < allValues.length; i++) {
    const deps          = String(allValues[i][depIdx] || '').trim();
    if (!deps)          continue;
    const currentStatus = String(allValues[i][statusIdx]).trim();
    if (['Done','Cancelled'].includes(currentStatus)) continue;

    const depIds        = deps.split(',').map(d => d.trim()).filter(Boolean);
    const allDone       = depIds.every(id => statusMap[id] === 'Done');
    const anyIncomplete = depIds.some(id => statusMap[id] && statusMap[id] !== 'Done');

    if (anyIncomplete && currentStatus !== 'Blocked') {
      sheet.getRange(i+1, statusIdx+1).setValue('Blocked');
      if (updIdx >= 0) sheet.getRange(i+1, updIdx+1).setValue(new Date().toISOString());
      changed++;
    } else if (allDone && currentStatus === 'Blocked') {
      // % Progress was FROZEN at whatever it was the moment this task got
      // Blocked (see PROGRESS_LADDERS comment) — so on unblock, restore
      // the ladder status that actually matches that frozen value instead
      // of blindly resetting to 'In Progress', which would silently wipe
      // out real progress (e.g. a task blocked while at 'In Review'/80%
      // must come back to 'In Review', not drop to 'In Progress'/40%).
      const progIdxDep = headers.indexOf('% Progress');
      const frozenPct  = progIdxDep >= 0 ? (parseInt(allValues[i][progIdxDep]) || 0) : null;
      const restored   = (frozenPct !== null && progressToStatus('1. Master Tasks', frozenPct)) || 'In Progress';
      sheet.getRange(i+1, statusIdx+1).setValue(restored);
      if (updIdx >= 0) sheet.getRange(i+1, updIdx+1).setValue(new Date().toISOString());
      changed++;
    }
  }
  return json({ success:true, tasksUpdated:changed });
}

// ═════════════════════════════════════
//  RECURRING TASK TEMPLATES
// ═════════════════════════════════════

function createRecurringTasks() {
  if (!isAutomationEnabled('recurring_tasks')) return json({ success:true, skipped:true });

  const sheet     = SS.getSheetByName('1. Master Tasks');
  const allValues = sheet.getDataRange().getValues();
  const headers   = allValues[0];
  const ci        = {}; // column index map
  ['ID','Task Description','Priority','Assigned To','Support Person','Supervisor','Supervisor_Email','Business / Dept',
   'Deadline','Status','% Progress','Assigned_To_Email','Last_Updated','Notes',
   'Manual Trigger Required','Depends On','Recurring','Recur Interval','Last Run'].forEach(h => {
    ci[h] = headers.indexOf(h);
  });

  if (ci['Recurring'] === -1 || ci['Recur Interval'] === -1) {
    return json({ success:false, error:'"Recurring" or "Recur Interval" columns not found.' });
  }

  const today = new Date(); today.setHours(0,0,0,0);
  let generated = 0;

  for (let i = 1; i < allValues.length; i++) {
    const row = allValues[i];
    if (String(row[ci['Recurring']]).trim().toLowerCase() !== 'yes') continue;

    const interval = parseInt(row[ci['Recur Interval']], 10);
    if (isNaN(interval) || interval <= 0) continue;

    const lastRun = row[ci['Last Run']] ? new Date(row[ci['Last Run']]) : null;
    const nextRun = lastRun ? new Date(lastRun.getTime() + interval * 86400000) : today;
    nextRun.setHours(0,0,0,0);
    if (today < nextRun) continue;

    const newDeadline = new Date(today.getTime() + interval * 86400000);
    const newRow      = headers.map(() => '');
    const now         = new Date().toISOString();

    newRow[ci['ID']]                     = '1-' + Date.now();
    newRow[ci['Task Description']]       = row[ci['Task Description']] + ' (Auto)';
    newRow[ci['Priority']]               = row[ci['Priority']];
    newRow[ci['Assigned To']]            = row[ci['Assigned To']];
    if (ci['Support Person'] >= 0)       newRow[ci['Support Person']]       = row[ci['Support Person']];
    if (ci['Supervisor'] >= 0)           newRow[ci['Supervisor']]           = row[ci['Supervisor']];
    if (ci['Supervisor_Email'] >= 0)     newRow[ci['Supervisor_Email']]     = row[ci['Supervisor_Email']];
    if (ci['Business / Dept'] >= 0)      newRow[ci['Business / Dept']]      = row[ci['Business / Dept']];
    newRow[ci['Deadline']]               = newDeadline.toISOString().split('T')[0];
    newRow[ci['Status']]                 = 'Not Started';
    newRow[ci['% Progress']]             = 0;
    if (ci['Assigned_To_Email'] >= 0)    newRow[ci['Assigned_To_Email']]    = row[ci['Assigned_To_Email']];
    if (ci['Last_Updated'] >= 0)         newRow[ci['Last_Updated']]         = now;
    if (ci['Notes'] >= 0)               newRow[ci['Notes']]                = 'Auto-generated recurring task';
    if (ci['Manual Trigger Required']>=0)newRow[ci['Manual Trigger Required']]='No';
    newRow[ci['Recurring']]              = 'No'; // instances are not templates

    sheet.appendRow(newRow);
    if (ci['Last Run'] >= 0) sheet.getRange(i+1, ci['Last Run']+1).setValue(today.toISOString().split('T')[0]);
    generated++;
  }
  return json({ success:true, tasksGenerated:generated });
}

function addRecurringTask(data, token) {
  // Was: trusted a client-supplied `user` string and did a loose substring
  // match against manager names — trivially spoofable (anyone typing "ayo"
  // passed). Now requires an actual leader session.
  const callerName = getLeaderIdentity(token);
  const caller = (callerName || '').toLowerCase();
  if (!callerName || !RBAC.editDeadline.some(m => caller.includes(m.toLowerCase()))) {
    return json({ error:'Permission denied: only managers can create recurring tasks.' });
  }
  data['Recurring']      = 'Yes';
  data['Recur Interval'] = data['Recur Interval'] || 7;
  data['Last Run']       = '';
  return appendRow('1. Master Tasks', data);
}

// ═════════════════════════════════════
//  GOOGLE CALENDAR SYNC
// ═════════════════════════════════════

function createCalendarEvents(callerEmail) {
  if (!isAutomationEnabled('calendar_sync')) return json({ success:true, skipped:true });
  const tasks    = sheetToObjects('1. Master Tasks');
  const calendar = CalendarApp.getDefaultCalendar();
  let created = 0, errors = 0;

  tasks.forEach(task => {
    if (['Done','Cancelled'].includes(task['Status'])) return;
    if (!task['Deadline']) return;
    try {
      const deadline = new Date(task['Deadline']); deadline.setHours(9,0,0,0);
      const title    = `[ABION] ${task['Task Description']}`;
      const desc     = `Assigned to: ${task['Assigned To']}\nPriority: ${task['Priority']}\nStatus: ${task['Status']}\nProgress: ${task['% Progress']}%${task['Notes']?'\nNotes: '+task['Notes']:''}`;
      const dayStart = new Date(deadline); dayStart.setHours(0,0,0,0);
      const dayEnd   = new Date(deadline); dayEnd.setHours(23,59,59,999);
      const existing = calendar.getEvents(dayStart, dayEnd).filter(ev => ev.getTitle() === title);
      if (!existing.length) {
        const event = calendar.createEvent(title, deadline, new Date(deadline.getTime()+3600000), { description:desc });
        if (task['Assigned_To_Email']) { try { event.addGuest(task['Assigned_To_Email']); } catch(ge) {} }
        created++;
      }
    } catch(e) { errors++; Logger.log('Calendar error for ' + task['ID'] + ': ' + e.message); }
  });
  return json({ success:true, eventsCreated:created, errors });
}

function _createSingleCalendarEvent(taskId, description, deadlineStr, email) {
  try {
    const deadline = new Date(deadlineStr); deadline.setHours(9,0,0,0);
    const event    = CalendarApp.getDefaultCalendar()
      .createEvent('[ABION] ' + description, deadline, new Date(deadline.getTime()+3600000));
    if (email) { try { event.addGuest(email); } catch(e) {} }
  } catch(e) { Logger.log('Calendar single event error: ' + e.message); }
}

// ═════════════════════════════════════
//  AUTOMATION TOGGLE
// ═════════════════════════════════════

function setAutomation(key, enabled, token) {
  const callerName = getLeaderIdentity(token);
  if (!callerName) {
    Logger.log('setAutomation BLOCKED — token: ' + (token ? token.substring(0,8)+'...' : 'NONE'));
    return json({ error: 'Not authorised. Leader token expired — please log out and log in again.' });
  }

  const sheet = SS.getSheetByName('8. Automation Controls');
  if (!sheet) return json({ error:'Automation Controls tab missing' });

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('Key')] === key) {
      const label = data[i][headers.indexOf('Label')] || key;
      sheet.getRange(i+1, headers.indexOf('Enabled')+1)    .setValue(enabled ? 'TRUE' : 'FALSE');
      sheet.getRange(i+1, headers.indexOf('LastToggled')+1).setValue(new Date().toISOString());
      sheet.getRange(i+1, headers.indexOf('ToggledBy')+1)  .setValue(callerName);
      // Full audit trail — every toggle, not just the most recent one.
      const logSheet = getOrCreateSheet('13. Automation Log', TAB_HEADERS['13. Automation Log']);
      logSheet.appendRow([new Date().toISOString(), key, label, enabled ? 'TRUE' : 'FALSE', callerName]);
      return json({ success:true, key, enabled });
    }
  }
  return json({ error:'Automation key not found: ' + key });
}

function isAutomationEnabled(key) {
  try {
    const rows = sheetToObjects('8. Automation Controls');
    const row  = rows.find(r => r['Key'] === key);
    return row ? (row['Enabled'] === 'TRUE' || row['Enabled'] === true) : true;
  } catch(e) { return true; }
}

// ═════════════════════════════════════
//  onEdit TRIGGER
// ═════════════════════════════════════

function onEdit(e) {
  const sheet     = e.range.getSheet();
  const sheetName = sheet.getName();
  const TRACKED   = ['1. Master Tasks','2. COS','3. Partnerships','4. Org Database','5. KPI Log','6. Marketing'];
  if (!TRACKED.includes(sheetName) || e.range.getRow() === 1) return;

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const updIdx  = headers.indexOf('Last_Updated');
  if (updIdx >= 0) sheet.getRange(e.range.getRow(), updIdx+1).setValue(new Date().toISOString());

  if (sheetName === '1. Master Tasks') {
    const emailIdx    = headers.indexOf('Assigned_To_Email');
    const deadlineIdx = headers.indexOf('Deadline');

    if (e.range.getColumn() === emailIdx+1 && e.value && isAutomationEnabled('email_assignments')) {
      const row     = sheet.getRange(e.range.getRow(),1,1,headers.length).getValues()[0];
      const payload = {};
      headers.forEach((h,i) => payload[h] = row[i]);
      sendAssignmentEmail(payload); // direct Gmail — always fires if automation is on
      if (isAutomationEnabled('n8n_webhooks')) pingN8N(N8N.ASSIGNMENT, payload);
    }

    if (e.range.getColumn() === deadlineIdx+1 && e.value && isAutomationEnabled('calendar_sync')) {
      const row  = sheet.getRange(e.range.getRow(),1,1,headers.length).getValues()[0];
      const desc = row[headers.indexOf('Task Description')] || 'Task';
      _createSingleCalendarEvent(row[headers.indexOf('ID')], desc, e.value, row[headers.indexOf('Assigned_To_Email')]||'');
    }
  }
}

// ═════════════════════════════════════
//  DAILY TRIGGERS (called by time-based triggers)
// ═════════════════════════════════════

function dailyDeadlineCheck()    { checkDeadlinesAndNotify(false); }

// ─── TEST: run this manually from Apps Script editor to check email sending ──
function diagEmailSetup() {
  // Run this to check your email config before going live
  const issues = [];
  if (FOUNDER_EMAIL.includes('your@')) issues.push('FOUNDER_EMAIL not set');
  if (COO_EMAIL.includes('your_'))     issues.push('COO_EMAIL not set (CC on 5d/7d reminders will be skipped)');
  if (DEPUTY_EMAIL.includes('your_'))  issues.push('DEPUTY_EMAIL not set (CC on 5d/7d reminders will be skipped)');
  if (TEAM_LEAD_EMAIL.includes('your_'))issues.push('TEAM_LEAD_EMAIL not set (CC on 2d reminders will be skipped)');
  const enabled = isAutomationEnabled('email_assignments');
  Logger.log('=== NEXUS Email Diagnostics ===');
  Logger.log('email_assignments automation: ' + (enabled?'ENABLED':'DISABLED ← FIX THIS'));
  Logger.log('FOUNDER_EMAIL: ' + FOUNDER_EMAIL);
  if (issues.length) { Logger.log('Warnings (non-fatal):'); issues.forEach(i=>Logger.log('  ⚠ '+i)); }
  else Logger.log('All emails configured.');
  Logger.log('Run testAssignmentEmail() to send a real test.');
  return issues;
}

function testFullFlow() {
  Logger.log('=== NEXUS Full Flow Test ===');
  try {
    Logger.log('Gmail quota: ' + MailApp.getRemainingDailyQuota() + ' — scope OK');
  } catch(e) {
    Logger.log('GMAIL ERROR: ' + e.message + ' — re-authorize the script');
  }
  try {
    const enabled = isAutomationEnabled('email_assignments');
    Logger.log('email_assignments: ' + (enabled ? 'ENABLED' : 'DISABLED — fix in Automation Controls sheet'));
  } catch(e) {
    Logger.log('Automation read error: ' + e.message);
  }
  try {
    GmailApp.sendEmail(FOUNDER_EMAIL, '[NEXUS Test] Email check', 'GmailApp is authorized and working.\n\n— NEXUS testFullFlow()', { name:'NEXUS · ABION' });
    Logger.log('Test email sent to ' + FOUNDER_EMAIL + ' — check inbox');
  } catch(e) {
    Logger.log('Test email FAILED: ' + e.message);
  }
  try {
    const resendKey = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');
    if (!resendKey) {
      Logger.log('Resend check SKIPPED — RESEND_API_KEY not set in Script Properties');
    } else {
      const ok = sendViaResend(FOUNDER_EMAIL, '[NEXUS Test] Resend check', '<p>Resend is authorized and working.</p><p>— NEXUS testFullFlow()</p>', 'Resend is authorized and working.\n\n— NEXUS testFullFlow()');
      Logger.log('Resend test email ' + (ok ? 'sent to ' + FOUNDER_EMAIL + ' — check inbox' : 'FAILED — see sendViaResend log above'));
    }
  } catch(e) {
    Logger.log('Resend test error: ' + e.message);
  }
  try {
    PropertiesService.getScriptProperties().setProperty('nexus_test', 'ok');
    const v = PropertiesService.getScriptProperties().getProperty('nexus_test');
    Logger.log('PropertiesService (token storage): ' + (v === 'ok' ? 'OK' : 'FAIL'));
  } catch(e) {
    Logger.log('PropertiesService ERROR: ' + e.message);
  }
  Logger.log('=== Done ===');
}

function testAssignmentEmail() {
  // Uses the first real task with an assignee email from the sheet.
  // Falls back to the script runner's email if none found.
  const tasks = sheetToObjects('1. Master Tasks');
  const realTask = tasks.find(t => t['Assigned_To_Email']);
  const testTask = realTask || {
    'ID'               : 'TEST-001',
    'Task Description' : 'Test Task — please ignore',
    'Assigned To'      : 'Test User',
    'Assigned_To_Email': Session.getActiveUser().getEmail(),
    'Priority'         : 'HIGH',
    'Business / Dept'  : 'Operations',
    'Support Person'   : 'Ayo T. Joshua',
    'Deadline'         : new Date(Date.now() + 7 * 86400000).toISOString().split('T')[0],
    'Notes'            : 'This is a test of the assignment email system.',
  };
  Logger.log('Sending test assignment email to: ' + testTask['Assigned_To_Email']);
  sendAssignmentEmail(testTask);
  Logger.log('Done — check inbox and execution logs.');
}
// Exercises the multi-assignee / multi-support fan-out specifically: two
// assignee names and two support names, but all mapped to the same running
// user's email (so this only needs one real inbox to test) — since every
// support email matches an assignee email, CC should end up completely
// empty. Swap in real distinct addresses to verify actual multi-recipient
// delivery.
function testMultiAssignmentEmail() {
  const you = Session.getActiveUser().getEmail();
  const testTask = {
    'ID'                    : 'TEST-MULTI-001',
    'Task Description'      : 'Test Task (multi-assignee) — please ignore',
    'Assigned To'           : 'Test User A, Test User B',
    'Assigned_To_Email'     : you + ', ' + you,
    'Priority'              : 'HIGH',
    'Business / Dept'       : 'Operations',
    'Support Person'        : 'Test User B, Test Support C',
    'Support_Person_Email'  : you + ', ' + you,
    'Deadline'              : new Date(Date.now() + 7 * 86400000).toISOString().split('T')[0],
    'Notes'                 : 'This is a test of the multi-assignee fan-out.',
  };
  Logger.log('Sending test multi-assignee email — to: ' + testTask['Assigned_To_Email']);
  sendAssignmentEmail(testTask);
  Logger.log('Done — check inbox (should be one email, no duplicate CC) and execution logs.');
}
// Manually exercises the branded deadline email — this fires for real now
// whenever a task's Deadline is set/changed on an existing task independent
// of reassignment (see updateRow). Not fired at initial task creation:
// sendAssignmentEmail already covers the deadline + deep link there, so
// wiring both would double-email every new task.
function testDeadlineEmail() {
  const you = Session.getActiveUser().getEmail();
  const testTask = {
    'ID'                : 'TEST-DEADLINE-001',
    'Task Description'  : 'Test Task (deadline email) — please ignore',
    'Assigned To'       : 'Test User A',
    'Assigned_To_Email' : you,
    'Priority'          : 'HIGH',
    'Deadline'          : new Date().toISOString().split('T')[0],
    'Deadline_Time'     : '15:00',
  };
  Logger.log('Sending test deadline email to: ' + testTask['Assigned_To_Email']);
  sendDeadlineEmail(testTask, [you]);
  Logger.log('Done — check inbox and execution logs.');
}
function dailyOverdueCheck()     { triggerOverdueCheck(); }
function dailyDependencyCheck()  { updateDependencyStatuses(); }
function dailyRecurringTasks()   { createRecurringTasks(); }
function dailyFollowUpEmails()   { runFollowUpAutomation(); }  // ← NEW — was missing
function dailyEODNudge()         { checkMissingKPI(); }        // ← FIXED — now sends Gmail directly

function triggerOverdueCheck() {
  const today   = new Date();
  const overdue = sheetToObjects('1. Master Tasks').filter(t => {
    if (!t['Assigned_To_Email'])                    return false;
    if (['Done','Cancelled'].includes(t['Status'])) return false;
    const ref  = t['Last_Updated'] || t['Start Date'];
    if (!ref)                                        return false;
    const days = Math.floor((today - new Date(ref)) / 86400000);
    t._daysSince = days;
    return [2,5,7].includes(days);
  });
  if (overdue.length && isAutomationEnabled('n8n_webhooks')) {
    pingN8N(N8N.OVERDUE, { tasks:overdue, triggeredAt:today.toISOString() });
  }
  return json({ success:true, overdueCount:overdue.length });
}

// ═════════════════════════════════════
//  ✅ INSTALL TRIGGERS — run this ONCE manually
//  Extensions → Apps Script → Run → installTriggers
// ═════════════════════════════════════

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // dailyDeadlineCheck's automatic run was removed — its deadline-proximity
  // logic is now merged into dailyFollowUpEmails' tier system (see
  // ESCALATION_TIERS), so running both daily would double-notify. The
  // function itself is untouched and still callable manually via the
  // leader-only "sendDeadlineReport" action.
  ScriptApp.newTrigger('dailyOverdueCheck')        .timeBased().everyDays(1).atHour(8) .create(); // 8 AM n8n ping
  ScriptApp.newTrigger('sendDailyMorningDigest')   .timeBased().everyDays(1).atHour(8) .create(); // 8 AM personal digest (Mon=full, Tue-Fri=light)
  ScriptApp.newTrigger('dailyDependencyCheck')     .timeBased().everyDays(1).atHour(9) .create(); // 9 AM dependency sweep
  ScriptApp.newTrigger('dailyRecurringTasks')      .timeBased().everyDays(1).atHour(10).create(); // 10 AM recurring gen
  ScriptApp.newTrigger('sendCOBReminders')         .timeBased().everyDays(1).atHour(16).create(); // 4 PM COB personal reminder to unfiled members
  ScriptApp.newTrigger('dailyFollowUpEmails')      .timeBased().everyDays(1).atHour(18).create(); // 6 PM priority-scaled escalation tiers (staleness + deadline merged)
  ScriptApp.newTrigger('dailyEODNudge')            .timeBased().everyDays(1).atHour(18).create(); // 6 PM missing EOD summary to Founder
  // Hourly follow-ups now installed here too — previously required a separate
  // manual installHourlyTrigger() call, and since this function deletes ALL
  // project triggers on every run, that hourly one kept getting silently
  // wiped out on redeploy without ever being recreated. Folding it in fixes that.
  ScriptApp.newTrigger('runHourlyFollowUps')       .timeBased().everyHours(1).create(); // hourly — only pings tasks with Follow_Up_Hours set
  // Monday digest now handled by sendDailyMorningDigest() → calls sendMondayDigest() on Mondays

  Logger.log([
    'Triggers installed:',
    '  dailyOverdueCheck       @ 8am',
    '  sendDailyMorningDigest  @ 8am  ← NEW (Mon=full digest, Tue-Fri=light digest, to each member)',
    '  dailyDependencyCheck    @ 9am',
    '  dailyRecurringTasks     @ 10am',
    '  sendCOBReminders        @ 4pm  ← NEW (personal EOD reminder to each unfiled member)',
    '  dailyFollowUpEmails     @ 6pm  ← priority-scaled escalation tiers, staleness + deadline merged (see ESCALATION_TIERS), email + push',
    '  dailyEODNudge           @ 6pm  ← EOD summary to Founder/COO',
    '  runHourlyFollowUps      @ every hour  ← only pings tasks with "Follow_Up_Hours" set (8am-7pm, weekdays)',
    '',
    'dailyDeadlineCheck no longer runs automatically — merged into dailyFollowUpEmails above. Still callable manually via the leader "Send deadline report" action.',
  ].join('\n'));
}

// ═════════════════════════════════════
//  ONE-TIME: remap legacy status values + backfill % Progress
// ═════════════════════════════════════
// Run once, after ensureProgressColumns(). Only COS's old 'Pending' value
// isn't in the new ladder (Partnerships' old 4 values are all still valid
// in the new 6-value ladder, nothing to remap there). Also fills
// % Progress on every existing row to match its current status.
function migrateOldStatusValues() {
  const LEGACY = { '2. COS': { 'Pending': 'Not Started' } };
  ['1. Master Tasks', '2. COS', '3. Partnerships'].forEach(name => {
    const sheet = SS.getSheetByName(name);
    if (!sheet) return;
    const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2) return;
    const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
    const statusField = PROGRESS_STATUS_FIELD[name];
    const statusIdx = headers.indexOf(statusField);
    const progIdx   = headers.indexOf('% Progress');
    if (statusIdx === -1 || progIdx === -1) return;
    const rows = sheet.getRange(2,1,lastRow-1,lastCol).getValues();
    const remap = LEGACY[name] || {};
    rows.forEach((row, i) => {
      let status = row[statusIdx];
      if (remap[status]) {
        status = remap[status];
        sheet.getRange(i+2, statusIdx+1).setValue(status);
      }
      const computed = statusToProgress(name, status);
      if (computed !== null) sheet.getRange(i+2, progIdx+1).setValue(computed);
    });
  });
}

// ═════════════════════════════════════
//  ONE-TIME: add missing '% Progress' column to COS/Partnerships
// ═════════════════════════════════════
// Run once manually from the Apps Script editor (select function, Run).
// No-ops on tabs that already have the column. Appends as the last column;
// COLUMNS_BY_TAB looks it up by header name, so exact position doesn't matter.
function ensureProgressColumns() {
  ['2. COS', '3. Partnerships'].forEach(name => {
    const sheet = SS.getSheetByName(name);
    if (!sheet) return;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
    if (headers.indexOf('% Progress') === -1) {
      sheet.getRange(1, lastCol+1).setValue('% Progress');
    }
  });
}

// ═════════════════════════════════════
//  SHEET FORMATTER
// ═════════════════════════════════════

function formatABIONSheet() {
  const RED_MED='#B71C1D',RED_LIGHT='#FFEBEE',WHITE='#FFFFFF',
        GREEN='#E8F5E9',BLUE='#E3F2FD',YELLOW='#FFF8E1',
        ORANGE='#FFF3E0',GRAY='#F5F5F5';

  function fmt(name, validations, cfRules) {
    const sheet = SS.getSheetByName(name); if (!sheet) return;
    const lastCol = sheet.getLastColumn(); if (!lastCol) return;
    const hdrRange = sheet.getRange(1,1,1,lastCol);
    hdrRange.setBackground(RED_MED).setFontColor(WHITE).setFontWeight('bold').setFontSize(10);
    sheet.setFrozenRows(1);
    if (sheet.getLastRow() > 1) sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).setFontSize(10);
    const headers = hdrRange.getValues()[0];
    validations.forEach(v => {
      const ci = headers.indexOf(v.h); if (ci < 0) return;
      sheet.getRange(2,ci+1,Math.max(sheet.getLastRow()-1,1),1)
           .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(v.vals,true).setAllowInvalid(false).build());
    });
    cfRules.forEach(cf => {
      const ci = headers.indexOf(cf.h); if (ci < 0) return;
      const range = sheet.getRange(2,ci+1,Math.max(sheet.getLastRow()-1,1),1);
      const rules = sheet.getConditionalFormatRules();
      cf.rules.forEach(r => rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(r.v).setBackground(r.bg).setFontColor(r.fc||'#333')
          .setRanges([range]).build()
      ));
      sheet.setConditionalFormatRules(rules);
    });
    sheet.autoResizeColumns(1, lastCol);
  }

  fmt('1. Master Tasks',
    [{h:'Priority',vals:['URGENT','HIGH','MEDIUM','LOW']},
     {h:'Status',vals:['Not Started','Acknowledged','In Progress','Revisions Needed','In Review','Done','Blocked','Cancelled']}],
    [{h:'Status',rules:[{v:'Not Started',bg:RED_LIGHT,fc:'#8B0000'},{v:'Acknowledged',bg:YELLOW},{v:'In Progress',bg:BLUE},{v:'Revisions Needed',bg:ORANGE},{v:'In Review',bg:YELLOW},{v:'Done',bg:GREEN},{v:'Blocked',bg:ORANGE},{v:'Cancelled',bg:GRAY,fc:'#999'}]},
     {h:'Priority',rules:[{v:'URGENT',bg:'#D32F2F',fc:WHITE},{v:'HIGH',bg:'#FF9800',fc:WHITE},{v:'MEDIUM',bg:'#FFC107'},{v:'LOW',bg:'#8BC34A',fc:WHITE}]}]);

  fmt('2. COS',
    [{h:'WhatsApp Group Created',vals:['Yes','No']},{h:'Status',vals:['Not Started','Acknowledged','In Progress','Follow-Up Needed','Active']}],
    [{h:'Status',rules:[{v:'Not Started',bg:RED_LIGHT,fc:'#8B0000'},{v:'Acknowledged',bg:YELLOW},{v:'In Progress',bg:BLUE},{v:'Follow-Up Needed',bg:ORANGE},{v:'Active',bg:GREEN}]}]);

  fmt('3. Partnerships',
    [{h:'Category',vals:['NGO','Government','Corporate','Individual','Academic','Media']},
     {h:'Current Status',vals:['Not Started','In Discussion','Negotiating','Cold','Active','Declined']},
     {h:'Document Status',vals:['Not Started','Draft Sent','Under Review','Signed','N/A']}],
    [{h:'Current Status',rules:[{v:'Not Started',bg:RED_LIGHT,fc:'#8B0000'},{v:'In Discussion',bg:BLUE},{v:'Negotiating',bg:YELLOW},{v:'Cold',bg:GRAY,fc:'#999'},{v:'Active',bg:GREEN},{v:'Declined',bg:RED_LIGHT,fc:'#8B0000'}]}]);

  fmt('4. Org Database',
    [{h:'Onboarding Status',vals:['Pending','In Progress','Active','Declined']},
     {h:'WhatsApp Added',vals:['Yes','No']},
     {h:'Letter Status',vals:['Not Sent','Sent','Acknowledged']}],[]);

  fmt('5. KPI Log',
    [{h:'Day',vals:['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']},{h:'Report Submitted (Y/N)',vals:['Y','N']}],
    [{h:'Report Submitted (Y/N)',rules:[{v:'Y',bg:GREEN},{v:'N',bg:RED_LIGHT,fc:'#8B0000'}]}]);

  fmt('6. Marketing',
    [{h:'Content Type',vals:['Reel','Post','Story','Thread','Article','Infographic','Video','Podcast','Other']},
     {h:'Status',vals:['Not Started','In Progress','In Review','Done','Cancelled']}],
    [{h:'Status',rules:[{v:'Done',bg:GREEN},{v:'In Progress',bg:BLUE},{v:'In Review',bg:YELLOW},{v:'Cancelled',bg:GRAY,fc:'#999'}]}]);

  fmt('8. Automation Controls',
    [{h:'Enabled',vals:['TRUE','FALSE']}],
    [{h:'Enabled',rules:[{v:'TRUE',bg:GREEN},{v:'FALSE',bg:RED_LIGHT,fc:'#8B0000'}]}]);

  const colors = {
    '1. Master Tasks':'#B71C1D','2. COS':'#FF6F00','3. Partnerships':'#1565C0',
    '4. Org Database':'#6A1B9A','5. KPI Log':'#00838F','6. Marketing':'#2E7D32',
    '8. Automation Controls':'#424242',
  };
  SS.getSheets().forEach(s => { if (colors[s.getName()]) s.setTabColor(colors[s.getName()]); });
}

// ═════════════════════════════════════
//  UTILS
// ═════════════════════════════════════

// Returns 'YYYY-MM-DD' in the script's LOCAL timezone (not UTC).
// Fixes KPI date bug — toISOString() returns UTC which is 1h behind Nigeria (UTC+1).
// ─── ONE-TIME MIGRATION ────────────────────────────────────────
// Run this once from the Apps Script editor (select it in the function
// dropdown → Run) after deploying this update. It safely adds the new
// "Support_Person_Email" column to the existing "1. Master Tasks" sheet
// without touching any existing rows/data. Safe to run more than once.
function migrateAddSupportPersonEmailColumn() {
  const sheet = SS.getSheetByName('1. Master Tasks');
  if (!sheet) { Logger.log('1. Master Tasks sheet not found.'); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.includes('Support_Person_Email')) { Logger.log('Column already exists — nothing to do.'); return; }
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Support_Person_Email').setFontWeight('bold');
  Logger.log('Added "Support_Person_Email" column to 1. Master Tasks.');
}

// Run once. Adds "Supervisor" and "Supervisor_Email" columns — a distinct
// oversight role, separate from Assigned To / Support Person (single-value,
// unlike those two). Several places in this file already referenced a
// "Supervisor" concept in name only (a stray, unused entry in
// createRecurringTasks' column map, and a local variable in the reminder-tier
// email code that was actually just aliasing Support Person) — this makes it
// a real field for the first time.
function migrateAddSupervisorColumns() {
  const sheet = SS.getSheetByName('1. Master Tasks');
  if (!sheet) { Logger.log('1. Master Tasks sheet not found.'); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ['Supervisor', 'Supervisor_Email'].forEach(col => {
    if (headers.includes(col)) { Logger.log('Column "' + col + '" already exists — skipping.'); return; }
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col).setFontWeight('bold');
    Logger.log('Added "' + col + '" column to 1. Master Tasks.');
  });
}

// One-time migration: adds the Calendar_Event_ID column used to track each
// task's deadline calendar event, so addDeadlineToCalendar can find and
// update/delete a prior event instead of creating a duplicate one every
// time a task is reassigned or its deadline is edited. Safe to re-run —
// skips if the column already exists.
function migrateAddCalendarEventIdColumn() {
  const sheet = SS.getSheetByName('1. Master Tasks');
  if (!sheet) { Logger.log('1. Master Tasks sheet not found.'); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.includes('Calendar_Event_ID')) { Logger.log('Column "Calendar_Event_ID" already exists — skipping.'); return; }
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Calendar_Event_ID').setFontWeight('bold');
  Logger.log('Added "Calendar_Event_ID" column to 1. Master Tasks.');
}

// One-time migration: adds the "Status" column to "9. Team Members" (the
// Pending/Active invite gate — see TEAM_INVITE_PROP comment). Every EXISTING
// row is backfilled to "Active" rather than left blank, since those people
// already have working logins today; only rows added after this migration
// via addTeamMember start life as "Pending". Safe to re-run — skips if the
// column already exists, and only fills genuinely blank Status cells.
function migrateAddTeamStatusColumn() {
  const sheet = SS.getSheetByName('9. Team Members');
  if (!sheet) { Logger.log('9. Team Members sheet not found.'); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let colIdx = headers.indexOf('Status');
  if (colIdx === -1) {
    colIdx = headers.length;
    sheet.getRange(1, colIdx + 1).setValue('Status').setFontWeight('bold');
    Logger.log('Added "Status" column to 9. Team Members.');
  } else {
    Logger.log('"Status" column already exists — checking for blank rows to backfill.');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const statusRange = sheet.getRange(2, colIdx + 1, lastRow - 1, 1);
  const values = statusRange.getValues();
  let backfilled = 0;
  const updated = values.map(row => {
    if (String(row[0] || '').trim() === '') { backfilled++; return ['Active']; }
    return row;
  });
  if (backfilled > 0) {
    statusRange.setValues(updated);
    Logger.log('Backfilled ' + backfilled + ' existing row(s) to Status=Active.');
  }
}

function localDateISO(d) {
  d = d || new Date();
  return d.getFullYear() + '-' +
    String(d.getMonth()+1).padStart(2,'0') + '-' +
    String(d.getDate()).padStart(2,'0');
}

// Sheets stores "Time"-formatted cells as a Date whose date portion is the
// Sheets/Lotus epoch (1899-12-30) with only the time-of-day meaningful.
// Detect that epoch and format as "HH:mm" instead of running it through
// localDateISO(), which was collapsing every Date — including these
// time-only cells — down to just "1899-12-30" and shipping that straight
// into <input type="time"> on the frontend (hence the browser's
// "does not conform to the required format HH:mm" errors).
function isSheetsTimeEpoch(d) {
  return d.getFullYear() === 1899 && d.getMonth() === 11 && d.getDate() === 30;
}
function localTimeHM(d) {
  return String(d.getHours()).padStart(2,'0') + ':' + String(d.getMinutes()).padStart(2,'0');
}

function sheetToObjects(tabName) {
  const sheet = SS.getSheetByName(tabName);
  if (!sheet || sheet.getLastColumn() === 0 || sheet.getLastRow() < 2) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows
    .filter(r => r.some(c => c !== ''))
    .map((r, ri) => {
      const obj = { _rowIndex: ri+1 };
      headers.forEach((h, i) => {
        // Use localDateISO() not toISOString() — toISOString() is UTC and shifts
        // the date back by 1 hour in Nigeria (UTC+1), causing e.g. "2026-07-15"
        // stored in the sheet to appear as "2026-07-14" everywhere in the app.
        const cell = r[i];
        if (cell instanceof Date) {
          obj[h] = isSheetsTimeEpoch(cell) ? localTimeHM(cell) : localDateISO(cell);
        } else {
          obj[h] = cell;
        }
      });
      return obj;
    });
}

function getOrCreateSheet(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function clearData(sheet) {
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2,1,last-1,sheet.getLastColumn()).clearContent();
}

function pingN8N(url, payload) {
  if (!url || url.startsWith('PASTE_')) return;
  try {
    UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
  } catch(e) { Logger.log('n8n ping failed: ' + e.message); }
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
// NEW (v3.2) — run once from the "NEXUS Setup" menu in the Google Sheet
// (or manually from the Apps Script editor). Creates the '19. Goals' sheet
// if it doesn't exist yet, and backfills the two new columns onto
// '16. Pending AI Tasks' if that sheet already has rows from before this
// feature. Safe to re-run — every step checks before it writes.
function setupGoalDecomposition() {
  getOrCreateSheet('19. Goals', TAB_HEADERS['19. Goals']);
  migrateAddGoalColumns();
  SpreadsheetApp.getUi().alert(
    'Goal Decomposition set up',
    "'19. Goals' is ready and '16. Pending AI Tasks' has its Goal_ID / Depends_On_Index columns. You're good to go — try \"Describe a Goal\" in the app.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// One-time migration for existing installs — adds Goal_ID and
// Depends_On_Index to '16. Pending AI Tasks' live, same pattern as the
// other migrateAdd* functions elsewhere in this file. Safe to re-run;
// skips any column that's already there.
function migrateAddGoalColumns() {
  const sheet = SS.getSheetByName('16. Pending AI Tasks');
  if (!sheet) { Logger.log('16. Pending AI Tasks sheet not found — nothing to migrate yet, it will be created fresh with the new columns.'); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ['Goal_ID', 'Depends_On_Index'].forEach(col => {
    if (headers.includes(col)) { Logger.log('Column "' + col + '" already exists — skipping.'); return; }
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col).setFontWeight('bold');
    Logger.log('Added "' + col + '" column to 16. Pending AI Tasks.');
  });
}

function onOpen() {
  // NEW (v3.2) — one-click Goal Decomposition setup, so this ships without
  // any manual sheet/column work: open the Sheet, click the menu item once.
  // Safe to click more than once — everything it calls is idempotent.
  SpreadsheetApp.getUi()
    .createMenu('NEXUS Setup')
    .addItem('Set Up Goal Decomposition', 'setupGoalDecomposition')
    .addToUi();

  // 1. Tell the script to look at your new "sys_config" tab
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sys_config");
  
  // Safety check: in case you accidentally deleted the tab or misspelled it
  if (!sheet) return; 
 
  // 2. Look at cell A1 on that tab
  var killSwitch = sheet.getRange("A1").getValue(); 

  if (killSwitch === "DOWN") {
    // Show the scary system error
    SpreadsheetApp.getUi().alert(
      "CRITICAL SYSTEM ERROR (Code: 503-CONN-FAIL)",
      "The connection to the central database has been interrupted. " +
      "Active sessions have been suspended to prevent data corruption. " +
      "\n\nStandard recovery protocol is underway. Please close this tab and contact your administrator.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Lock them out with the sidebar
    var html = HtmlService.createHtmlOutput(
      "<div style='font-family: sans-serif; text-align: center; padding: 50px; color: #721c24; background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 5px;'>" +
      "<h3>[SYSTEM LOCKDOWN]</h3>" +
      "<p>Active database maintenance in progress. All writes are blocked.</p>" +
      "</div>"
    )
    .setTitle("Database Disconnected");
    
    SpreadsheetApp.getUi().showSidebar(html);
  }
}
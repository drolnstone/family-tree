/**
 * ═══════════════════════════════════════════════════════════════════════════
 * FAMILY HERITAGE TREE — Apps Script backend
 * Build Once. Grow Forever.
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * This script does two jobs:
 *
 *   1. It gives the spreadsheet a "Family Tree" menu with the tools a sole
 *      administrator needs — assigning Person IDs, recalculating generations,
 *      checking the record for mistakes, approving photographs, and backing
 *      everything up.
 *
 *   2. It runs a small web app that lets family members submit photographs and
 *      stories from the website. Everything they send lands in an approval
 *      queue; nothing appears on the public site until you approve it.
 *
 * The website does NOT need this script in order to display the family tree —
 * it reads the published spreadsheet directly. If this script ever breaks, the
 * tree keeps working. That is deliberate.
 *
 * Installation is covered step by step in SETUP.md.
 */

// ─── Settings ──────────────────────────────────────────────────────────────

/** Leave blank when the script is bound to the spreadsheet (the normal case). */
const SHEET_ID = '';

/** Name of the top-level Google Drive folder holding the family media. */
const DRIVE_ROOT_NAME = 'FAMILY TREE';

/** Largest upload accepted from the website, in megabytes. */
const MAX_UPLOAD_MB = 12;

/** How many dated backups to keep before the oldest is discarded. */
const KEEP_BACKUPS = 30;

const TABS = ['PEOPLE','RELATIONSHIPS','PLACES','OCCUPATIONS','EDUCATION',
              'EVENTS','PHOTOS','STORIES','SETTINGS'];

/**
 * This is a blood-descent tree. A person hangs off the family by ParentID — one
 * ungendered column. Parent2ID is only for the rare case where both parents are
 * on record, such as the originating couple. FatherID and MotherID are still
 * read so that a sheet which has not yet been renamed keeps working.
 */
const PARENT_FIELDS = ['FatherID', 'MotherID', 'ParentID', 'Parent2ID'];

/**
 * True for someone who married into the family rather than descending from it.
 *
 * Marked either by Bloodline = No or by an S-prefixed Person ID (S001, S002 …),
 * which is the convention this script uses when it creates a spouse. Accepting
 * either means a row added by hand with an S id stays out of the line of
 * descent even if nobody remembered to fill in Bloodline.
 */
function inLine_(id) { return /^P/i.test(String(id || '').trim()); }

function marriedIn_(r) {
  if (!r) return false;
  return /^no$/i.test(String(r.Bloodline || '').trim()) ||
         /^S\d/i.test(String(r.PersonID || '').trim());
}

/** Every parent of a row that actually exists and belongs to the line. */
/** Both parents as recorded — a wife who married in is still a mother. */
function allParentsOfRow_(r, byId) {
  const out = [];
  PARENT_FIELDS.forEach(function (f) {
    const v = String(r[f] || '').trim();
    if (v && v !== r.PersonID && byId[v] && out.indexOf(v) === -1) out.push(v);
  });
  return out;
}

/** Only the parent who carries the line. Generations and lines follow this. */
function parentsOfRow_(r, byId) {
  return allParentsOfRow_(r, byId).filter(function (x) { return inLine_(x); });
}

/** The parent column to write into: ParentID if the sheet has been renamed. */
function primaryParentField_(header) {
  return header.indexOf('FatherID') >= 0 ? 'FatherID' : 'ParentID';
}
function secondParentField_(header) {
  return header.indexOf('MotherID') >= 0 ? 'MotherID' : 'Parent2ID';
}

/** The column a parent of this gender belongs in. */
function parentFieldFor_(header, gender) {
  return String(gender || '').toUpperCase() === 'F'
    ? secondParentField_(header) : primaryParentField_(header);
}

// ─── Spreadsheet helpers ───────────────────────────────────────────────────

function book_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = book_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name + '. Has a tab been renamed?');
  return sh;
}

/** Read a tab as { header:[], rows:[{}], sheet } — blank rows dropped. */
function table_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  if (!values.length) return { header: [], rows: [], sheet: sh };
  const header = values[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (var r = 1; r < values.length; r++) {
    if (values[r].every(function (v) { return String(v).trim() === ''; })) continue;
    const o = { _row: r + 1 };
    header.forEach(function (h, i) { o[h] = values[r][i] == null ? '' : String(values[r][i]).trim(); });
    rows.push(o);
  }
  return { header: header, rows: rows, sheet: sh };
}

function appendRow_(name, obj) {
  const t = table_(name);
  const line = t.header.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  t.sheet.appendRow(line);
  return t.sheet.getLastRow();
}

/** Next free ID for a column, e.g. nextId_('PEOPLE','PersonID','P',3) → 'P020'. */
function nextId_(tab, col, prefix, width) {
  const rows = table_(tab).rows;
  var max = 0;
  rows.forEach(function (r) {
    const m = String(r[col] || '').match(new RegExp('^' + prefix + '(\\d+)$'));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  var n = String(max + 1);
  while (n.length < width) n = '0' + n;
  return prefix + n;
}

function settings_() {
  const out = {};
  table_('SETTINGS').rows.forEach(function (r) { if (r.Key) out[r.Key] = r.Value; });
  return out;
}

function writeSetting_(key, value) {
  const t = table_('SETTINGS');
  const iKey = t.header.indexOf('Key') + 1, iVal = t.header.indexOf('Value') + 1;
  const hit = t.rows.filter(function (r) { return r.Key === key; })[0];
  if (hit) t.sheet.getRange(hit._row, iVal).setValue(value);
  else t.sheet.appendRow([key, value, '']);
}

// ─── Web app: reading ──────────────────────────────────────────────────────

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET ?action=data      → every tab as JSON (a fallback if the published
 *                         spreadsheet is ever unreachable from the website)
 * GET ?action=pending   → the approval queue, for the administrator
 * GET (anything else)   → a short status page
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'data') {
      const out = {};
      TABS.forEach(function (t) {
        try {
          var rows = table_(t).rows.map(function (r) { delete r._row; return r; });
          // Same rule as the website: nothing unapproved leaves this endpoint.
          if (t === 'PHOTOS' || t === 'STORIES') {
            rows = rows.filter(function (r) {
              const st = String(r.ApprovalStatus || '').trim();
              return !st || /^approved$/i.test(st);
            });
          }
          out[t] = rows;
        } catch (err) { out[t] = []; }
      });
      return json_(out);
    }
    if (action === 'pending') {
      return json_({ waiting: inboxRows_().rows.filter(function (r) { return /^pending$/i.test(r.Status); }) });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
  return HtmlService.createHtmlOutput(
    '<p style="font:15px system-ui;padding:24px">Family Heritage Tree backend is running.</p>');
}

// ─── Web app: writing ──────────────────────────────────────────────────────

/**
 * POST a JSON body. Supported actions:
 *   uploadPhoto  { personId, filename, mimeType, data(base64), caption, photoDate, place, uploader }
 *   addStory     { personId, title, story, toldBy }
 *   suggest      { personId, message, from }
 */
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Could not read the submission.' }); }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    switch (body.action) {
      case 'uploadPhoto': return json_(uploadPhoto_(body));
      case 'addStory':    return json_(addStory_(body));
      case 'suggestPerson':     return json_(suggestPerson_(body));
      case 'suggestCorrection': return json_(suggestCorrection_(body));
      case 'suggest':     return json_(suggestCorrection_(body));   // older name, kept working
      default:            return json_({ ok: false, error: 'Unknown action.' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// ─── The approval inbox ────────────────────────────────────────────────────
//
// Submissions from the website land here, NOT in PHOTOS or STORIES.
//
// This matters more than it looks. The website reads whole tabs from the
// published spreadsheet, so anything sitting in PHOTOS or STORIES — even a row
// marked Pending — is downloaded to every visitor's browser and can be read
// with developer tools. Keeping unvetted material in a separate sheet that the
// website never fetches is what makes "nothing appears until you approve it"
// actually true. INBOX must never be added to the TABS list above.

const INBOX = 'INBOX';
const INBOX_HEADER = [
  // shared by every kind of submission
  'When','Kind','PersonID','Person','Title','Body','DriveFileID',
  'PhotoDate','Place','Uploader','Status','Published as',
  // used when someone suggests a person who is missing from the tree
  'Relation','Name','Gender','BirthDate','BirthPlace','DeathDate','DeathPlace','Living','Contact','Spouse'
];

/** Largest number of unreviewed submissions to hold. A crude spam ceiling. */
const MAX_PENDING = 400;

function inbox_() {
  var sh = book_().getSheetByName(INBOX);
  if (!sh) {
    sh = book_().insertSheet(INBOX);
    sh.appendRow(INBOX_HEADER);
    sh.getRange(1, 1, 1, INBOX_HEADER.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#8C6F4A');
    sh.setFrozenRows(1);
    [150, 70, 90, 170, 220, 420, 260, 110, 150, 140, 100, 120]
      .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange('F2:F1000').setWrap(true);
    sh.setTabColor('8C6F4A');
    return sh;
  }
  ensureInboxColumns_(sh);
  return sh;
}

/**
 * An INBOX created by an earlier version of this script will be missing the
 * columns used by person suggestions. Add whatever is absent, on the end, so
 * upgrading never means rebuilding the sheet or losing what is in it.
 */
function ensureInboxColumns_(sh) {
  const width = Math.max(sh.getLastColumn(), 1);
  const have = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
  const missing = INBOX_HEADER.filter(function (h) { return have.indexOf(h) === -1; });
  if (!missing.length) return;
  sh.getRange(1, have.length + 1, 1, missing.length)
    .setValues([missing])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#8C6F4A');
}

function inboxRows_() {
  const sh = inbox_();
  const values = sh.getDataRange().getValues();
  const header = values[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (var r = 1; r < values.length; r++) {
    if (values[r].every(function (v) { return String(v).trim() === ''; })) continue;
    const o = { _row: r + 1 };
    header.forEach(function (h, i) { o[h] = values[r][i] == null ? '' : String(values[r][i]).trim(); });
    rows.push(o);
  }
  return { sheet: sh, header: header, rows: rows };
}

function uploadPhoto_(b) {
  if (!b.personId || !b.data) return { ok: false, error: 'Missing person or file.' };
  if (!table_('PEOPLE').rows.some(function (r) { return r.PersonID === b.personId; }))
    return { ok: false, error: 'Unknown person.' };

  const bytes = Utilities.base64Decode(b.data);
  if (bytes.length > MAX_UPLOAD_MB * 1024 * 1024)
    return { ok: false, error: 'That file is larger than ' + MAX_UPLOAD_MB + ' MB.' };

  const mime = String(b.mimeType || '');
  if (mime.indexOf('image/') !== 0) return { ok: false, error: 'Only image files can be uploaded.' };

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const safe = String(b.filename || 'photo.jpg').replace(/[^\w.\-]+/g, '_').slice(-60);
  const blob = Utilities.newBlob(bytes, mime, b.personId + '_' + stamp + '_' + safe);

  const folder = personFolder_(b.personId);
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // Some Workspace accounts forbid link sharing. The photo is stored safely
    // either way; the administrator can share it by hand at approval time.
  }

  appendInbox_({
    'When': new Date(), 'Kind': 'Photo', 'PersonID': b.personId, 'Person': personName_(b.personId),
    'Title': b.caption || '', 'DriveFileID': file.getId(),
    'PhotoDate': b.photoDate || '', 'Place': b.place || '',
    'Uploader': b.uploader || 'Anonymous', 'Status': 'Pending'
  });
  notifyAdmin_('New photograph awaiting approval',
    (b.uploader || 'Someone') + ' submitted a photograph for ' + personName_(b.personId) + '.');
  return { ok: true };
}

function addStory_(b) {
  if (!b.personId || !b.story) return { ok: false, error: 'Missing person or story.' };
  if (!table_('PEOPLE').rows.some(function (r) { return r.PersonID === b.personId; }))
    return { ok: false, error: 'Unknown person.' };

  appendInbox_({
    'When': new Date(), 'Kind': 'Story', 'PersonID': b.personId, 'Person': personName_(b.personId),
    'Title': b.title || 'Untitled', 'Body': String(b.story).slice(0, 40000),
    'Uploader': b.toldBy || '', 'Status': 'Pending'
  });
  notifyAdmin_('New story awaiting approval',
    (b.toldBy || 'Someone') + ' shared a story about ' + personName_(b.personId) + '.');
  return { ok: true };
}

function pendingCount_() {
  return inboxRows_().rows.filter(function (r) { return /^pending$/i.test(r.Status); }).length;
}

const clip_ = function (v, n) { return String(v == null ? '' : v).trim().slice(0, n); };

/**
 * A family member proposes somebody the tree is missing.
 *
 * They give a name and say how that person connects to someone already on
 * record. Nothing is written to PEOPLE here — it waits in the inbox until the
 * administrator approves it, at which point publishPerson_ creates the row with
 * the parentage wired up properly.
 */
function suggestPerson_(b) {
  const newName = clip_(b.newName, 120);
  if (!newName) return { ok: false, error: 'Please give the person a name.' };
  if (pendingCount_() >= MAX_PENDING)
    return { ok: false, error: 'The review queue is full. Please try again in a few days.' };

  const related = table_('PEOPLE').rows.filter(function (r) { return r.PersonID === clip_(b.relatedId, 20); })[0];
  if (!related) return { ok: false, error: 'Please say who this person is related to.' };
  // Somebody who married in holds no place in the line, so nobody can be
  // attached through them. Refusing here beats writing a parent link that the
  // site will then refuse to follow, leaving the new person stranded.
  if (marriedIn_(related))
    return { ok: false, error: (related.DisplayName || related.PersonID) + ' married into the family, so ' +
             'relatives cannot be added through them. Please choose their husband or wife, or another blood relative.' };

  // Three blood relations, plus marriage — which attaches a person to the
  // family without putting them in the line. Anything else falls back safely.
  var relIn = clip_(b.relation, 20).toLowerCase();
  if (relIn === 'spouse of' || relIn === 'partner of') relIn = 'married to';   // same thing, said differently
  const rel = ['child of', 'sibling of', 'parent of', 'married to'].indexOf(relIn) >= 0 ? relIn : 'child of';
  const gender = ['M', 'F', 'U'].indexOf(clip_(b.gender, 1).toUpperCase()) >= 0
    ? clip_(b.gender, 1).toUpperCase() : 'U';
  const living = ['Yes', 'No', 'Unknown'].indexOf(clip_(b.living, 10)) >= 0 ? clip_(b.living, 10) : 'Unknown';

  const row = {};
  row['When'] = new Date();
  row['Kind'] = 'Person';
  row['PersonID'] = related.PersonID;
  row['Person'] = related.DisplayName || related.PersonID;
  row['Title'] = newName + ' — ' + rel + ' ' + (related.DisplayName || related.PersonID);
  row['Body'] = clip_(b.notes, 8000);
  row['Relation'] = rel;
  row['Name'] = newName;
  row['Gender'] = gender;
  row['BirthDate'] = clip_(b.birthDate, 40);
  row['BirthPlace'] = clip_(b.birthPlace, 120);
  row['DeathDate'] = clip_(b.deathDate, 40);
  row['DeathPlace'] = clip_(b.deathPlace, 120);
  row['Living'] = living;
  row['Spouse'] = clip_(b.spouse, 120);      // who they married, if the contributor knew
  row['Uploader'] = clip_(b.from, 120) || 'Anonymous';
  row['Contact'] = clip_(b.contact, 160);
  row['Status'] = 'Pending';
  appendInbox_(row);

  notifyAdmin_('Someone new suggested for the tree',
    (row['Uploader']) + ' says ' + newName + ' is the ' + rel + ' ' +
    (related.DisplayName || related.PersonID) + '.');
  return { ok: true };
}

/** A family member reports something wrong. Applied by hand, never automatically. */
function suggestCorrection_(b) {
  const message = clip_(b.message, 8000);
  if (!message) return { ok: false, error: 'Please say what should be changed.' };
  if (pendingCount_() >= MAX_PENDING)
    return { ok: false, error: 'The review queue is full. Please try again in a few days.' };

  const who = clip_(b.personId, 20);
  const row = {};
  row['When'] = new Date();
  row['Kind'] = 'Correction';
  row['PersonID'] = who;
  row['Person'] = who ? personName_(who) : '';
  row['Title'] = 'Correction';
  row['Body'] = message;
  row['Uploader'] = clip_(b.from, 120) || 'Anonymous';
  row['Contact'] = clip_(b.contact, 160);
  row['Status'] = 'Pending';
  appendInbox_(row);

  notifyAdmin_('A correction has been suggested',
    row['Uploader'] + ' suggests a change' + (who ? ' to ' + personName_(who) : '') + '.');
  return { ok: true };
}

/** Append by column name, so the order of the inbox columns never matters. */
function appendInbox_(obj) {
  const sh = inbox_();
  const header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  sh.appendRow(header.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; }));
}

function addSuggestion_(b) {
  const t = book_().getSheetByName('SUGGESTIONS') || book_().insertSheet('SUGGESTIONS');
  if (t.getLastRow() === 0)
    t.appendRow(['When', 'PersonID', 'From', 'Message', 'Handled']);
  t.appendRow([new Date(), b.personId || '', b.from || '', b.message || '', 'No']);
  return { ok: true };
}

function personName_(id) {
  const hit = table_('PEOPLE').rows.filter(function (r) { return r.PersonID === id; })[0];
  return hit ? (hit.DisplayName || hit.FullName || id) : id;
}

/** The INBOX sheet's own address, so the email can go straight to it. */
function inboxUrl_() {
  try {
    return 'https://docs.google.com/spreadsheets/d/' + book_().getId() +
           '/edit#gid=' + inbox_().getSheetId();
  } catch (err) { return ''; }
}

const escHtml_ = function (s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
};

function notifyAdmin_(subject, body) {
  const to = settings_().contact_email || Session.getEffectiveUser().getEmail();
  if (!to) return;

  const url = inboxUrl_();
  var waiting = 0;
  try { waiting = pendingCount_(); } catch (err) {}
  const queue = waiting === 1 ? '1 submission is waiting for you.'
              : waiting > 1  ? waiting + ' submissions are waiting for you.' : '';

  const text = [
    body,
    queue,
    url ? 'Open the approval queue:\n' + url : 'Open the spreadsheet and use Family Tree ▸ Review the approval queue.',
    'Select the rows you want, then use Family Tree ▸ Approve the selected rows (or Reject).'
  ].filter(function (x) { return x; }).join('\n\n');

  const html =
    '<div style="font:15px/1.55 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#2b2b2b;max-width:520px">' +
      '<p style="margin:0 0 14px">' + escHtml_(body) + '</p>' +
      (queue ? '<p style="margin:0 0 18px;color:#6b6b6b">' + escHtml_(queue) + '</p>' : '') +
      (url
        ? '<p style="margin:0 0 18px"><a href="' + escHtml_(url) + '" ' +
          'style="display:inline-block;background:#8C6F4A;color:#fff;text-decoration:none;' +
          'padding:11px 20px;border-radius:6px;font-weight:600">Open the approval queue</a></p>'
        : '<p style="margin:0 0 18px">Open the spreadsheet and use <b>Family Tree ▸ Review the approval queue</b>.</p>') +
      '<p style="margin:0;color:#6b6b6b;font-size:13px">Select the rows you want, then use ' +
        '<b>Family Tree ▸ Approve the selected rows</b> (or Reject). Nothing appears on the ' +
        'website until you approve it.</p>' +
    '</div>';

  try {
    MailApp.sendEmail({
      to: to, name: 'Family Tree',
      subject: '[Family Tree] ' + subject,
      body: text, htmlBody: html
    });
  } catch (err) { /* quota exhausted — not worth failing the upload over */ }
}

// ─── Google Drive ──────────────────────────────────────────────────────────

function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function driveRoot_() {
  const id = settings_().drive_folder_id;
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  return childFolder_(DriveApp.getRootFolder(), DRIVE_ROOT_NAME);
}

function personFolder_(personId) {
  return childFolder_(childFolder_(driveRoot_(), 'PEOPLE'), personId);
}

/** Menu ▸ Create the Google Drive folders. */
function createDriveFolders() {
  const root = childFolder_(DriveApp.getRootFolder(), DRIVE_ROOT_NAME);
  childFolder_(root, 'PEOPLE');
  const fam = childFolder_(root, 'FAMILY');
  ['Historical Photographs', 'Documents', 'General Family Archive'].forEach(function (n) {
    childFolder_(fam, n);
  });
  childFolder_(root, 'Backups');
  writeSetting_('drive_folder_id', root.getId());

  var made = 0;
  table_('PEOPLE').rows.forEach(function (r) {
    if (r.PersonID) { childFolder_(childFolder_(root, 'PEOPLE'), r.PersonID); made++; }
  });
  ui_().alert('Drive is ready',
    'Created "' + DRIVE_ROOT_NAME + '" with folders for ' + made + ' people.\n\n' +
    'Folder ID saved to SETTINGS ▸ drive_folder_id:\n' + root.getId(), ui_().ButtonSet.OK);
}

// ─── Administration tools ──────────────────────────────────────────────────

function ui_() { return SpreadsheetApp.getUi(); }

/**
 * The menu, grouped by what you are actually doing rather than by what the
 * code happens to contain:
 *
 *   adding people  →  dealing with what the family has sent  →  tidying up
 *   after a change →  checking →  keeping it safe
 *
 * The things you run once when setting a sheet up, and never again, are folded
 * away under "Set up & repair". They are not deleted, because a rebuilt sheet —
 * or whoever keeps this record after you — will need them exactly once too.
 */
function onOpen() {
  const ui = ui_();
  ui.createMenu('Family Tree')
    .addItem('Add a person…', 'addPersonPrompt')
    .addItem('Add a husband or wife…', 'addSpousePrompt')
    .addSeparator()

    .addItem('Review the approval queue', 'reviewQueue')
    .addItem('Approve the selected rows', 'approveSelection')
    .addItem('Reject the selected rows', 'rejectSelection')
    .addItem('Make the selected photo the profile photo', 'makeProfilePhoto')
    .addSeparator()

    .addItem('Recalculate generations', 'recalcGenerations')
    .addItem('Give IDs to any new rows', 'assignMissingIds')
    .addSeparator()

    .addItem('Check the record for problems', 'validateRecord')
    .addItem('Family statistics', 'showStats')
    .addSeparator()

    .addItem('Back up now', 'backupNow')
    .addSubMenu(ui.createMenu('Set up & repair')
      .addItem('Create the Google Drive folders', 'createDriveFolders')
      .addItem('Nightly backups: on / off…', 'toggleNightlyBackup')
      .addItem('Explain what each column is for', 'annotateColumns')
      .addItem('Setup details for the website', 'showSetupDetails')
      .addSeparator()
      .addItem('Turn spouse names into records', 'spouseNamesToRecords')
      .addItem('Remove the old Spouse name column', 'removeSpouseNameColumn')
      .addItem('Update the sheet to the latest layout', 'updateSheetLayout')
      .addItem('Move old pending rows into the inbox', 'migratePendingToInbox'))
    .addToUi();
}

/**
 * Menu ▸ Add a person — somebody born into the family.
 *
 * This is for the line of descent. A husband or wife is a different thing and
 * has its own command; adding one here would put them in the tree, which is
 * exactly what must not happen.
 */
function addPersonPrompt() {
  const ui = ui_();
  ensureMarriageColumns_();

  const nameAsk = ui.prompt('Add a person',
    'Somebody born into the family. For a husband or wife, cancel this and use\n' +
    'Family Tree ▸ Add a husband or wife… instead.\n\n' +
    'What is their name? (Leave blank for an unknown name.)', ui.ButtonSet.OK_CANCEL);
  if (nameAsk.getSelectedButton() !== ui.Button.OK) return;

  const parentAsk = ui.prompt('Add a person',
    'Person ID of the parent they descend from — the one already in this family.\n' +
    'Leave blank if unknown.', ui.ButtonSet.OK_CANCEL);
  if (parentAsk.getSelectedButton() !== ui.Button.OK) return;

  const t = table_('PEOPLE');
  const byId = {};
  t.rows.forEach(function (r) { byId[r.PersonID] = r; });

  const parentId = parentAsk.getResponseText().trim().toUpperCase();
  if (parentId && !byId[parentId]) throw new Error('There is no person with the ID ' + parentId + '.');
  const parent = byId[parentId];
  if (parent && marriedIn_(parent))
    throw new Error(parentId + ' married into the family, so nobody descends through them. ' +
      'Give the Person ID of the blood parent — you can add ' + parentId +
      ' as the second parent at the next question.');

  // The other parent, usually the mother who married in. Offered rather than
  // required, and only sensible once the blood parent is known.
  var second = '';
  if (parent) {
    const spouses = String(parent.SpouseID || '').split(/[;,]/)
      .map(function (x) { return x.trim(); }).filter(function (x) { return x && byId[x]; });
    const hint = spouses.length
      ? '\n\n' + (parent.DisplayName || parentId) + ' is married to: ' +
        spouses.map(function (s) { return s + ' (' + (byId[s].DisplayName || s) + ')'; }).join(', ')
      : '';
    const secondAsk = ui.prompt('Add a person',
      'Person ID of the other parent — usually the mother, if she is on record.\n' +
      'Leave blank if you do not know, or if she has no record yet.' + hint, ui.ButtonSet.OK_CANCEL);
    if (secondAsk.getSelectedButton() !== ui.Button.OK) return;
    second = secondAsk.getResponseText().trim().toUpperCase();
    if (second && !byId[second]) throw new Error('There is no person with the ID ' + second + '.');
    if (second === parentId) second = '';
  }

  const id = nextId_('PEOPLE', 'PersonID', 'P', 3);
  const row = {
    PersonID: id,
    DisplayName: nameAsk.getResponseText().trim() || 'Unknown',
    Gender: 'U',
    Generation: parent && parent.Generation ? (parseInt(parent.Generation, 10) + 1) : '',
    Living: 'Unknown',
    Privacy: 'Public',
    Status: 'Reported by family'
  };
  row[primaryParentField_(t.header)] = parentId;
  if (second) row[secondParentField_(t.header)] = second;
  appendRow_('PEOPLE', row);
  try { personFolder_(id); } catch (err) { /* Drive can wait */ }

  ui.alert('Added',
    id + ' has been added' +
    (parent ? ', child of ' + (parent.DisplayName || parentId) +
      (second ? ' and ' + (byId[second].DisplayName || second) : '') : '') + '.\n' +
    'A Drive folder has been made for their photographs.\n\n' +
    'Open the PEOPLE sheet to fill in the rest of what you know — dates, places, gender, ' +
    'and SortOrder if you know where they came among their brothers and sisters.',
    ui.ButtonSet.OK);
}

/** Menu ▸ Give IDs to any new rows. */
function assignMissingIds() {
  const specs = [
    ['PEOPLE','PersonID','P'], ['RELATIONSHIPS','RelID','R'], ['PLACES','PlaceRecID','L'],
    ['OCCUPATIONS','OccRecID','O'], ['EDUCATION','EduRecID','E'], ['EVENTS','EventID','V'],
    ['PHOTOS','PhotoID','F'], ['STORIES','StoryID','S']
  ];
  var filled = 0;
  specs.forEach(function (s) {
    const t = table_(s[0]);
    const col = t.header.indexOf(s[1]) + 1;
    if (!col) return;
    t.rows.forEach(function (r) {
      if (!r[s[1]]) {
        t.sheet.getRange(r._row, col).setValue(nextId_(s[0], s[1], s[2], 3));
        filled++;
      }
    });
  });
  ui_().alert('Done', filled ? ('Gave IDs to ' + filled + ' row(s).') : 'Every row already has an ID.', ui_().ButtonSet.OK);
}

/** Menu ▸ Recalculate generations. Generation 1 is anyone with no known parent. */
function recalcGenerations() {
  const t = table_('PEOPLE');
  const col = t.header.indexOf('Generation') + 1;
  if (!col) throw new Error('The PEOPLE sheet has no Generation column.');

  const byId = {};
  t.rows.forEach(function (r) { byId[r.PersonID] = r; });

  const gen = {};
  t.rows.forEach(function (r) {
    if (!marriedIn_(r) && !parentsOfRow_(r, byId).length) gen[r.PersonID] = 1;
  });

  // Repeatedly push generations downward. Bounded so a cycle cannot hang it.
  for (var pass = 0; pass < t.rows.length + 2; pass++) {
    var changed = false;
    t.rows.forEach(function (r) {
      const parents = parentsOfRow_(r, byId).filter(function (p) { return gen[p]; });
      if (!parents.length) return;
      const want = Math.max.apply(null, parents.map(function (p) { return gen[p]; })) + 1;
      if (gen[r.PersonID] !== want) { gen[r.PersonID] = want; changed = true; }
    });
    if (!changed) break;
  }

  var n = 0;
  t.rows.forEach(function (r) {
    const v = marriedIn_(r) ? '' : (gen[r.PersonID] || '');
    if (String(r.Generation) !== String(v)) { t.sheet.getRange(r._row, col).setValue(v); n++; }
  });
  ui_().alert('Generations updated', n + ' row(s) changed.', ui_().ButtonSet.OK);
}



/**
 * The three columns a marriage needs, with the note that gets attached to each
 * heading so the sheet explains itself without anybody having to remember.
 */
const MARRIAGE_COLUMNS = [
  ['SpouseID',
   'Only for a wife or husband who deserves a record of their own — a photograph, where they ' +
   'came from, their story. Put their Person ID here, and give them a row of their own with ' +
   'Bloodline set to No. Recording the marriage on either row is enough; you need not do both.'],
  ['Bloodline',
   'No = this person married into the family. They will never appear in the tree, never count ' +
   'as anybody\'s ancestor or descendant, and cannot be given as a parent. Leave it blank for ' +
   'everyone born into the family.']
];

/**
 * Add whatever is missing of the three, on the end, matching the look of the
 * headings already there. Additive and safe to run again; no existing column,
 * value or Person ID is touched.
 */
function ensureMarriageColumns_() {
  const sh = sheet_('PEOPLE');
  const width = Math.max(sh.getLastColumn(), 1);
  const have = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
  const missing = MARRIAGE_COLUMNS.filter(function (c) { return have.indexOf(c[0]) === -1; });
  if (!missing.length) return [];

  // They belong with the other family columns, immediately after the parents —
  // which is where the sample workbook puts them. Only if there is no parent
  // column to anchor to do they go on the end.
  // Keep the three together: if one is already in place, the others join it.
  const anchor = Math.max(
    have.indexOf('Bloodline'), have.indexOf('SpouseID'), have.indexOf('Spouse'),
    have.indexOf('Parent2ID'), have.indexOf('MotherID'),
    have.indexOf('ParentID'), have.indexOf('FatherID'));
  const at = anchor >= 0 ? anchor + 2 : width + 1;
  if (anchor >= 0) sh.insertColumnsAfter(anchor + 1, missing.length);
  sh.getRange(1, at, 1, missing.length).setValues([missing.map(function (c) { return c[0]; })]);
  // Borrow the formatting of the first heading so the new ones do not look bolted on.
  try { sh.getRange(1, 1).copyFormatToRange(sh, at, at + missing.length - 1, 1, 1); } catch (err) {}

  missing.forEach(function (c, i) {
    sh.getRange(1, at + i).setNote(c[1]);
    try { sh.setColumnWidth(at + i, c[0] === 'Bloodline' ? 100 : 150); } catch (err) {}
  });

  // A Yes/No list on Bloodline, permissive so that a blank cell stays legal —
  // blank is the normal case and means "born into the family".
  const iB = missing.map(function (c) { return c[0]; }).indexOf('Bloodline');
  if (iB >= 0) {
    try {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No'], true).setAllowInvalid(true)
        .setHelpText('No = married into the family. Blank or Yes = born into it.').build();
      sh.getRange(2, at + iB, Math.max(sh.getMaxRows() - 1, 1)).setDataValidation(rule);
    } catch (err) {}
  }
  return missing.map(function (c) { return c[0]; });
}


/**
 * What every column in PEOPLE is for. Attached to the headings as notes, so the
 * sheet answers the question itself instead of anybody having to remember or ask.
 */
const COLUMN_NOTES = {
  'PersonID': 'Their permanent identifier. P001, P002 … for the family by blood; S001, S002 … ' +
    'for a husband or wife who married in. Never reuse one and never change one — every other ' +
    'sheet, and every photograph, points at a person by this.',
  'DisplayName': 'The name the family actually uses. This is what appears on the website.',
  'FullName': 'The full formal name, if it differs from the one above.',
  'OtherNames': 'Other names they were known by — a maiden name, a praise name, a name from ' +
    'another language. These are searchable on the website.',
  'Nickname': 'What they were called.',
  'Gender': 'M, F or U for unknown. Decides the silhouette shown when there is no photograph, ' +
    'and whether a marriage reads Wife or Husband.',
  'ParentID': 'The Person ID of the parent through whom this person belongs to the family. ' +
    'This single column is what builds the whole tree.',
  'Parent2ID': 'The other parent, when they are also on record — usually the mother who ' +
    'married in. Shown on the page, but descent is reckoned through ParentID alone.',
  'SpouseID': 'The Person ID of the husband or wife. Recording the marriage on either row is ' +
    'enough; the website reads it from both sides.',
  'Bloodline': 'No = this person married into the family. They keep a page of their own but ' +
    'hold no place in the line of descent. Leave blank for everyone born into the family.',
  'Branch': 'No longer used. Which line someone belongs to is worked out from the tree. ' +
    'Family Tree ▸ Name the family branches. The website groups and filters people by it. ' +
    'Blank simply means that command has not been run yet.',
  'Generation': 'How many generations down from the earliest known ancestor, counting the ' +
    'ancestor as 1. Filled in by Family Tree ▸ Recalculate generations — do not type it by hand.',
  'BirthDate': 'Any form is understood: 1948, c.1950, March 1948. Only the year is used for ' +
    'sorting and for working out lifespans.',
  'BirthPlace': 'Where they were born.',
  'DeathDate': 'Leave blank for the living.',
  'DeathPlace': 'Where they died.',
  'BurialPlace': 'Where they are buried.',
  'Living': 'Yes, No or Unknown. Anyone marked Yes has their dates and details withheld from ' +
    'the public site while SETTINGS ▸ hide_living_details is on.',
  'Privacy': 'Public, or Private to keep the person off the website altogether except as a ' +
    'position in the tree.',
  'ProfilePhoto': 'Normally left blank. Flag the photograph you want on the PHOTOS sheet ' +
    'instead, using Family Tree ▸ Make the selected photo the profile photo.',
  'ShortBio': 'A paragraph or two, shown at the top of their page.',
  'Status': 'How sure you are: Confirmed by family, Reported by family, Name unknown, ' +
    'Needs clarification. Shown as a small badge on their page.',
  'SortOrder': 'The order brothers and sisters appear in, within one family — 1 for the ' +
    'eldest, 2 for the next, and so on. It is only ever compared between children of the same ' +
    'parent, so the numbers may start again at 1 in every family. Leave it blank and they are ' +
    'ordered by year of birth instead, or by name where no year is known. Worth filling in ' +
    'where you know the birth order but not the dates.',
  'Notes': 'Anything for you, the keeper of the record. Never shown on the website.',
  'Spouse': 'No longer used. A husband or wife now has a record of their own with an S id. ' +
    'Family Tree ▸ Turn spouse names into records converts anything still written here, and ' +
    'the column can then be deleted.'
};

/** Menu ▸ Explain what each column is for. */
function annotateColumns() {
  const ui = ui_();
  const sh = sheet_('PEOPLE');
  const width = Math.max(sh.getLastColumn(), 1);
  const have = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
  var done = 0;
  const unknown = [];
  have.forEach(function (h, i) {
    if (!h) return;
    if (COLUMN_NOTES[h]) { sh.getRange(1, i + 1).setNote(COLUMN_NOTES[h]); done++; }
    else unknown.push(h);
  });
  ui.alert('The columns now explain themselves',
    done + ' heading(s) on the PEOPLE sheet now carry a note saying what belongs in them.\n\n' +
    'Hover over any heading — or tap it on a phone — to read it.' +
    (unknown.length ? '\n\nNo note for: ' + unknown.join(', ') + '. These are columns this script ' +
      'does not know about; the website ignores them, so they are yours to use as you like.' : ''),
    ui.ButtonSet.OK);
}

/**
 * Menu ▸ Remove the old Spouse name column.
 *
 * The column held a husband or wife as bare text, before they had records of
 * their own. It is refused while anything is still written in it, because that
 * text is somebody's name and deleting it would lose them.
 */
function removeSpouseNameColumn() {
  const ui = ui_();
  const t = table_('PEOPLE');
  const i = t.header.indexOf('Spouse');
  if (i < 0) {
    ui.alert('Already gone', 'The PEOPLE sheet has no Spouse column. Nothing to remove.', ui.ButtonSet.OK);
    return;
  }
  const stillThere = t.rows.filter(function (r) { return String(r.Spouse || '').trim(); });
  if (stillThere.length) {
    ui.alert('Not yet — there are still names in it',
      stillThere.length + ' row(s) still have a name in the Spouse column:\n\n' +
      stillThere.slice(0, 10).map(function (r) {
        return '  ' + r.PersonID + ' ' + (r.DisplayName || '') + ' — ' + r.Spouse;
      }).join('\n') +
      (stillThere.length > 10 ? '\n  … and ' + (stillThere.length - 10) + ' more' : '') + '\n\n' +
      'Run Family Tree ▸ Turn spouse names into records first. That gives each of them a page ' +
      'of their own and empties the column, and then this will remove it.', ui.ButtonSet.OK);
    return;
  }
  const go = ui.alert('Remove the Spouse column',
    'The Spouse column is empty, so nothing is lost by removing it. Husbands and wives are now ' +
    'recorded as people in their own right and linked by SpouseID.\n\n' +
    'Column ' + String.fromCharCode(65 + i) + ' will be deleted. Go ahead?', ui.ButtonSet.YES_NO);
  if (go !== ui.Button.YES) return;

  t.sheet.deleteColumn(i + 1);
  ui.alert('Removed',
    'The Spouse column is gone. The website never needed it — it reads marriages from SpouseID.',
    ui.ButtonSet.OK);
}

/** The next free S id, for somebody who married into the family. */
function nextSpouseId_() { return nextId_('PEOPLE', 'PersonID', 'S', 3); }

/**
 * Create the row for a husband or wife: a full member of the family with a page
 * of their own, but no place in the line of descent.
 *
 * The marriage is written on BOTH rows. The website only needs one — it reads a
 * marriage from either side — but a sheet you maintain by hand should say so
 * plainly wherever you happen to be looking, and a partner whose SpouseID stays
 * blank looks unmarried when you scan the column.
 */
function makeSpouseRow_(nm, partner, gender, notes) {
  const id = nextSpouseId_();
  appendRow_('PEOPLE', {
    PersonID: id,
    DisplayName: nm,
    Gender: gender || 'U',
    SpouseID: partner.PersonID,
    Bloodline: 'No',
    Generation: '',
    Living: 'Unknown',
    Privacy: 'Public',
    Status: 'Reported by family',
    Notes: notes || ''
  });
  linkSpouseBack_(partner.PersonID, id);
  try { personFolder_(id); } catch (err) { /* Drive can wait; the record matters more */ }
  return id;
}

/** Write a spouse's id into the partner's own SpouseID cell, without duplicating. */
function linkSpouseBack_(partnerId, spouseId) {
  const t = table_('PEOPLE');
  const i = t.header.indexOf('SpouseID') + 1;
  if (!i) return '';
  const row = t.rows.filter(function (x) { return x.PersonID === partnerId; })[0];
  if (!row) return '';
  const have = String(row.SpouseID || '').split(/[;,]/)
    .map(function (x) { return x.trim(); }).filter(Boolean);
  if (have.indexOf(spouseId) >= 0) return have.join('; ');
  have.push(spouseId);
  t.sheet.getRange(row._row, i).setValue(have.join('; '));
  return have.join('; ');
}

/** The likely gender of a husband or wife, given the gender of who they married. */
function oppositeGender_(g) {
  const G = String(g || '').toUpperCase();
  return G === 'M' ? 'F' : G === 'F' ? 'M' : 'U';
}

/**
 * Menu ▸ Turn spouse names into records.
 *
 * A name typed into the Spouse column is only a name: it has no page, no
 * photographs and no folder. This gives each one a record of their own — an S
 * id, Bloodline = No, and the marriage recorded on both rows — and, where the
 * husband or wife had only the one spouse, writes them in as the second parent
 * of that couple's children so the family is explicit in the sheet rather than
 * guessed at by the website.
 */
function spouseNamesToRecords() {
  const ui = ui_();
  ensureMarriageColumns_();
  const t = table_('PEOPLE');
  const iSpouse = t.header.indexOf('Spouse') + 1;
  const iSpouseID = t.header.indexOf('SpouseID') + 1;
  const SECOND = secondParentField_(t.header);
  const iSecond = t.header.indexOf(SECOND) + 1;

  const byId = {};
  t.rows.forEach(function (x) { if (x.PersonID) byId[x.PersonID] = x; });

  // Every name waiting to become a person.
  const jobs = [];
  t.rows.forEach(function (r) {
    if (!r.PersonID || marriedIn_(r)) return;
    String(r.Spouse || '').split(/;/).forEach(function (nm) {
      nm = nm.trim();
      if (nm) jobs.push({ person: r, name: nm });
    });
  });

  if (!jobs.length) {
    ui.alert('Nothing to convert',
      'No names are waiting in the Spouse column.\n\n' +
      'Type a husband or wife\'s name into Spouse on their partner\'s row, then run this again — ' +
      'or use Family Tree ▸ Add a spouse… to create one directly.', ui.ButtonSet.OK);
    return;
  }

  // How many DIFFERENT spouses each person will end up with, so that children
  // are only attributed where there is no doubt whose they are. Counting
  // identities rather than mentions matters: a marriage written on both rows is
  // still one marriage, and must not look like two.
  const spouseSets = {};
  const noteSpouse = function (who, key) {
    if (!who || !key) return;
    (spouseSets[who] = spouseSets[who] || {})[key] = true;
  };
  jobs.forEach(function (j) { noteSpouse(j.person.PersonID, 'name:' + j.name.toLowerCase()); });
  t.rows.forEach(function (r) {
    String(r.SpouseID || '').split(/[;,]/).forEach(function (x) {
      x = x.trim(); if (!x || !byId[x]) return;
      noteSpouse(r.PersonID, 'id:' + x);          // written on this row
      noteSpouse(x, 'id:' + r.PersonID);          // and so true of the other one
    });
  });
  const spouseCount = function (who) { return Object.keys(spouseSets[who] || {}).length; };

  const preview = jobs.slice(0, 12).map(function (j) {
    return '  ' + j.name + ' — ' + (j.person.DisplayName || j.person.PersonID);
  }).join('\n');
  const go = ui.alert('Turn spouse names into records',
    jobs.length + ' name(s) will each become a person with a page of their own:\n\n' + preview +
    (jobs.length > 12 ? '\n  … and ' + (jobs.length - 12) + ' more' : '') + '\n\n' +
    'Each gets an S id, Bloodline = No, and a link to the person they married. Where someone has\n' +
    'only one husband or wife, that spouse is also written into ' + SECOND + ' on their children,\n' +
    'so the couple\'s children show on both their pages.\n\n' +
    'Nobody is added to the line of descent, no Person ID changes, and the tree is untouched.\n' +
    'The name is cleared from the Spouse column once the record exists.\n\n' +
    'Back up first if you would rather (Family Tree ▸ Back up now). Go ahead?',
    ui.ButtonSet.YES_NO);
  if (go !== ui.Button.YES) return;

  var made = 0, attributed = 0;
  const ambiguous = {}, summary = [];

  jobs.forEach(function (j) {
    const partner = j.person;
    const id = makeSpouseRow_(j.name, partner, oppositeGender_(partner.Gender),
      'Was recorded only as a name on ' + partner.PersonID + '\'s row. Gender assumed from the ' +
      'marriage — please correct it if wrong.');
    made++;

    // makeSpouseRow_ has already written the marriage onto the partner's row;
    // keep this copy of it in step so the count below stays right.
    partner.SpouseID = String(partner.SpouseID || '').trim()
      ? partner.SpouseID + '; ' + id : id;

    // The children of this marriage, but only where there is one wife or husband
    // and so no question of whose children they are.
    if (spouseCount(partner.PersonID) === 1 && iSecond) {
      t.rows.forEach(function (kid) {
        if (!kid.PersonID || kid.PersonID === partner.PersonID) return;
        if (parentsOfRow_(kid, byId).indexOf(partner.PersonID) < 0) return;
        if (String(kid[SECOND] || '').trim()) return;            // never overwrite
        t.sheet.getRange(kid._row, iSecond).setValue(id);
        kid[SECOND] = id;
        attributed++;
      });
    } else if (spouseCount(partner.PersonID) > 1) {
      ambiguous[partner.PersonID] = partner.DisplayName || partner.PersonID;
    }
    summary.push('  ' + id + '  ' + j.name + ' — married ' + (partner.DisplayName || partner.PersonID));
  });

  // Clear the names last, so nothing is lost if the run is interrupted partway.
  if (iSpouse) jobs.forEach(function (j) { t.sheet.getRange(j.person._row, iSpouse).setValue(''); });

  ui.alert('Spouses now have records',
    made + ' record(s) created:\n' + summary.slice(0, 20).join('\n') +
    (summary.length > 20 ? '\n  … and ' + (summary.length - 20) + ' more' : '') + '\n\n' +
    (attributed ? attributed + ' child(ren) now name their mother or father as well as their blood parent.\n' : '') +
    (Object.keys(ambiguous).length
      ? Object.keys(ambiguous).map(function (k) { return ambiguous[k]; }).join(' and ') +
        ' had more than one wife or husband, so their children were left alone — there is no way to ' +
        'tell whose they are. Write the right S id into ' + SECOND + ' on each child by hand.\n'
      : '') +
    '\nGender was guessed from the marriage; correct any that are wrong.\n' +
    'Each of them has a Drive folder for photographs, and the marriage is written on both rows.\n\n' +
    'Reload the website: they now appear under People in "Married into the family", each with a page ' +
    'of their own showing their husband or wife and the children of the marriage.',
    ui.ButtonSet.OK);
}

/** Menu ▸ Add a spouse… — create a husband or wife directly. */
function addSpousePrompt() {
  const ui = ui_();
  ensureMarriageColumns_();
  const partnerAsk = ui.prompt('Add a spouse',
    'Person ID of the family member they married (for example P014).', ui.ButtonSet.OK_CANCEL);
  if (partnerAsk.getSelectedButton() !== ui.Button.OK) return;

  const t = table_('PEOPLE');
  const pid = partnerAsk.getResponseText().trim().toUpperCase();
  const partner = t.rows.filter(function (x) { return x.PersonID === pid; })[0];
  if (!partner) throw new Error('There is no person with the ID ' + pid + '.');
  if (marriedIn_(partner))
    throw new Error(pid + ' married into the family themselves. Give the Person ID of somebody in the line.');

  const nameAsk = ui.prompt('Add a spouse',
    'Their name, as the family knew them.', ui.ButtonSet.OK_CANCEL);
  if (nameAsk.getSelectedButton() !== ui.Button.OK) return;
  const nm = nameAsk.getResponseText().trim();
  if (!nm) { ui.alert('Please give them a name.'); return; }

  const id = makeSpouseRow_(nm, partner, oppositeGender_(partner.Gender),
    'Gender assumed from the marriage — please correct it if wrong.');

  const SECOND = secondParentField_(t.header);
  const kids = t.rows.filter(function (x) {
    return x.PersonID && x.PersonID !== partner.PersonID &&
           PARENT_FIELDS.some(function (f) { return String(x[f] || '').trim() === partner.PersonID; }) &&
           !String(x[SECOND] || '').trim();
  });
  var attached = 0;
  if (kids.length) {
    const ask = ui.alert('Are these their children too?',
      (partner.DisplayName || pid) + ' has ' + kids.length + ' child(ren) with no second parent recorded:\n\n' +
      kids.slice(0, 12).map(function (k) { return '  ' + (k.DisplayName || k.PersonID); }).join('\n') +
      (kids.length > 12 ? '\n  … and ' + (kids.length - 12) + ' more' : '') + '\n\n' +
      'Are they ' + nm + '\'s children as well? Say yes and ' + id + ' goes into ' + SECOND +
      ' on each, so they appear on her page too.\n\n' +
      'Say no if any of them were by someone else — you can then fill in ' + SECOND + ' by hand.',
      ui.ButtonSet.YES_NO);
    if (ask === ui.Button.YES) {
      const iSecond = t.header.indexOf(SECOND) + 1;
      kids.forEach(function (k) { t.sheet.getRange(k._row, iSecond).setValue(id); attached++; });
    }
  }

  ui.alert('Added',
    nm + ' has been added as ' + id + ', married to ' + (partner.DisplayName || pid) + '.\n' +
    'The marriage is recorded on both rows, and a Drive folder has been made for their photographs.\n\n' +
    (attached ? attached + ' child(ren) now name her as a parent, so they show on her page.\n\n' : '') +
    'She has a page of her own but no place in the line of descent.' +
    (attached || !kids.length ? '' :
      '\n\nTo show the children of this marriage on her page, put ' + id + ' into ' + SECOND +
      ' on each of those children\'s rows.'),
    ui.ButtonSet.OK);
}

/** Menu ▸ Check the record for problems. */
function validateRecord() {
  const peopleTable = table_('PEOPLE');
  const people = peopleTable.rows;
  const byId = {};
  const problems = [], warnings = [];

  // Without these three columns the site does not break — it quietly behaves as
  // though nobody in the family ever married, and any spouse name approved from
  // the inbox is dropped on the way in. Silence is the whole problem, so say so.
  const absent = ['SpouseID', 'Bloodline'].filter(function (h) {
    return peopleTable.header.indexOf(h) < 0;
  });
  const andOr = function (a) {
    return a.length < 2 ? a.join('') : a.slice(0, -1).join(', ') + ' or ' + a[a.length - 1];
  };
  if (absent.length)
    problems.push('The PEOPLE sheet has no ' + andOr(absent) + ' column' +
      (absent.length > 1 ? 's' : '') + ', so marriages cannot be recorded. ' +
      'Run Family Tree ▸ Add the marriage columns and ' +
      (absent.length > 1 ? 'they' : 'it') + ' will be added for you.');

  people.forEach(function (r) {
    if (!r.PersonID) { problems.push('A row in PEOPLE has no Person ID (row ' + r._row + ').'); return; }
    if (byId[r.PersonID]) problems.push('Person ID ' + r.PersonID + ' is used twice (rows ' + byId[r.PersonID]._row + ' and ' + r._row + ').');
    byId[r.PersonID] = r;
  });

  const yr = function (s) { const m = String(s || '').match(/\b(\d{4})\b/); return m ? +m[1] : null; };
  const marriedInIds = {};
  people.forEach(function (r) { if (marriedIn_(r)) marriedInIds[r.PersonID] = true; });
  people.forEach(function (r) {
    if (marriedIn_(r) && parentsOfRow_(r, byId).length)
      warnings.push(r.PersonID + ' is marked as married into the family but also has a parent recorded. ' +
                    'Decide which they are.');
    // A husband or wife who married in may perfectly well be somebody's mother
    // or father — that is what the second parent column is for. What is wrong is
    // a child whose ONLY recorded parent married in, because then nothing joins
    // them to the line of descent and they drop out of the tree entirely.
    const hasBlood = parentsOfRow_(r, byId).length > 0;
    PARENT_FIELDS.forEach(function (k) {
      if (r[k] && marriedInIds[r[k]] && !hasBlood)
        warnings.push(r.PersonID + '\'s only recorded parent is ' + r[k] + ', who married into the family. ' +
                      'Add the blood parent as well, or ' + r.PersonID + ' will not appear in the tree.');
      if (r[k] && !byId[r[k]]) problems.push(r.PersonID + ' has ' + k + ' = ' + r[k] + ', but no such person exists.');
      if (r[k] && r[k] === r.PersonID) problems.push(r.PersonID + ' is recorded as their own parent.');
    });
    // A mistyped SpouseID is simply ignored by the website — the marriage
    // vanishes without a word — so it has to be caught here.
    String(r.SpouseID || '').split(/[;,]/).forEach(function (x) {
      x = x.trim(); if (!x) return;
      if (x === r.PersonID) problems.push(r.PersonID + ' is recorded as their own spouse.');
      else if (!byId[x]) problems.push(r.PersonID + ' has SpouseID = ' + x + ', but no such person exists. ' +
        'Use the Spouse column for a wife or husband who has no record of their own.');
    });
    parentsOfRow_(r, byId).forEach(function (pid) {
      const c = yr(r.BirthDate), pb = yr(byId[pid].BirthDate);
      if (c && pb && c <= pb + 12)
        warnings.push(r.PersonID + ' (b. ' + c + ') is barely younger than their parent ' + pid + ' (b. ' + pb + ').');
    });
  });

  // Cycles: nobody may be their own ancestor.
  people.forEach(function (r) {
    var cur = r.PersonID, seen = {}, hops = 0;
    while (cur && byId[cur] && hops++ < 200) {
      if (seen[cur]) { problems.push('Circular parentage involving ' + r.PersonID + '.'); break; }
      seen[cur] = true;
      cur = parentsOfRow_(byId[cur], byId)[0];
    }
  });

  // Sub-tables pointing at people who do not exist.
  ['RELATIONSHIPS','PLACES','OCCUPATIONS','EDUCATION','EVENTS','PHOTOS','STORIES'].forEach(function (tab) {
    table_(tab).rows.forEach(function (r) {
      ['PersonID','Person1ID','Person2ID'].forEach(function (k) {
        if (r[k] && !byId[r[k]]) problems.push(tab + ' row ' + r._row + ' refers to ' + r[k] + ', who does not exist.');
      });
    });
  });

  // Anyone floating free of the main tree.
  const linked = {};
  people.forEach(function (r) {
    const ps = parentsOfRow_(r, byId);
    if (ps.length) linked[r.PersonID] = true;
    ps.forEach(function (p) { linked[p] = true; });
  });
  table_('RELATIONSHIPS').rows.forEach(function (r) {
    if (byId[r.Person1ID] && byId[r.Person2ID]) { linked[r.Person1ID] = true; linked[r.Person2ID] = true; }
  });
  // A marriage attaches a spouse to the family, because that is the only way
  // they are attached at all. It does NOT attach anyone in the line: being
  // married is exactly compatible with having been entered without a ParentID,
  // which is the mistake this check exists to catch.
  people.forEach(function (r) {
    if (marriedIn_(r) && (String(r.Spouse || '').trim() || String(r.SpouseID || '').trim()))
      linked[r.PersonID] = true;
    String(r.SpouseID || '').split(/[;,]/).forEach(function (x) {
      x = x.trim(); if (x && byId[x]) linked[x] = true;
    });
  });
  const loose = people.filter(function (r) { return r.PersonID && !linked[r.PersonID]; })
                      .map(function (r) { return r.PersonID + ' (' + (r.DisplayName || '?') + ')'; });
  if (loose.length) warnings.push('Not connected to anyone yet: ' + loose.join(', ') + '.');

  const head = problems.length
    ? problems.length + ' thing(s) need fixing'
    : (warnings.length ? 'No errors — but ' + warnings.length + ' thing(s) worth a look' : 'The record looks sound');
  const body =
    (problems.length ? 'MUST FIX\n' + problems.map(function (p) { return '• ' + p; }).join('\n') + '\n\n' : '') +
    (warnings.length ? 'WORTH CHECKING\n' + warnings.map(function (w) { return '• ' + w; }).join('\n') : '') ||
    'Every Person ID is unique, every parent link points at a real person, and nothing is orphaned.';
  ui_().alert(head, body.slice(0, 8000), ui_().ButtonSet.OK);
}

/** Menu ▸ Family statistics. */
function showStats() {
  const people = table_('PEOPLE').rows;
  const gens = people.map(function (r) { return parseInt(r.Generation, 10); })
                     .filter(function (n) { return !isNaN(n); });
  const branches = {};
  // Which line someone is on comes from the tree, not from a column.
  const byId = {}, kids = {};
  people.forEach(function (r) { byId[r.PersonID] = r; });
  people.forEach(function (r) {
    parentsOfRow_(r, byId).forEach(function (x) { (kids[x] = kids[x] || []).push(r.PersonID); });
  });
  const root = people.filter(function (r) {
    return inLine_(r.PersonID) && !parentsOfRow_(r, byId).length;
  })[0];
  (root ? (kids[root.PersonID] || []) : []).forEach(function (head) {
    var n = 0, q = [head];
    while (q.length) { const c = q.shift(); n++; (kids[c] || []).forEach(function (k) { q.push(k); }); }
    branches[(byId[head].DisplayName || head) + '\u2019s line'] = n;
  });
  const unnamed = people.filter(function (r) { return /name unknown/i.test(r.Status) || /^unknown/i.test(r.DisplayName); }).length;
  const married = people.filter(marriedIn_).length;
  const living = people.filter(function (r) { return /^yes$/i.test(r.Living); }).length;

  ui_().alert('The family so far',
    'People recorded: ' + people.length + '  (in the line: ' + (people.length - married) +
      ', married in: ' + married + ')\n' +
    'Generations: ' + (gens.length ? Math.max.apply(null, gens) : 0) + '\n' +
    'Living: ' + living + '\n' +
    'Names still to recover: ' + unnamed + '\n\n' +
    'Branches\n' + Object.keys(branches).sort().map(function (b) { return '  ' + b + ': ' + branches[b]; }).join('\n') + '\n\n' +
    'Photographs: ' + table_('PHOTOS').rows.length + '  (approved: ' +
      table_('PHOTOS').rows.filter(function (r) { return /^approved$/i.test(r.ApprovalStatus); }).length + ')\n' +
    'Stories: ' + table_('STORIES').rows.length + '\n' +
    'Places recorded: ' + table_('PLACES').rows.length + '\n' +
    'Occupations recorded: ' + table_('OCCUPATIONS').rows.length,
    ui_().ButtonSet.OK);
}

// ─── Approval queue ────────────────────────────────────────────────────────

function reviewQueue() {
  const t = inboxRows_();
  const waiting = t.rows.filter(function (r) { return /^pending$/i.test(r.Status); });
  const legacy = table_('PHOTOS').rows.filter(function (r) { return /^pending$/i.test(r.ApprovalStatus); }).length
               + table_('STORIES').rows.filter(function (r) { return /^pending$/i.test(r.ApprovalStatus); }).length;

  const warn = legacy
    ? '\n\n⚠ ' + legacy + ' older pending row(s) are still sitting in PHOTOS/STORIES, where the\n' +
      'website downloads them. Run Family Tree ▸ Move old pending rows into the inbox.'
    : '';

  if (!waiting.length) {
    ui_().alert('Approval queue', 'Nothing is waiting.' + warn, ui_().ButtonSet.OK);
    return;
  }
  t.sheet.activate();
  ui_().alert('Approval queue',
    waiting.length + ' submission(s) waiting:\n\n' +
    waiting.map(function (r) {
      return '  row ' + r._row + ' · ' + r.Kind + ' · ' + (r.Person || r.PersonID) +
             ' · from ' + (r.Uploader || '?') + (r.Title ? ' · "' + r.Title + '"' : '');
    }).join('\n') +
    '\n\nRead them on the INBOX sheet, select the rows you want, then use\n' +
    'Family Tree ▸ Approve the selected rows (or Reject).' + warn,
    ui_().ButtonSet.OK);
}

/**
 * Turn an approved person-suggestion into a real PEOPLE row, wired into the
 * tree according to the relationship the contributor described.
 *
 * The new person gets a fresh permanent Person ID. Nothing existing is
 * overwritten: if the relationship would clash with parentage already on
 * record, the conflict is written into Notes for the administrator to settle
 * rather than silently applied.
 */
function publishPerson_(r) {
  // A contributor may have named who this person married. If the column is not
  // there the value would be dropped without a word, so make sure it is.
  ensureMarriageColumns_();
  const t = table_('PEOPLE');
  const related = t.rows.filter(function (x) { return x.PersonID === r.PersonID; })[0];
  if (!related) throw new Error('The person this suggestion hangs off (' + r.PersonID + ') no longer exists.');
  // Checked again at approval time: the suggestion may predate the guard in
  // suggestPerson_, or the person may have been marked Bloodline = No since.
  if (marriedIn_(related))
    throw new Error((related.DisplayName || related.PersonID) + ' is marked as married into the family, so ' +
      'nobody can be attached through them. Reject this row, or change who it hangs off, and approve it again.');

  const byId = {};
  t.rows.forEach(function (x) { byId[x.PersonID] = x; });

  const PARENT = primaryParentField_(t.header);
  var rel = String(r.Relation || 'child of').toLowerCase();
  if (rel.indexOf('spouse') === 0 || rel.indexOf('partner') === 0) rel = 'married to';
  const gender = ['M', 'F'].indexOf(String(r.Gender || '').toUpperCase()) >= 0
    ? String(r.Gender).toUpperCase() : 'U';
  const relGen = parseInt(related.Generation, 10);
  const notes = [];
  const fields = {};
  var generation = '';

  if (rel.indexOf('child') === 0) {
      // A mother belongs in MotherID and a father in FatherID.
      fields[parentFieldFor_(t.header, related.Gender)] = related.PersonID;
    if (!isNaN(relGen)) generation = relGen + 1;
      if (['M', 'F'].indexOf(String(related.Gender || '').toUpperCase()) < 0)
        notes.push('The gender of ' + related.PersonID + ' is not recorded, so they were entered as '
                   + 'the father. Move to ' + secondParentField_(t.header) + ' if that is wrong.');

  } else if (rel.indexOf('sibling') === 0) {
    // Take the same parents as the sibling — including a mother who married in.
    const ps = allParentsOfRow_(related, byId);
    if (ps.length) fields[PARENT] = ps[0];
    if (ps.length > 1) fields[secondParentField_(t.header)] = ps[1];
    if (!ps.length)
      notes.push('No parent is recorded for ' + related.PersonID + ', so this sibling is not linked to the tree yet. ' +
                 'Give them both the same ParentID.');
    if (!isNaN(relGen)) generation = relGen;

  } else if (rel.indexOf('parent') === 0) {
    if (!isNaN(relGen)) generation = Math.max(1, relGen - 1);

  } else if (rel.indexOf('married') === 0) {
    // A husband or wife joins the family without joining the line. They get an
    // S id, no generation and no parent link; the marriage is the whole of
    // their attachment, and it is written onto both rows below.
    fields['SpouseID'] = related.PersonID;
    fields['Bloodline'] = 'No';
    generation = '';
  }

  const wed = rel.indexOf('married') === 0;
  const id = nextId_('PEOPLE', 'PersonID', wed ? 'S' : 'P', 3);
  const row = {
    PersonID: id,
    DisplayName: r.Name || 'Unknown',
    Gender: gender,
    Generation: generation,
    BirthDate: r.BirthDate || '',
    BirthPlace: r.BirthPlace || '',
    DeathDate: r.DeathDate || '',
    DeathPlace: r.DeathPlace || '',
    // Anyone not clearly recorded as dead is treated as living, so the privacy
    // screen covers them from the moment they appear.
    Living: /^no$/i.test(String(r.Living || '')) ? 'No' : 'Yes',
    Privacy: 'Public',
    Status: 'Reported by family',
    Notes: ['Suggested by ' + (r.Uploader || 'a family member') +
            (r.Contact ? ' (' + r.Contact + ')' : '') + ', ' + rel + ' ' + related.PersonID + '.',
            r.Body ? 'They wrote: ' + r.Body : ''].concat(notes)
           .filter(function (x) { return x; }).join(' ')
  };
  Object.keys(fields).forEach(function (k) { row[k] = fields[k]; });
  appendRow_('PEOPLE', row);

  // A marriage belongs on both rows, so the person already on record shows as
  // married when you look at their row rather than only on the new one.
  if (wed) {
    linkSpouseBack_(related.PersonID, id);
    try { personFolder_(id); } catch (err) {}
  }

  // If the contributor also said who this person married, that name becomes a
  // person too rather than a note in a column — a page of their own, and the
  // marriage recorded between the two.
  if (!wed && String(r.Spouse || '').trim()) {
    const fresh = table_('PEOPLE').rows.filter(function (x) { return x.PersonID === id; })[0];
    String(r.Spouse).split(/;/).forEach(function (nm) {
      nm = nm.trim(); if (!nm || !fresh) return;
      makeSpouseRow_(nm, fresh, oppositeGender_(gender),
        'Named by ' + (r.Uploader || 'a family member') + ' as the husband or wife of ' + id +
        '. Gender assumed from the marriage — please correct it if wrong.');
    });
  }

  // "Parent of" points the other way round: the existing person gains a parent.
  // Use the main column if it is free, otherwise the second one. Never overwrite.
  if (rel.indexOf('parent') === 0) {
    const t2 = table_('PEOPLE');
    const target = t2.rows.filter(function (x) { return x.PersonID === related.PersonID; })[0];
    const SECOND = secondParentField_(t2.header);
    const iNotes = t2.header.indexOf('Notes') + 1;
    if (target) {
      // A mother belongs in MotherID and a father in FatherID. When the gender is
      // known, that is the only column they may go in — a father whose slot is
      // already taken is a conflict to settle, not someone to file as a mother.
      // Only an unrecorded gender may fall back to whichever column is free.
      const want = parentFieldFor_(t2.header, gender);
      const other = want === PARENT ? SECOND : PARENT;
      const known = ['M', 'F'].indexOf(gender) >= 0;
      const free = !String(target[want] || '').trim() ? want
                 : (!known && !String(target[other] || '').trim()) ? other : '';
      if (free) {
        t2.sheet.getRange(target._row, t2.header.indexOf(free) + 1).setValue(id);
      } else if (iNotes) {
        t2.sheet.getRange(target._row, iNotes).setValue(
          (target.Notes ? target.Notes + ' ' : '') +
          'A family member suggested ' + id + ' as a parent, but two are already recorded. Please resolve.');
      }
    }
  }

  return id;
}

/** Copy one approved inbox row into the public sheets. */
function publishInboxRow_(r) {
  if (/^person$/i.test(r.Kind))     return publishPerson_(r);
  if (/^correction$/i.test(r.Kind)) return '';   // applied by hand; approving just files it
  if (/^photo$/i.test(r.Kind)) {
    const id = nextId_('PHOTOS', 'PhotoID', 'F', 3);
    appendRow_('PHOTOS', {
      PhotoID: id, PersonID: r.PersonID, DriveFileID: r.DriveFileID,
      Caption: r.Title || '', PhotoDate: r.PhotoDate || '', Place: r.Place || '',
      PeopleShown: r.PersonID, Uploader: r.Uploader || '', UploadedAt: r.When || '',
      ApprovalStatus: 'Approved', IsProfile: 'No', Notes: 'Approved from the inbox.'
    });
    return id;
  }
  const id = nextId_('STORIES', 'StoryID', 'S', 3);
  appendRow_('STORIES', {
    StoryID: id, PersonID: r.PersonID, Title: r.Title || 'Untitled', Story: r.Body || '',
    ToldBy: r.Uploader || '', RecordedDate: String(r.When || '').slice(0, 10),
    Category: 'Memory', ApprovalStatus: 'Approved', Notes: 'Approved from the inbox.'
  });
  return id;
}

function actOnSelection_(approve) {
  const sh = book_().getActiveSheet();
  const rng = sh.getActiveRange();

  if (sh.getName() === INBOX) {
    const t = inboxRows_();
    const iStatus = t.header.indexOf('Status') + 1;
    const iPub = t.header.indexOf('Published as') + 1;
    var done = 0, skipped = 0, corrections = 0;
    const addedPeople = [], refused = [];

    for (var r = rng.getRow(); r < rng.getRow() + rng.getNumRows(); r++) {
      if (r === 1) continue;
      const hit = t.rows.filter(function (x) { return x._row === r; })[0];
      if (!hit) continue;
      if (!/^pending$/i.test(hit.Status)) { skipped++; continue; }

      if (approve) {
        // One row that cannot be published must not abandon the rest of the
        // selection half-approved, so each is dealt with on its own.
        var id;
        try { id = publishInboxRow_(hit); }
        catch (err) { refused.push('row ' + r + ': ' + String(err && err.message || err)); continue; }
        sh.getRange(r, iStatus).setValue('Approved');
        sh.getRange(r, iPub).setValue(id || 'applied by hand');
        if (/^person$/i.test(hit.Kind)) addedPeople.push((hit.Name || id) + ' → ' + id);
        if (/^correction$/i.test(hit.Kind)) corrections++;
      } else {
        sh.getRange(r, iStatus).setValue('Rejected');
        if (hit.DriveFileID) {
          try { DriveApp.getFileById(hit.DriveFileID).setTrashed(true); } catch (err) {}
        }
      }
      done++;
    }
    ui_().alert(approve ? 'Approved' : 'Rejected',
      done + ' submission(s) ' + (approve ? 'published to the site' : 'rejected') + '.' +
      (skipped ? '\n' + skipped + ' row(s) skipped — they had already been dealt with.' : '') +
      (refused.length
        ? '\n\n' + refused.length + ' row(s) could not be published and are still Pending:\n  ' +
          refused.join('\n  ')
        : '') +
      (addedPeople.length
        ? '\n\nAdded to PEOPLE:\n  ' + addedPeople.join('\n  ') +
          '\n\nEach new person is marked Living = Yes and Status = Reported by family, and anything\n' +
          'the contributor wrote is in their Notes. Read those notes — they may flag a conflict.\n' +
          'Then run Family Tree ▸ Recalculate generations, and Name the family branches if you use them.'
        : '') +
      (corrections
        ? '\n\n' + corrections + ' correction(s) filed. Corrections are never applied automatically —\n' +
          'read the Body column and make the edit yourself.'
        : '') +
      (approve && done ? '\n\nReload the website to see the changes.' : ''),
      ui_().ButtonSet.OK);
    return;
  }

  // Rows you entered by hand on PHOTOS or STORIES.
  const nm = sh.getName();
  if (nm !== 'PHOTOS' && nm !== 'STORIES') {
    ui_().alert('Select rows on the INBOX sheet first.',
      'Family Tree ▸ Review the approval queue will take you there.', ui_().ButtonSet.OK);
    return;
  }
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  const col = header.indexOf('ApprovalStatus') + 1;
  if (!col) { ui_().alert('That sheet has no ApprovalStatus column.'); return; }
  var n = 0;
  for (var q = rng.getRow(); q < rng.getRow() + rng.getNumRows(); q++) {
    if (q === 1) continue;
    sh.getRange(q, col).setValue(approve ? 'Approved' : 'Rejected');
    n++;
  }
  ui_().alert(approve ? 'Approved' : 'Rejected', n + ' row(s) updated.', ui_().ButtonSet.OK);
}
function approveSelection() { actOnSelection_(true); }
function rejectSelection()  { actOnSelection_(false); }

/**
 * Menu ▸ Move old pending rows into the inbox.
 * One-off tidy-up for anything submitted before the inbox existed. Pending rows
 * in PHOTOS/STORIES are visible to every website visitor; this gets them out.
 */
function migratePendingToInbox() {
  inbox_();
  var moved = 0;

  [['PHOTOS', 'Photo'], ['STORIES', 'Story']].forEach(function (spec) {
    const t = table_(spec[0]);
    const pending = t.rows.filter(function (r) { return /^pending$/i.test(r.ApprovalStatus); });
    // Delete from the bottom up so earlier row numbers stay valid.
    pending.sort(function (a, b) { return b._row - a._row; }).forEach(function (r) {
      appendInbox_({
        'When': r.UploadedAt || r.RecordedDate || new Date(), 'Kind': spec[1],
        'PersonID': r.PersonID, 'Person': personName_(r.PersonID),
        'Title': r.Caption || r.Title || '', 'Body': r.Story || '', 'DriveFileID': r.DriveFileID || '',
        'PhotoDate': r.PhotoDate || '', 'Place': r.Place || '',
        'Uploader': r.Uploader || r.ToldBy || '', 'Status': 'Pending'
      });
      t.sheet.deleteRow(r._row);
      moved++;
    });
  });

  ui_().alert(moved ? 'Moved to the inbox' : 'Nothing to move',
    moved
      ? moved + ' pending row(s) moved out of the public sheets and into INBOX.\n\n' +
        'They are no longer downloaded by visitors. Review them there as usual.'
      : 'No pending rows were left in PHOTOS or STORIES. Nothing needed moving.',
    ui_().ButtonSet.OK);
}

/** Menu ▸ Make the selected photo the profile photo. */
function makeProfilePhoto() {
  const sh = book_().getActiveSheet();
  if (sh.getName() !== 'PHOTOS') { ui_().alert('Select a row on the PHOTOS sheet first.'); return; }
  const row = sh.getActiveRange().getRow();
  if (row === 1) { ui_().alert('Select a photo row, not the header.'); return; }

  const t = table_('PHOTOS');
  const hit = t.rows.filter(function (r) { return r._row === row; })[0];
  if (!hit) { ui_().alert('That row is empty.'); return; }
  if (!hit.PersonID) { ui_().alert('That photo has no PersonID.'); return; }

  const iProf = t.header.indexOf('IsProfile') + 1;
  const iStat = t.header.indexOf('ApprovalStatus') + 1;
  t.rows.forEach(function (r) {
    if (r.PersonID === hit.PersonID) sh.getRange(r._row, iProf).setValue(r._row === row ? 'Yes' : 'No');
  });
  sh.getRange(row, iStat).setValue('Approved');

  // Clear any hand-entered override so the flagged photo wins.
  const pt = table_('PEOPLE');
  const iPhoto = pt.header.indexOf('ProfilePhoto') + 1;
  pt.rows.forEach(function (r) {
    if (r.PersonID === hit.PersonID) pt.sheet.getRange(r._row, iPhoto).setValue('');
  });

  ui_().alert('Done',
    'The silhouette for ' + personName_(hit.PersonID) + ' will be replaced by this photograph.\n\n' +
    'Their Person ID has not changed. Their previous photographs are still on record.',
    ui_().ButtonSet.OK);
}

// ─── Backups ───────────────────────────────────────────────────────────────

/** Menu ▸ Back up now. Also runs nightly once the trigger is installed. */
function backupNow() {
  const folder = childFolder_(driveRoot_(), 'Backups');
  const id = book_().getId();
  const url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx';
  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob();

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  blob.setName('FamilyTree_' + stamp + '.xlsx');
  folder.createFile(blob);

  // Keep the most recent KEEP_BACKUPS files.
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(KEEP_BACKUPS).forEach(function (f) { f.setTrashed(true); });

  try {
    ui_().alert('Backed up', 'Saved to ' + DRIVE_ROOT_NAME + ' ▸ Backups ▸ ' + blob.getName(), ui_().ButtonSet.OK);
  } catch (err) { /* running from a trigger — no UI available */ }
}

/**
 * Menu ▸ Nightly backups: on / off.
 *
 * Off by default. A dated copy every night fills the Drive folder with files
 * that are almost all identical, so the useful pattern is to take a backup when
 * you actually want one — before a big edit, or to work on the sheet offline.
 */
/**
 * Menu ▸ Update the sheet to the latest layout.
 *
 * Brings a live spreadsheet in step with what the website reads:
 *   • the parent columns become FatherID and MotherID — everyone has both
 *   • Bloodline retires, because a P or S number already says it
 *   • Branch retires, because which line someone is on is worked out from the tree
 *
 * A column is only ever removed when it is empty. Nothing that holds data is
 * deleted, no value moves cell, and no ID changes. Safe to run twice.
 */
function updateSheetLayout() {
  const ui = ui_();
  const sh = sheet_('PEOPLE');
  var header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]
                 .map(function (h) { return String(h).trim(); });
  const done = [], warn = [];

  // 1. gendered parent columns: everyone has a father and a mother
  [['ParentID', 'FatherID'], ['Parent2ID', 'MotherID']].forEach(function (pair) {
    const i = header.indexOf(pair[0]);
    if (i >= 0 && header.indexOf(pair[1]) < 0) {
      sh.getRange(1, i + 1).setValue(pair[1]);
      header[i] = pair[1];
      done.push(pair[0] + ' renamed to ' + pair[1]);
    }
  });

  // 2. columns the website no longer reads, removed only when empty
  ['Bloodline', 'Branch'].forEach(function (name) {
    const i = header.indexOf(name);
    if (i < 0) return;
    const last = Math.max(sh.getLastRow() - 1, 0);
    const used = last
      ? sh.getRange(2, i + 1, last, 1).getValues()
          .filter(function (v) { return String(v[0]).trim() !== ''; }).length
      : 0;
    if (!used) {
      sh.deleteColumn(i + 1); header.splice(i, 1);
      done.push(name + ' removed — it was empty and is no longer read');
    } else if (name === 'Bloodline') {
      warn.push(used + ' row(s) still carry a Bloodline value. It is now the P or S in the ' +
                'ID that decides, so those values are ignored. Clear the column and run this ' +
                'again to remove it.');
    } else {
      warn.push(used + ' row(s) still carry a Branch value. Which line someone is on is worked ' +
                'out from the tree now, so it is ignored. Clear the column and run this again.');
    }
  });

  // 3. anyone married in should carry an S number, or they will show in the tree
  const strays = table_('PEOPLE').rows.filter(function (r) {
    return String(r.SpouseID || '').trim() && !inLine_(r.PersonID) === false &&
           /^no$/i.test(String(r.Bloodline || '').trim());
  });
  if (strays.length)
    warn.push(strays.length + ' person(s) are marked Bloodline = No but still have a P number, ' +
              'so they would appear in the tree: ' +
              strays.slice(0, 8).map(function (r) { return r.PersonID; }).join(', ') +
              '. Give them S numbers.');

  if (!done.length && !warn.length) {
    ui.alert('Already up to date',
      'The sheet already matches what the website reads. Nothing needed changing.\n\n' +
      'If the site still looks unchanged, it is a cached page, not the sheet — open ' +
      'your site at /#/admin and check the build number shown there.', ui.ButtonSet.OK);
    return;
  }

  ui.alert(done.length ? 'Sheet updated' : 'Needs your attention',
    (done.length ? done.map(function (d) { return '• ' + d; }).join('\n') + '\n\n' : '') +
    (warn.length ? 'WORTH CHECKING\n' + warn.map(function (w) { return '• ' + w; }).join('\n') + '\n\n' : '') +
    'No value moved and no ID changed.\n\n' +
    'A P number means descended from the family; an S number means married into it. ' +
    'That letter is the only thing deciding who appears in the tree.\n\n' +
    'Reload the website to see the change.', ui.ButtonSet.OK);
}

function toggleNightlyBackup() {
  const ui = ui_();
  const existing = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'backupNow';
  });

  if (existing.length) {
    existing.forEach(function (t) { ScriptApp.deleteTrigger(t); });
    ui.alert('Nightly backups are off',
      'No more backups will be made on their own. Use Family Tree ▸ Back up now whenever ' +
      'you want a copy — before a big edit, or to work on the sheet offline.\n\n' +
      'Backups already in Drive are left alone; delete any you do not want.',
      ui.ButtonSet.OK);
    return;
  }

  const go = ui.alert('Turn nightly backups on?',
    'A dated copy of this spreadsheet will be saved every night, and the most recent ' +
    KEEP_BACKUPS + ' kept.\n\nMost of them will be identical to each other. Run this again at ' +
    'any time to switch it back off.', ui.ButtonSet.YES_NO);
  if (go !== ui.Button.YES) return;

  ScriptApp.newTrigger('backupNow').timeBased().atHour(3).everyDays(1).create();
  ui.alert('Nightly backups are on', 'One dated copy a night, keeping the most recent ' +
    KEEP_BACKUPS + '.', ui.ButtonSet.OK);
}

function installNightlyBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupNow').timeBased().atHour(3).everyDays(1).create();
  ui_().alert('Nightly backups are on',
    'A dated copy of this spreadsheet will be saved to ' + DRIVE_ROOT_NAME + ' ▸ Backups every night.\n' +
    'The most recent ' + KEEP_BACKUPS + ' are kept.', ui_().ButtonSet.OK);
}

// ─── Setup helper ──────────────────────────────────────────────────────────

function showSetupDetails() {
  const id = book_().getId();
  writeSetting_('sheet_id', id);
  var scriptUrl = '';
  try { scriptUrl = ScriptApp.getService().getUrl() || ''; } catch (err) {}

  ui_().alert('Details for index.html',
    'SHEET_ID:\n' + id + '\n\n' +
    'APPS_SCRIPT_URL:\n' + (scriptUrl || '(deploy the web app first: Deploy ▸ New deployment ▸ Web app)') + '\n\n' +
    'Paste these into the CONFIG block at the top of index.html, commit, and the\n' +
    'website will start reading live family data.\n\n' +
    'Remember: this spreadsheet must be shared as "Anyone with the link ▸ Viewer"\n' +
    'for the website to read it.',
    ui_().ButtonSet.OK);
}

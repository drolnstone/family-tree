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
  'Relation','Name','Gender','BirthDate','BirthPlace','DeathDate','DeathPlace','Living','Contact'
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

  inbox_().appendRow([
    new Date(), 'Photo', b.personId, personName_(b.personId),
    b.caption || '', '', file.getId(),
    b.photoDate || '', b.place || '', b.uploader || 'Anonymous', 'Pending', ''
  ]);
  notifyAdmin_('New photograph awaiting approval',
    (b.uploader || 'Someone') + ' submitted a photograph for ' + personName_(b.personId) + '.');
  return { ok: true };
}

function addStory_(b) {
  if (!b.personId || !b.story) return { ok: false, error: 'Missing person or story.' };
  if (!table_('PEOPLE').rows.some(function (r) { return r.PersonID === b.personId; }))
    return { ok: false, error: 'Unknown person.' };

  inbox_().appendRow([
    new Date(), 'Story', b.personId, personName_(b.personId),
    b.title || 'Untitled', String(b.story).slice(0, 40000), '',
    '', '', b.toldBy || '', 'Pending', ''
  ]);
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

  const rel = ['child of', 'sibling of', 'spouse of', 'parent of'].indexOf(clip_(b.relation, 20)) >= 0
    ? clip_(b.relation, 20) : 'child of';
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

function notifyAdmin_(subject, body) {
  const to = settings_().contact_email || Session.getEffectiveUser().getEmail();
  if (!to) return;
  try {
    MailApp.sendEmail(to, '[Family Tree] ' + subject,
      body + '\n\nOpen the spreadsheet and use Family Tree ▸ Review the approval queue.');
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

function onOpen() {
  ui_().createMenu('Family Tree')
    .addItem('Add a person…', 'addPersonPrompt')
    .addItem('Give IDs to any new rows', 'assignMissingIds')
    .addItem('Recalculate generations', 'recalcGenerations')
    .addItem('Name the family branches…', 'assignBranches')
    .addSeparator()
    .addItem('Check the record for problems', 'validateRecord')
    .addItem('Family statistics', 'showStats')
    .addSeparator()
    .addItem('Review the approval queue', 'reviewQueue')
    .addItem('Move old pending rows into the inbox', 'migratePendingToInbox')
    .addItem('Approve the selected rows', 'approveSelection')
    .addItem('Reject the selected rows', 'rejectSelection')
    .addItem('Make the selected photo the profile photo', 'makeProfilePhoto')
    .addSeparator()
    .addItem('Create the Google Drive folders', 'createDriveFolders')
    .addItem('Back up now', 'backupNow')
    .addItem('Turn on nightly backups', 'installNightlyBackup')
    .addSeparator()
    .addItem('Setup details for the website', 'showSetupDetails')
    .addToUi();
}

/** Menu ▸ Add a person — fills in the ID, generation and branch for you. */
function addPersonPrompt() {
  const ui = ui_();
  const nameAsk = ui.prompt('Add a person', 'What is their name? (Leave blank for an unknown name.)', ui.ButtonSet.OK_CANCEL);
  if (nameAsk.getSelectedButton() !== ui.Button.OK) return;

  const fatherAsk = ui.prompt('Add a person', "Father's Person ID (blank if unknown):", ui.ButtonSet.OK_CANCEL);
  if (fatherAsk.getSelectedButton() !== ui.Button.OK) return;
  const motherAsk = ui.prompt('Add a person', "Mother's Person ID (blank if unknown):", ui.ButtonSet.OK_CANCEL);
  if (motherAsk.getSelectedButton() !== ui.Button.OK) return;

  const father = fatherAsk.getResponseText().trim().toUpperCase();
  const mother = motherAsk.getResponseText().trim().toUpperCase();
  const people = table_('PEOPLE').rows;
  const byId = {};
  people.forEach(function (r) { byId[r.PersonID] = r; });

  [father, mother].forEach(function (p) {
    if (p && !byId[p]) throw new Error('There is no person with the ID ' + p + '.');
  });

  const parent = byId[father] || byId[mother];
  const id = nextId_('PEOPLE', 'PersonID', 'P', 3);
  appendRow_('PEOPLE', {
    PersonID: id,
    DisplayName: nameAsk.getResponseText().trim() || 'Unknown',
    Gender: 'U',
    FatherID: father,
    MotherID: mother,
    Branch: parent ? (parent.Branch || '') : '',
    Generation: parent && parent.Generation ? (parseInt(parent.Generation, 10) + 1) : '',
    Living: 'Unknown',
    Privacy: 'Public',
    Status: 'Reported by family'
  });
  ui.alert('Added', id + ' has been added. Open the PEOPLE sheet to fill in the rest of what you know.', ui.ButtonSet.OK);
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

  const byId = {}, kids = {};
  t.rows.forEach(function (r) { byId[r.PersonID] = r; });
  t.rows.forEach(function (r) {
    [r.FatherID, r.MotherID].forEach(function (p) {
      if (p && byId[p]) (kids[p] = kids[p] || []).push(r.PersonID);
    });
  });

  const gen = {};
  t.rows.forEach(function (r) {
    const hasParent = (r.FatherID && byId[r.FatherID]) || (r.MotherID && byId[r.MotherID]);
    if (!hasParent) gen[r.PersonID] = 1;
  });

  // Repeatedly push generations downward. Bounded so a cycle cannot hang it.
  for (var pass = 0; pass < t.rows.length + 2; pass++) {
    var changed = false;
    t.rows.forEach(function (r) {
      const parents = [r.FatherID, r.MotherID].filter(function (p) { return p && gen[p]; });
      if (!parents.length) return;
      const want = Math.max.apply(null, parents.map(function (p) { return gen[p]; })) + 1;
      if (gen[r.PersonID] !== want) { gen[r.PersonID] = want; changed = true; }
    });
    if (!changed) break;
  }

  var n = 0;
  t.rows.forEach(function (r) {
    const v = gen[r.PersonID] || '';
    if (String(r.Generation) !== String(v)) { t.sheet.getRange(r._row, col).setValue(v); n++; }
  });
  ui_().alert('Generations updated', n + ' row(s) changed.', ui_().ButtonSet.OK);
}

/**
 * Menu ▸ Name the family branches.
 *
 * Fills in the Branch column for everybody at once. You choose which generation
 * heads the branches; each person in that generation gives their name to a
 * branch, and every one of their descendants inherits the label. People above
 * that generation keep whatever label they already have. Nobody's Person ID,
 * parentage or any other column is touched.
 */
function assignBranches() {
  const ui = ui_();
  const t = table_('PEOPLE');
  const col = t.header.indexOf('Branch') + 1;
  if (!col) throw new Error('The PEOPLE sheet has no Branch column.');

  const byId = {}, kids = {};
  t.rows.forEach(function (r) { if (r.PersonID) byId[r.PersonID] = r; });
  t.rows.forEach(function (r) {
    [r.FatherID, r.MotherID].forEach(function (p) {
      if (p && byId[p]) (kids[p] = kids[p] || []).push(r.PersonID);
    });
  });

  // Work out generations from the data rather than trusting the column.
  const gen = {};
  t.rows.forEach(function (r) {
    if (!((r.FatherID && byId[r.FatherID]) || (r.MotherID && byId[r.MotherID]))) gen[r.PersonID] = 1;
  });
  for (var pass = 0; pass < t.rows.length + 2; pass++) {
    var changed = false;
    t.rows.forEach(function (r) {
      const ps = [r.FatherID, r.MotherID].filter(function (p) { return p && gen[p]; });
      if (!ps.length) return;
      const want = Math.max.apply(null, ps.map(function (p) { return gen[p]; })) + 1;
      if (gen[r.PersonID] !== want) { gen[r.PersonID] = want; changed = true; }
    });
    if (!changed) break;
  }

  const deepest = Math.max.apply(null, t.rows.map(function (r) { return gen[r.PersonID] || 1; }));
  const ask = ui.prompt('Name the family branches',
    'Each person in the generation you choose gives their name to a branch, and all of their\n' +
    'descendants get that label.\n\n' +
    'Generation 2 is usual — the children of the earliest known ancestor.\n' +
    'Choose 3 for finer branches. Your tree currently runs to generation ' + deepest + '.\n\n' +
    'Which generation heads the branches?', ui.ButtonSet.OK_CANCEL);
  if (ask.getSelectedButton() !== ui.Button.OK) return;

  const level = parseInt(ask.getResponseText().trim() || '2', 10);
  if (isNaN(level) || level < 2) { ui.alert('Please give a generation of 2 or more.'); return; }

  const heads = t.rows.filter(function (r) { return gen[r.PersonID] === level; });
  if (!heads.length) {
    ui.alert('Nobody is in generation ' + level + '.',
      'Run Family Tree ▸ Recalculate generations first, then try again.', ui.ButtonSet.OK);
    return;
  }

  // Give every head and its descendants a label. First claim wins, so a person
  // who descends from two heads keeps the earlier one rather than flickering.
  const label = {};
  heads.forEach(function (h) {
    const nm = (h.DisplayName || h.FullName || h.PersonID).trim();
    const text = /branch$/i.test(nm) ? nm : nm + ' Branch';
    const queue = [h.PersonID];
    while (queue.length) {
      const cur = queue.shift();
      if (label[cur]) continue;
      label[cur] = text;
      (kids[cur] || []).forEach(function (k) { if (!label[k]) queue.push(k); });
    }
  });

  // Married-in spouses have no parents on record, so take their partner's branch.
  table_('RELATIONSHIPS').rows.forEach(function (r) {
    if (!/spouse|partner|married/i.test(r.Type || '')) return;
    [[r.Person1ID, r.Person2ID], [r.Person2ID, r.Person1ID]].forEach(function (pair) {
      if (byId[pair[0]] && !label[pair[0]] && label[pair[1]]) label[pair[0]] = label[pair[1]];
    });
  });

  var written = 0, kept = 0, blank = 0;
  t.rows.forEach(function (r) {
    if (!r.PersonID) return;
    if (label[r.PersonID]) {
      if (String(r.Branch) !== label[r.PersonID]) {
        t.sheet.getRange(r._row, col).setValue(label[r.PersonID]);
        written++;
      }
    } else if (gen[r.PersonID] && gen[r.PersonID] < level) {
      if (!String(r.Branch).trim()) {
        t.sheet.getRange(r._row, col).setValue('Founding generation');
        written++;
      } else kept++;
    } else blank++;
  });

  const names = heads.map(function (h) {
    const nm = (h.DisplayName || h.PersonID).trim();
    return '  ' + (/branch$/i.test(nm) ? nm : nm + ' Branch');
  });
  ui.alert('Branches named',
    names.length + ' branch(es) created from generation ' + level + ':\n' + names.join('\n') + '\n\n' +
    written + ' row(s) updated' + (kept ? ', ' + kept + ' existing label(s) left alone' : '') +
    (blank ? ', ' + blank + ' person(s) not connected to any branch yet' : '') + '.\n\n' +
    'Reload the website to see them.', ui.ButtonSet.OK);
}

/** Menu ▸ Check the record for problems. */
function validateRecord() {
  const people = table_('PEOPLE').rows;
  const byId = {};
  const problems = [], warnings = [];

  people.forEach(function (r) {
    if (!r.PersonID) { problems.push('A row in PEOPLE has no Person ID (row ' + r._row + ').'); return; }
    if (byId[r.PersonID]) problems.push('Person ID ' + r.PersonID + ' is used twice (rows ' + byId[r.PersonID]._row + ' and ' + r._row + ').');
    byId[r.PersonID] = r;
  });

  people.forEach(function (r) {
    ['FatherID', 'MotherID'].forEach(function (k) {
      if (r[k] && !byId[r[k]]) problems.push(r.PersonID + ' has ' + k + ' = ' + r[k] + ', but no such person exists.');
    });
    if (r.FatherID && byId[r.FatherID] && byId[r.FatherID].Gender === 'F')
      warnings.push(r.PersonID + "'s father " + r.FatherID + ' is recorded as female.');
    if (r.MotherID && byId[r.MotherID] && byId[r.MotherID].Gender === 'M')
      warnings.push(r.PersonID + "'s mother " + r.MotherID + ' is recorded as male.');

    const yr = function (s) { const m = String(s || '').match(/\b(\d{4})\b/); return m ? +m[1] : null; };
    [['FatherID','father'],['MotherID','mother']].forEach(function (pair) {
      const p = byId[r[pair[0]]];
      if (!p) return;
      const c = yr(r.BirthDate), pb = yr(p.BirthDate);
      if (c && pb && c <= pb + 12)
        warnings.push(r.PersonID + ' (b. ' + c + ') is barely younger than their ' + pair[1] + ' ' + p.PersonID + ' (b. ' + pb + ').');
    });
  });

  // Cycles: nobody may be their own ancestor.
  people.forEach(function (r) {
    var cur = r.PersonID, seen = {}, hops = 0;
    while (cur && byId[cur] && hops++ < 200) {
      if (seen[cur]) { problems.push('Circular parentage involving ' + r.PersonID + '.'); break; }
      seen[cur] = true;
      cur = byId[cur].FatherID || byId[cur].MotherID;
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
    if ((r.FatherID && byId[r.FatherID]) || (r.MotherID && byId[r.MotherID])) linked[r.PersonID] = true;
    [r.FatherID, r.MotherID].forEach(function (p) { if (p && byId[p]) linked[p] = true; });
  });
  table_('RELATIONSHIPS').rows.forEach(function (r) {
    if (byId[r.Person1ID] && byId[r.Person2ID]) { linked[r.Person1ID] = true; linked[r.Person2ID] = true; }
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
  people.forEach(function (r) { const b = r.Branch || 'Unassigned'; branches[b] = (branches[b] || 0) + 1; });
  const unnamed = people.filter(function (r) { return /name unknown/i.test(r.Status) || /^unknown/i.test(r.DisplayName); }).length;
  const living = people.filter(function (r) { return /^yes$/i.test(r.Living); }).length;

  ui_().alert('The family so far',
    'People recorded: ' + people.length + '\n' +
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
  const t = table_('PEOPLE');
  const related = t.rows.filter(function (x) { return x.PersonID === r.PersonID; })[0];
  if (!related) throw new Error('The person this suggestion hangs off (' + r.PersonID + ') no longer exists.');

  const rel = String(r.Relation || 'child of').toLowerCase();
  const gender = ['M', 'F'].indexOf(String(r.Gender || '').toUpperCase()) >= 0
    ? String(r.Gender).toUpperCase() : 'U';
  const relGen = parseInt(related.Generation, 10);
  const notes = [];
  var father = '', mother = '', generation = '';

  if (rel.indexOf('child') === 0) {
    if (related.Gender === 'F') mother = related.PersonID;
    else father = related.PersonID;
    if (related.Gender !== 'M' && related.Gender !== 'F')
      notes.push('Parent ' + related.PersonID + "'s gender is not recorded, so they were put in FatherID. " +
                 'Move to MotherID if that is wrong.');
    if (!isNaN(relGen)) generation = relGen + 1;

  } else if (rel.indexOf('sibling') === 0) {
    father = related.FatherID || '';
    mother = related.MotherID || '';
    if (!father && !mother)
      notes.push('No parents are recorded for ' + related.PersonID + ', so this sibling is not linked to the tree yet. ' +
                 'Add a FatherID or MotherID to both of them.');
    if (!isNaN(relGen)) generation = relGen;

  } else if (rel.indexOf('spouse') === 0 || rel.indexOf('partner') === 0) {
    if (!isNaN(relGen)) generation = relGen;

  } else if (rel.indexOf('parent') === 0) {
    if (!isNaN(relGen)) generation = Math.max(1, relGen - 1);
  }

  const id = nextId_('PEOPLE', 'PersonID', 'P', 3);
  appendRow_('PEOPLE', {
    PersonID: id,
    DisplayName: r.Name || 'Unknown',
    Gender: gender,
    FatherID: father,
    MotherID: mother,
    Branch: related.Branch || '',
    Generation: generation,
    BirthDate: r.BirthDate || '',
    BirthPlace: r.BirthPlace || '',
    DeathDate: r.DeathDate || '',
    DeathPlace: r.DeathPlace || '',
    // Anyone whose status is not clearly "no" is treated as living, so the
    // privacy screen covers them from the moment they appear.
    Living: /^no$/i.test(String(r.Living || '')) ? 'No' : (/^yes$/i.test(String(r.Living || '')) ? 'Yes' : 'Yes'),
    Privacy: 'Public',
    Status: 'Reported by family',
    Notes: ['Suggested by ' + (r.Uploader || 'a family member') +
            (r.Contact ? ' (' + r.Contact + ')' : '') + ', ' + rel + ' ' + related.PersonID + '.',
            r.Body ? 'They wrote: ' + r.Body : ''].concat(notes)
           .filter(function (x) { return x; }).join(' ')
  });

  // "Parent of" points the other way round: the existing person gains a parent.
  if (rel.indexOf('parent') === 0) {
    const t2 = table_('PEOPLE');
    const target = t2.rows.filter(function (x) { return x.PersonID === related.PersonID; })[0];
    const field = gender === 'F' ? 'MotherID' : 'FatherID';
    const col = t2.header.indexOf(field) + 1;
    const iNotes = t2.header.indexOf('Notes') + 1;
    if (target && col) {
      if (!String(target[field] || '').trim()) {
        t2.sheet.getRange(target._row, col).setValue(id);
      } else {
        t2.sheet.getRange(target._row, iNotes).setValue(
          (target.Notes ? target.Notes + ' ' : '') +
          'A family member suggested ' + id + ' as another ' + field.replace('ID', '') +
          ', but one is already recorded. Please resolve.');
      }
    }
  }

  if (rel.indexOf('spouse') === 0 || rel.indexOf('partner') === 0) {
    appendRow_('RELATIONSHIPS', {
      RelID: nextId_('RELATIONSHIPS', 'RelID', 'R', 3),
      Person1ID: id, Person2ID: related.PersonID, Type: 'Spouse',
      Status: 'Reported by family', Notes: 'From a family suggestion.'
    });
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
    const addedPeople = [];

    for (var r = rng.getRow(); r < rng.getRow() + rng.getNumRows(); r++) {
      if (r === 1) continue;
      const hit = t.rows.filter(function (x) { return x._row === r; })[0];
      if (!hit) continue;
      if (!/^pending$/i.test(hit.Status)) { skipped++; continue; }

      if (approve) {
        const id = publishInboxRow_(hit);
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
  const box = inbox_();
  var moved = 0;

  [['PHOTOS', 'Photo'], ['STORIES', 'Story']].forEach(function (spec) {
    const t = table_(spec[0]);
    const pending = t.rows.filter(function (r) { return /^pending$/i.test(r.ApprovalStatus); });
    // Delete from the bottom up so earlier row numbers stay valid.
    pending.sort(function (a, b) { return b._row - a._row; }).forEach(function (r) {
      box.appendRow([
        r.UploadedAt || r.RecordedDate || new Date(), spec[1], r.PersonID, personName_(r.PersonID),
        r.Caption || r.Title || '', r.Story || '', r.DriveFileID || '',
        r.PhotoDate || '', r.Place || '', r.Uploader || r.ToldBy || '', 'Pending', ''
      ]);
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

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
          out[t] = table_(t).rows.map(function (r) { delete r._row; return r; });
        } catch (err) { out[t] = []; }
      });
      return json_(out);
    }
    if (action === 'pending') {
      return json_({
        photos: table_('PHOTOS').rows.filter(function (r) { return /^pending$/i.test(r.ApprovalStatus); }),
        stories: table_('STORIES').rows.filter(function (r) { return /^pending$/i.test(r.ApprovalStatus); })
      });
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
      case 'suggest':     return json_(addSuggestion_(body));
      default:            return json_({ ok: false, error: 'Unknown action.' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
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

  const photoId = nextId_('PHOTOS', 'PhotoID', 'F', 3);
  appendRow_('PHOTOS', {
    PhotoID: photoId,
    PersonID: b.personId,
    DriveFileID: file.getId(),
    Caption: b.caption || '',
    PhotoDate: b.photoDate || '',
    Place: b.place || '',
    PeopleShown: b.personId,
    Uploader: b.uploader || 'Anonymous',
    UploadedAt: new Date().toISOString().slice(0, 19).replace('T', ' '),
    ApprovalStatus: 'Pending',
    IsProfile: 'No',
    Notes: 'Submitted through the website.'
  });
  notifyAdmin_('New photograph awaiting approval',
    (b.uploader || 'Someone') + ' submitted a photograph for ' + personName_(b.personId) + '.');
  return { ok: true, photoId: photoId };
}

function addStory_(b) {
  if (!b.personId || !b.story) return { ok: false, error: 'Missing person or story.' };
  const storyId = nextId_('STORIES', 'StoryID', 'S', 3);
  appendRow_('STORIES', {
    StoryID: storyId,
    PersonID: b.personId,
    Title: b.title || 'Untitled',
    Story: b.story,
    ToldBy: b.toldBy || '',
    RecordedDate: new Date().toISOString().slice(0, 10),
    Category: b.category || 'Memory',
    ApprovalStatus: 'Pending',
    Notes: 'Submitted through the website.'
  });
  notifyAdmin_('New story awaiting approval',
    (b.toldBy || 'Someone') + ' shared a story about ' + personName_(b.personId) + '.');
  return { ok: true, storyId: storyId };
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
  const photos = table_('PHOTOS').rows.filter(function (r) { return /^pending$/i.test(r.ApprovalStatus); });
  const stories = table_('STORIES').rows.filter(function (r) { return /^pending$/i.test(r.ApprovalStatus); });
  if (!photos.length && !stories.length) {
    ui_().alert('Approval queue', 'Nothing is waiting.', ui_().ButtonSet.OK);
    return;
  }
  ui_().alert('Approval queue',
    'Photographs waiting: ' + photos.length + '\n' +
    (photos.length ? photos.map(function (r) { return '  ' + r.PhotoID + ' · ' + personName_(r.PersonID) + ' · from ' + (r.Uploader || '?'); }).join('\n') + '\n' : '') +
    '\nStories waiting: ' + stories.length + '\n' +
    (stories.length ? stories.map(function (r) { return '  ' + r.StoryID + ' · ' + personName_(r.PersonID) + ' · "' + r.Title + '"'; }).join('\n') : '') +
    '\n\nTo act on one: go to the PHOTOS or STORIES sheet, select the row, then use\n' +
    'Family Tree ▸ Approve the selected rows.',
    ui_().ButtonSet.OK);
}

function setSelectionStatus_(value) {
  const sh = book_().getActiveSheet();
  const nm = sh.getName();
  if (nm !== 'PHOTOS' && nm !== 'STORIES') {
    ui_().alert('Select rows on the PHOTOS or STORIES sheet first.');
    return;
  }
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  const col = header.indexOf('ApprovalStatus') + 1;
  if (!col) { ui_().alert('That sheet has no ApprovalStatus column.'); return; }

  const rng = sh.getActiveRange();
  var n = 0;
  for (var r = rng.getRow(); r < rng.getRow() + rng.getNumRows(); r++) {
    if (r === 1) continue;
    sh.getRange(r, col).setValue(value);
    n++;
  }
  ui_().alert(value, n + ' row(s) set to ' + value + '.', ui_().ButtonSet.OK);
}
function approveSelection() { setSelectionStatus_('Approved'); }
function rejectSelection()  { setSelectionStatus_('Rejected'); }

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

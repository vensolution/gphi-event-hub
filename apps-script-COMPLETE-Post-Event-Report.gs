/**
 * POST EVENT REPORT  ->  Google Sheet tab + Drive photo folder + Google Doc  (+ promoter photos in ONE Google Doc)
 * COMPLETE SCRIPT: delete everything in Code.gs and paste this whole file.
 * Deploy as Web app (Execute as: Me, Who has access: Anyone) > Manage deployments > pencil > New version.
 *
 * Script properties (Project Settings > Script properties):
 *   ADMIN_PASSCODE   same passcode you use to sign in to the admin page
 *   GEMINI_API_KEY   (for "Improve writing") free key from aistudio.google.com
 *   GEMINI_MODEL     (optional) defaults to gemini-4-hydrogen
 */

var PARENT_FOLDER_NAME = 'POST EVENT REPORTS';   // created automatically next to your Sheet

/* ------------------------------ router ------------------------------ */

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSCODE');
    if (!pw || String(d.pass || '') !== String(pw)) return json({ ok: false, error: 'Wrong passcode' });
    switch (d.action) {
      case 'photo':          return savePhoto(d);
      case 'doc':            return makeDoc(d);
      case 'log':            return saveLog(d);
      case 'statuses':       return getStatuses();
      case 'improve':        return improveText(d);
      case 'master':         return updateMaster(d);
      case 'promoterFolder': return json(promoterFolder_(d));
      case 'promoterPhoto':  return json(promoterPhoto_(d));
      case 'promoterDoc':    return json(promoterDoc_(d));
      default:
        if (!d.action) return makeReport(d);
        return json({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* File name:  PR_11/04/2026_FEEDING PROGRAM_BRGY 43 ZONE 6 PASAY CITY_POST_EVENT_REPORT
 * ("PR_" when a MarComms person in charge is filled in, otherwise "EVENTS_") */
function baseName(d) {
  var p = String(d.eventDateText || '').split('/');
  var date = p.length === 3 ? ('0' + p[0]).slice(-2) + '/' + ('0' + p[1]).slice(-2) + '/' + p[2] : String(d.eventDateText || '');
  var pr = String(d.picMarcomms || '').trim() ? 'PR_' : 'EVENTS_';
  return (pr + date + '_' + d.eventName + '_' + d.venue).toUpperCase().replace(/\s+/g, ' ').trim();
}

/* ------------------ 1) one event = one Google Sheet file ------------------ */

function makeReport(d) {
  var fileName = baseName(d) + '_POST_EVENT_REPORT';
  var top = topFolder();
  var folder = firstFolder(top, fileName) || top.createFolder(fileName);   // folder is named like the Sheet file

  // re-sending reuses the same file (same link) and rebuilds it
  var ss = null, files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    var f = files.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) { ss = SpreadsheetApp.openById(f.getId()); break; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  }

  var sh = ss.insertSheet('tmp' + Date.now());
  ss.getSheets().forEach(function (s) { if (s.getSheetId() !== sh.getSheetId()) ss.deleteSheet(s); });
  sh.setName('POST EVENT REPORT');
  buildReport(sh, d);

  return json({ ok: true, tab: fileName, sheetUrl: ss.getUrl(), folderId: folder.getId(), folderUrl: folder.getUrl() });
}

/* tiny rich-text builder: label normal, value bold */
function RT() { this.t = ''; this.r = []; }
RT.prototype.add = function (s, bold) {
  s = String(s == null ? '' : s);
  if (bold && s) this.r.push([this.t.length, this.t.length + s.length]);
  this.t += s; return this;
};
RT.prototype.nl = function () { this.t += '\n'; return this; };
RT.prototype.kv = function (label, value) { return this.add(label + ' ').add(value, true).nl(); };
RT.prototype.block = function (label, value) {            // long answers: label, then answer below
  return this.add(label).nl().add(value, true).nl().nl();
};
RT.prototype.tick = function (label, value, options) {
  this.add(label + ' ');
  var self = this;
  options.forEach(function (o, i) {
    var sel = value === o;
    self.add('(' + (sel ? 'X' : ' ') + ') ' + o, sel);
    if (i < options.length - 1) self.add(' ');
  });
  return this.nl();
};

function buildReport(sh, d) {
  sh.setColumnWidth(1, 750);
  sh.setColumnWidth(2, 170);

  // header rows (same layout as the template)
  sh.getRange('A1:B1').merge().setValue('EVENT COORDINATOR REPORT & SURVEY FORM')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2').setValue('Event Name: ' + d.eventName);
  sh.getRange('B2').setValue('Event Date: ' + d.eventDateText);
  sh.getRange('A3').setValue('Event Venue (Barangay and Municipality): ' + d.venue);
  sh.getRange('B3').setValue('Report Date: ' + d.reportDateText);
  sh.getRange('A4').setValue('Person In Charge: ' + d.picEvents + ' (Events) ' + d.picMarcomms + ' (MarComms)');
  sh.getRange('A2:B4').setFontWeight('bold');
  sh.setRowHeight(5, 14);

  var YN = ['Yes', 'No'], RATE = ['Excellent', 'Good', 'Fair', 'Needs Improvement'];
  var blocks = [];

  var a = new RT();
  a.add('I. ACTUAL EVENT ATTENDANCE', true).nl().nl()
   .kv('Total actual participants:', d.total)
   .kv('Participants aged 21 and above:', d.over21)
   .kv('Participants aged 21 below:', d.under21)
   .kv('New promoters recruited during the event:', d.newRecruited)
   .kv('Existing promoters who participated:', d.existingPromoters);
  blocks.push(a);

  var b = new RT();
  b.add('II. PROMOTER ENGAGEMENT', true).nl().nl()
   .kv('Participants interested in becoming promoters:', d.interested)
   .kv('New promoters successfully registered:', d.newRegistered)
   .kv('New promoters who actively participated:', d.newActive);
  blocks.push(b);

  var c = new RT();
  c.add('III. BARANGAY FEEDBACK AND COMMUNITY RESPONSE', true).nl().nl()
   .kv('Name and position of the barangay representative:', d.brgyRep)
   .kv('Feedback of the barangay officials regarding the event:', d.brgyFeedback)
   .tick('Were barangay officials or representatives present during the event?', d.brgyPresent, YN)
   .tick('Did the barangay express support or interest in future activities?', d.brgySupport, YN)
   .tick('Were there any concerns or recommendations raised by the barangay?', d.brgyConcerns, YN)
   .tick("Overall assessment of the barangay's response:", d.brgyRating, RATE)
   .nl()
   .kv('Recommendations of the Barangay for future events:', d.brgyReco);
  blocks.push(c);

  var e = new RT();
  e.add('IV. EVENT FEEDBACK', true).nl().nl()
   .block('Most successful activity:', d.mostSuccessful)
   .block('Challenges encountered:', d.challenges)
   .block('Actions taken to address these challenges:', d.actionsTaken)
   .tick('Did the event meet its target attendance?', d.metAttendance, YN)
   .tick('Did the team achieve its target for new promoters?', d.metPromoters, YN)
   .tick('Overall event rating:', d.eventRating, RATE)
   .nl()
   .block('Recommendations for future events:', d.recommendations);
  blocks.push(e);

  var f = new RT();
  f.add('Prepared by: ', true).add(d.preparedBy)
   .add('          Signature: ', true).add('______________________________').nl()
   .add('Date: ', true).add(d.preparedDateText)
   .add('                    Reviewed by: ', true).add(d.reviewedBy);
  blocks.push(f);

  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  var row = 6;
  blocks.forEach(function (rt) {
    var text = rt.t.replace(/\n+$/, '');
    var builder = SpreadsheetApp.newRichTextValue().setText(text);
    rt.r.forEach(function (x) { if (x[1] <= text.length) builder.setTextStyle(x[0], x[1], bold); });
    sh.getRange(row, 1, 1, 2).merge().setRichTextValue(builder.build());

    var lines = 0;                                           // merged cells do not auto-resize
    text.split('\n').forEach(function (l) { lines += Math.max(1, Math.ceil(l.length / 135)); });
    sh.setRowHeight(row, lines * 14 + 36);
    row++;
  });

  var all = sh.getRange(1, 1, row - 1, 2);
  all.setFontFamily('Arial').setFontSize(10).setWrap(true).setVerticalAlignment('middle');
  sh.getRange('A1:B1').setHorizontalAlignment('center');
  all.setBorder(true, true, true, true, true, true);
  sh.setHiddenGridlines(true);
}

/* ---------------------- 2) Drive folder + photos --------------------- */

function topFolder() {
  var parent;
  try { parent = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next(); }
  catch (e) { parent = DriveApp.getRootFolder(); }
  return firstFolder(parent, PARENT_FOLDER_NAME) || parent.createFolder(PARENT_FOLDER_NAME);
}

function firstFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function savePhoto(d) {
  var folder = DriveApp.getFolderById(d.folderId);
  if (d.replace) {                                          // first photo of a re-send: clear the old ones
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      if (String(f.getMimeType()).indexOf('image/') === 0) f.setTrashed(true);
    }
  }
  var name = ('0' + d.index).slice(-2) + '.jpg';
  var blob = Utilities.newBlob(Utilities.base64Decode(d.data), d.mime || 'image/jpeg', name);
  folder.createFile(blob);
  return json({ ok: true, index: d.index });
}

/* ------------------------ 3) Photo documentation --------------------- */

/* shared: ONE-PAGE photo grid, picks columns x rows that make the photos as large as possible */
function photoGrid_(body, images, AH) {
  var N = images.length; if (!N) return;
  var AW = 504, best = null;
  for (var cols = 1; cols <= 6; cols++) {
    var rows = Math.ceil(N / cols);
    var s = Math.min(AW / cols, AH / rows), waste = cols * rows - N;
    if (!best || s > best.s + 0.01 || (Math.abs(s - best.s) <= 0.01 && waste < best.waste)) best = { cols: cols, rows: rows, s: s, waste: waste };
  }
  var cw = AW / best.cols, ch = Math.min(AH / best.rows, cw * 1.4);
  var data = [];
  for (var rr = 0; rr < best.rows; rr++) { var rw = []; for (var cc = 0; cc < best.cols; cc++) rw.push(''); data.push(rw); }
  var t = body.appendTable(data);
  t.setBorderWidth(0.5).setBorderColor('#d9d9d9');
  for (var c0 = 0; c0 < best.cols; c0++) t.setColumnWidth(c0, cw);
  var idx = 0;
  for (var r1 = 0; r1 < best.rows; r1++) {
    t.getRow(r1).setMinimumHeight(ch);
    for (var c1 = 0; c1 < best.cols; c1++, idx++) {
      var cell = t.getCell(r1, c1);
      cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(3).setPaddingRight(3)
          .setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      var par = cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
      if (idx < N) {
        var img = par.appendInlineImage(images[idx].getBlob());
        var w = img.getWidth(), h = img.getHeight();
        var sc = Math.min((cw - 8) / w, (ch - 14) / h);
        img.setWidth(Math.round(w * sc)).setHeight(Math.round(h * sc));
      }
    }
  }
}

/* shared: page setup + logo + 3 centered lines */
function docHeader_(doc, d) {
  var body = doc.getBody();
  body.setPageWidth(612).setPageHeight(792);                 // Letter (short bond)
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);
  var p0 = body.getParagraphs()[0].setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingBefore(0).setSpacingAfter(2);
  var logo = p0.appendInlineImage(Utilities.newBlob(Utilities.base64Decode(LOGO_B64), 'image/png', 'logo.png'));
  logo.setWidth(90).setHeight(Math.round(90 * logo.getHeight() / logo.getWidth()));
  [['LOCATION: ', d.location], ['EVENT NAME: ', d.eventName], ['DATE: ', d.dateLong]].forEach(function (x) {
    body.appendParagraph(x[0] + (x[1] || '')).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true)
      .setLineSpacing(1).setFontFamily('Arial').setFontSize(10).setSpacingBefore(0).setSpacingAfter(0);
  });
  return body;
}

function imagesIn_(folder) {
  var images = [], it = folder.getFiles();
  while (it.hasNext()) { var f = it.next(); if (String(f.getMimeType()).indexOf('image/') === 0) images.push(f); }
  images.sort(function (x, y) { return x.getName() < y.getName() ? -1 : 1; });
  return images;
}

function makeDoc(d) {
  var folder = DriveApp.getFolderById(d.folderId);
  var name = String(d.tab).replace(/_POST_EVENT_REPORT$/, '') + '_PHOTO_DOCUMENTATION';

  var old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);

  var images = imagesIn_(folder);
  var doc = DocumentApp.create(name);
  var body = docHeader_(doc, d);
  body.appendParagraph('').setFontSize(6).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  photoGrid_(body, images, 590);

  // keep the unavoidable last empty paragraph from pushing onto a 2nd page
  var ps = body.getParagraphs();
  ps[ps.length - 1].setFontSize(1).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);

  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);
  return json({ ok: true, docUrl: file.getUrl(), photos: images.length });
}

/* ------- 3b) PROMOTER photos: ONE Google Doc for ALL promoters of the event ------- */

function promoterFolder_(d) {
  var parent = DriveApp.getRootFolder();
  var m = String(d.parentFolderUrl || '').match(/[-\w]{25,}/);
  if (m) { try { parent = DriveApp.getFolderById(m[0]); } catch (e) {} }
  var name = 'PROMOTER PHOTOS - ' + d.eventName + (d.promoter ? ' - ' + d.promoter : '');
  var it = parent.getFoldersByName(name);
  var f = it.hasNext() ? it.next() : parent.createFolder(name);
  return { ok: true, folderId: f.getId(), folderUrl: f.getUrl() };
}

function promoterPhoto_(d) {
  var f = DriveApp.getFolderById(d.folderId), nm;
  if (d.promoterIndex) {                                    // new: p01_03.jpg = promoter 1, photo 3
    nm = 'p' + ('0' + d.promoterIndex).slice(-2) + '_' + ('0' + d.index).slice(-2) + '.jpg';
    var o = f.getFilesByName(nm); while (o.hasNext()) o.next().setTrashed(true);
  } else {                                                  // old site version: one folder per promoter
    nm = 'promoter_' + ('0' + d.index).slice(-2) + '.jpg';
    if (d.replace) { var all = f.getFiles(); while (all.hasNext()) { var x = all.next(); if (String(x.getMimeType()).indexOf('image/') === 0) x.setTrashed(true); } }
  }
  f.createFile(Utilities.newBlob(Utilities.base64Decode(d.data), d.mime || 'image/jpeg', nm));
  return { ok: true };
}

function promoterDoc_(d) {
  var folder = DriveApp.getFolderById(d.folderId);
  var name = (String(d.eventName).toUpperCase() + '_PROMOTER_PHOTO_DOCUMENTATION').replace(/\s+/g, ' ');
  var old = folder.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true);

  var legacy = !d.promoters || !d.promoters.length;          // old site version sends one promoter at a time
  var plist = legacy ? [{ name: d.promoterName, user: d.promoterUser, type: d.promoterType }] : d.promoters;
  var images = imagesIn_(folder);

  var doc = DocumentApp.create(name);
  var body = docHeader_(doc, d);
  plist.forEach(function (p, i) {
    if (i > 0) body.appendPageBreak();
    var u = String(p.user || '').replace(/^@/, '');
    var line = 'PROMOTER ' + (i + 1) + ': ' + String(p.name || '').toUpperCase() + (u ? ' (@' + u + ')' : '') + (p.type ? ' - ' + String(p.type).toUpperCase() : '');
    body.appendParagraph(line).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setLineSpacing(1)
        .setFontFamily('Arial').setFontSize(10).setSpacingBefore(i === 0 ? 8 : 0).setSpacingAfter(6);
    var tag = 'p' + ('0' + (i + 1)).slice(-2) + '_';
    var mine = legacy ? images : images.filter(function (x) { return x.getName().indexOf(tag) === 0; });
    photoGrid_(body, mine, i === 0 ? 530 : 620);
  });
  var ps = body.getParagraphs();
  ps[ps.length - 1].setFontSize(1).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);
  return { ok: true, docUrl: file.getUrl(), promoters: plist.length, photos: images.length };
}

/* ------------- 5) "Event Details" master list (existing Google Sheet) ------------- */

var MASTER_SHEET_ID = '16JqhR7leze6uHd2lyGA5NM7QRXD6UJbpKSbJgmXqTJ0';
var MASTER_GID = 0;                                           // the first tab
var STATUS_OPTIONS = ['Pending', 'Submitted', 'For Review', 'Approved'];
var STATUS_DEFAULT = 'Submitted';

function eventLine(d) {                                       // same wording as the file name: EVENT NAME_VENUE
  return (d.eventName + '_' + d.venue).toUpperCase().replace(/\s+/g, ' ').trim();
}

function setLink_(cell, url) {
  if (!url) { cell.setValue(''); return; }
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build())
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
}

function updateMaster(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var sh = null;
    ss.getSheets().forEach(function (s) { if (!sh && s.getName().trim().toLowerCase() === 'event details') sh = s; });
    if (!sh) sh = ss.getSheets()[0];

    // find this event's block (re-send) or the next free block (new)
    var id = (String(d.sheetUrl || '').match(/\/d\/([\w-]+)/) || [])[1] || '';
    var last = Math.max(sh.getLastRow(), 2);
    var B = sh.getRange(1, 2, last, 1).getValues(), C = sh.getRange(1, 3, last, 1).getValues();
    var start = 0, lastDate = 0;
    for (var i = 0; i < last; i++) {
      var b = String(B[i][0]).trim();
      if (b === 'Date') lastDate = i + 1;
      if (id && /^Link \(Post/i.test(b) && String(C[i][0]).indexOf(id) > -1) start = i + 1 - 2;
    }
    var isNew = !start;
    if (isNew) start = lastDate ? lastDate + 6 : 2;           // each block = 5 rows + 1 spacer row

    if (sh.getMaxRows() < start + 5) sh.insertRowsAfter(sh.getMaxRows(), start + 5 - sh.getMaxRows());

    var block = sh.getRange(start, 1, 6, 10);

    sh.getRange(start, 2, 5, 1)
      .setValues([['Date'], ['Event'], ['Link (Post Event)'], ['Link (photos/folder)'], ['Link (documentation)']])
      .setFontWeight('bold');

    var p = String(d.eventDateText || '').split('/');
    var dateText = p.length === 3 ? ('0' + p[0]).slice(-2) + '/' + ('0' + p[1]).slice(-2) + '/' + p[2] : String(d.eventDateText || '');
    sh.getRange(start, 3).setNumberFormat('@').setValue(dateText).setHorizontalAlignment('center');
    sh.getRange(start + 1, 3).setValue(eventLine(d));
    setLink_(sh.getRange(start + 2, 3), d.sheetUrl);
    setLink_(sh.getRange(start + 3, 3), d.folderUrl);
    setLink_(sh.getRange(start + 4, 3), d.docUrl);

    // name goes in the top-left cell of its merged cell, if it is merged
    try {
      var nameCell = sh.getRange(start + 2, 1), nm = nameCell.getMergedRanges();
      (nm.length ? nm[0].getCell(1, 1) : nameCell).setValue(d.preparedBy).setFontWeight('bold').setHorizontalAlignment('center');
    } catch (e) {}

    if (isNew) {
      try {
        var jm = sh.getRange(start, 10).getMergedRanges();
        var target = jm.length ? jm[0] : sh.getRange(start, 10, 5, 1);
        if (!jm.length && sh.getRange(start, 10, 5, 1).getMergedRanges().length === 0) target.merge();   // only merge when nothing overlaps
        target.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(false).build())
              .setHorizontalAlignment('center').setVerticalAlignment('middle');
        target.getCell(1, 1).setValue(STATUS_DEFAULT);
        SpreadsheetApp.flush();
      } catch (fmtErr) {}
      try { block.setBorder(true, true, true, true, true, true); SpreadsheetApp.flush(); } catch (e) {}
    }
    SpreadsheetApp.flush();
    return json({ ok: true, row: start, isNew: isNew });
  } finally {
    lock.releaseLock();
  }
}

/* Run this from the editor (choose checkMaster > Run) to test the connection to the Event Details sheet.
 * It only READS: it writes nothing. Read the result in the Execution log. */
function checkMaster() {
  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sh = null;
  ss.getSheets().forEach(function (s) { if (!sh && s.getName().trim().toLowerCase() === 'event details') sh = s; });
  if (!sh) sh = ss.getSheets()[0];
  var last = Math.max(sh.getLastRow(), 2), B = sh.getRange(1, 2, last, 1).getValues(), lastDate = 0, n = 0;
  for (var i = 0; i < last; i++) { if (String(B[i][0]).trim() === 'Date') { lastDate = i + 1; n++; } }
  Logger.log('Opened "' + ss.getName() + '", tab "' + sh.getName() + '". Blocks found: ' + n +
             '. Last "Date" row: ' + lastDate + '. The next block will start at row ' + (lastDate ? lastDate + 6 : 2) + '.');
}

/* ------------- 4) status log: tells the admin page SENT / PENDING ------------- */

function logSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('REPORT LOG');
  if (!sh) {
    sh = ss.insertSheet('REPORT LOG');
    sh.appendRow(['Key', 'Event', 'Event Date', 'Status', 'Report Sheet', 'Photo Folder', 'Google Doc', 'Sent At']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
  }
  return sh;
}

function saveLog(d) {
  if (!d.key) return json({ ok: false, error: 'Missing key' });
  var sh = logSheet_();
  var vals = sh.getDataRange().getValues(), row = -1;
  for (var i = 1; i < vals.length; i++) { if (vals[i][0] === d.key) { row = i + 1; break; } }
  var rec = [d.key, d.eventName, d.eventDateText, 'SENT', d.sheetUrl || '', d.folderUrl || '', d.docUrl || '', new Date()];
  if (row > 0) sh.getRange(row, 1, 1, rec.length).setValues([rec]); else sh.appendRow(rec);
  return json({ ok: true });
}

function getStatuses() {
  var vals = logSheet_().getDataRange().getValues(), map = {};
  for (var i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    map[vals[i][0]] = { sheetUrl: vals[i][4], folderUrl: vals[i][5], docUrl: vals[i][6], at: String(vals[i][7]) };
  }
  return json({ ok: true, map: map });
}

/* --------------------- FREE AI "Improve writing" --------------------- */

var IMPROVE_SYSTEM = "Rewrite these Philippine event-report notes as clear, grammatically correct, professional English. " +
  "Keep all facts, names, numbers and dates; add nothing. Translate Tagalog/Taglish to English. " +
  "Plain text only. Output only the rewritten text.";

function improveText(d) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) return json({ error: 'Not set up yet: add GEMINI_API_KEY in Script properties.' });

  var text = String(d.text || '').trim();
  if (text.length < 5) return json({ error: 'Text is too short.' });
  if (text.length > 4000) return json({ error: 'Text is too long (max 4000 characters).' });

  var model = props.getProperty('GEMINI_MODEL') || 'gemini-4-hydrogen';
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';

  // "thinking" levels to try, fastest first; the one that works is remembered so later calls make ONE request
  var LEVELS = ['minimal', 'low', null];
  var cache = CacheService.getScriptCache();
  var start = Number(cache.get('think_lv') || 0);
  var maxOut = Math.min(2048, Math.ceil(text.length / 2) + 256);

  function call(level) {
    var cfg = { temperature: 0.2, maxOutputTokens: maxOut };
    if (level) cfg.thinkingConfig = { thinkingLevel: level };
    return UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key.trim() },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        system_instruction: { parts: [{ text: IMPROVE_SYSTEM }] },
        contents: [{ role: 'user', parts: [{ text: 'Report field: ' + String(d.field || 'Event feedback').slice(0, 120) +
                                                    '\n\nNotes to rewrite:\n' + text }] }],
        generationConfig: cfg
      })
    });
  }

  try {
    var r, n = start;
    for (; n < LEVELS.length; n++) {
      r = call(LEVELS[n]);
      if (r.getResponseCode() !== 400) break;               // 400 = this thinking level is not accepted, try the next
    }
    if (n >= LEVELS.length) n = LEVELS.length - 1;
    if (n !== start) cache.put('think_lv', String(n), 21600);
    var code = r.getResponseCode();
    var data = JSON.parse(r.getContentText());
    if (code === 429) return json({ error: 'Free limit reached for now. Please try again in a minute.' });
    if (code !== 200) return json({ error: (data.error && data.error.message) || 'AI service error' });
    var parts = (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) || [];
    var out = parts.map(function (p) { return p.text || ''; }).join('').trim();
    if (!out) return json({ error: 'Empty response' });
    return json({ text: out });
  } catch (err) {
    return json({ error: 'Could not reach the AI service.' });
  }
}

/* Run this ONE time from the editor so Google asks you to approve Drive, Docs and external requests. */
function authorizeOnce() {
  DriveApp.getRootFolder();
  try { UrlFetchApp.fetch('https://generativelanguage.googleapis.com/'); } catch (e) {}
  Logger.log('Authorized. Now deploy a New version.');
}

function testMaster() {
  var out = updateMaster({
    eventName: 'TEST EVENT', venue: 'TEST VENUE', eventDateText: '11/23/2026', preparedBy: 'Test',
    sheetUrl: 'https://docs.google.com/spreadsheets/d/TESTID123/edit',
    folderUrl: 'https://drive.google.com/drive/folders/TESTFOLDER', docUrl: ''
  });
  Logger.log(out.getContent());
}

/* ------------------------------ logo (PNG) --------------------------- */
var LOGO_B64 = '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';

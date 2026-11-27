// ONE Google Doc for ALL promoters of an event. Paste at the bottom of the Apps Script behind PER_CONFIG.APPS_SCRIPT_URL.
// In doPost, after your passcode check and before  if (!d.action) return makeReport(d);  add:
//   if (d.action === 'promoterFolder') return json(promoterFolder_(d));
//   if (d.action === 'promoterPhoto')  return json(promoterPhoto_(d));
//   if (d.action === 'promoterDoc')    return json(promoterDoc_(d));
// Then Deploy > Manage deployments > pencil > New version > Deploy.
function promoterFolder_(d) {                       // one folder per EVENT (not per promoter)
  var parent = DriveApp.getRootFolder();
  var m = String(d.parentFolderUrl || '').match(/[-\w]{25,}/);
  if (m) { try { parent = DriveApp.getFolderById(m[0]); } catch (e) {} }
  var name = 'PROMOTER PHOTOS - ' + d.eventName;
  var it = parent.getFoldersByName(name);
  var f = it.hasNext() ? it.next() : parent.createFolder(name);
  return { ok: true, folderId: f.getId(), folderUrl: f.getUrl() };
}
function promoterPhoto_(d) {                        // file name = promoter number + photo number, so the order is kept
  var f = DriveApp.getFolderById(d.folderId);
  var nm = 'p' + ('0' + d.promoterIndex).slice(-2) + '_' + ('0' + d.index).slice(-2) + '.jpg';
  var old = f.getFilesByName(nm); while (old.hasNext()) old.next().setTrashed(true);
  f.createFile(Utilities.newBlob(Utilities.base64Decode(d.data), d.mime || 'image/jpeg', nm));
  return { ok: true };
}
function promoterDoc_(d) {                          // d.promoters = [{name, user, type}] in order
  var f = DriveApp.getFolderById(d.folderId);
  var doc = DocumentApp.create(d.eventName + ' - PROMOTER PHOTO DOCUMENTATION');
  var body = doc.getBody();
  body.appendParagraph('PHOTO DOCUMENTATION').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Event: ' + d.eventName);
  body.appendParagraph('Location: ' + (d.location || ''));
  body.appendParagraph('Date: ' + (d.dateLong || ''));
  body.appendParagraph('Promoters: ' + d.promoters.length);
  var files = []; var it = f.getFiles(); while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
  d.promoters.forEach(function (p, i) {
    body.appendPageBreak();
    var tag = 'p' + ('0' + (i + 1)).slice(-2) + '_';
    body.appendParagraph('PROMOTER ' + (i + 1) + ': ' + p.name + (p.user ? ' (@' + String(p.user).replace(/^@/, '') + ')' : '') + (p.type ? ' - ' + p.type : ''))
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    files.filter(function (x) { return x.getName().indexOf(tag) === 0; }).forEach(function (file) {
      var img = body.appendImage(file.getBlob());
      var w = img.getWidth(), h = img.getHeight(), max = 440;
      if (w > max) img.setWidth(max).setHeight(Math.round(h * max / w));
    });
  });
  doc.saveAndClose();
  var df = DriveApp.getFileById(doc.getId());
  f.addFile(df); DriveApp.getRootFolder().removeFile(df);
  return { ok: true, docUrl: doc.getUrl() };
}

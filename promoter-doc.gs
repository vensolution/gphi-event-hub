// Add these actions to the Apps Script that already handles the Post Event Report (the one at PER_CONFIG.APPS_SCRIPT_URL in bybet/admin.html),
// then Deploy > Manage deployments > Edit > New version.
// In your existing doPost, add before the "unknown action" fallback:
//   if (data.action === 'promoterFolder') return json_(promoterFolder_(data));
//   if (data.action === 'promoterPhoto')  return json_(promoterPhoto_(data));
//   if (data.action === 'promoterDoc')    return json_(promoterDoc_(data));
// Use the same passcode check and json_ helper your script already uses for the other actions.
// The Doc below is a standard photo documentation layout. To get EXACTLY the same format as your existing Photo Documentation,
// call your existing doc-building function inside promoterDoc_ instead of buildPromoterDoc_.

function promoterFolder_(d) {
  var parent = DriveApp.getRootFolder();
  var m = String(d.parentFolderUrl || '').match(/[-\w]{25,}/);
  if (m) { try { parent = DriveApp.getFolderById(m[0]); } catch (e) {} }
  var name = 'PROMOTER PHOTOS - ' + d.eventName + ' - ' + (d.promoter || '');
  var it = parent.getFoldersByName(name);
  var f = it.hasNext() ? it.next() : parent.createFolder(name);
  return { ok: true, folderId: f.getId(), folderUrl: f.getUrl() };
}
function promoterPhoto_(d) {
  var f = DriveApp.getFolderById(d.folderId);
  if (d.replace) { var old = f.getFiles(); while (old.hasNext()) old.next().setTrashed(true); }
  var blob = Utilities.newBlob(Utilities.base64Decode(d.data), d.mime || 'image/jpeg', 'promoter_' + ('0' + d.index).slice(-2) + '.jpg');
  f.createFile(blob);
  return { ok: true };
}
function promoterDoc_(d) {
  var f = DriveApp.getFolderById(d.folderId);
  var doc = DocumentApp.create(d.eventName + ' - PROMOTER PHOTO DOCUMENTATION');
  var body = doc.getBody();
  body.appendParagraph('PHOTO DOCUMENTATION').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Event: ' + d.eventName);
  body.appendParagraph('Location: ' + (d.location || ''));
  body.appendParagraph('Date: ' + (d.dateLong || ''));
  body.appendParagraph('Promoter: ' + d.promoterName + (d.promoterUser ? ' (@' + String(d.promoterUser).replace(/^@/, '') + ')' : '') + (d.promoterType ? ' - ' + d.promoterType : ''));
  var files = f.getFiles(), list = [];
  while (files.hasNext()) list.push(files.next());
  list.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
  list.forEach(function (file) {
    var img = body.appendImage(file.getBlob());
    var w = img.getWidth(), h = img.getHeight(), max = 440;
    if (w > max) { img.setWidth(max).setHeight(Math.round(h * max / w)); }
    body.appendParagraph('');
  });
  doc.saveAndClose();
  var df = DriveApp.getFileById(doc.getId());
  f.addFile(df); DriveApp.getRootFolder().removeFile(df);
  return { ok: true, docUrl: doc.getUrl() };
}

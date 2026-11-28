// ONE Google Doc for ALL promoters of an event.
// Paste these three functions at the very bottom of the Apps Script that handles the Post Event Report,
// REPLACING the older promoterFolder_, promoterPhoto_ and promoterDoc_ if you already pasted them.
// In doPost, the three lines go BEFORE:  if (!d.action) return makeReport(d);
//     if (d.action === 'promoterFolder') return json(promoterFolder_(d));
//     if (d.action === 'promoterPhoto')  return json(promoterPhoto_(d));
//     if (d.action === 'promoterDoc')    return json(promoterDoc_(d));
// Then: Deploy > Manage deployments > pencil > Version: New version > Deploy.

function promoterFolder_(d) {
  var parent = DriveApp.getRootFolder();
  var m = String(d.parentFolderUrl || '').match(/[-\w]{25,}/);
  if (m) { try { parent = DriveApp.getFolderById(m[0]); } catch (e) {} }
  var name = 'PROMOTER PHOTOS - ' + d.eventName + (d.promoter ? ' - ' + d.promoter : '');
  var it = parent.getFoldersByName(name);
  var f = it.hasNext() ? it.next() : parent.createFolder(name);
  return { ok: true, folderId: f.getId(), folderUrl: f.getUrl() };
}

// d.name is the file name, e.g. P01_03.jpg = promoter 1, photo 3. The first upload (replace = true) clears old files.
function promoterPhoto_(d) {
  var f = DriveApp.getFolderById(d.folderId);
  if (d.replace) { var old = f.getFiles(); while (old.hasNext()) old.next().setTrashed(true); }
  var blob = Utilities.newBlob(Utilities.base64Decode(d.data), d.mime || 'image/jpeg', d.name || ('photo_' + new Date().getTime() + '.jpg'));
  f.createFile(blob);
  return { ok: true };
}

// d.promoters = [{tag:'P01', name, user, type}, ...] in the order they appear on the page.
function promoterDoc_(d) {
  var f = DriveApp.getFolderById(d.folderId);
  var doc = DocumentApp.create(d.eventName + ' - PROMOTER PHOTO DOCUMENTATION');
  var body = doc.getBody();
  body.appendParagraph('PHOTO DOCUMENTATION').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Event: ' + d.eventName);
  body.appendParagraph('Location: ' + (d.location || ''));
  body.appendParagraph('Date: ' + (d.dateLong || ''));
  var list = d.promoters || [];
  body.appendParagraph('Promoters: ' + list.length);

  var files = {}, it = f.getFiles();
  while (it.hasNext()) { var fl = it.next(); files[fl.getName()] = fl; }

  list.forEach(function (p, idx) {
    if (idx > 0) body.appendPageBreak(); else body.appendParagraph('');
    body.appendParagraph('PROMOTER ' + (idx + 1) + ': ' + p.name).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Username: ' + (p.user ? '@' + String(p.user).replace(/^@/, '') : '-') + (p.type ? '   |   ' + p.type + ' promoter' : ''));
    var names = Object.keys(files).filter(function (n) { return n.indexOf(p.tag + '_') === 0; }).sort();
    var cols = 4, boxW = 118, boxH = 100;          // 4 x 5 = up to 20 photos fit on one page
    var table = body.appendTable();
    table.setBorderWidth(0);
    for (var i = 0; i < names.length; i += cols) {
      var row = table.appendTableRow();
      for (var c = 0; c < cols; c++) {
        var cell = row.appendTableCell('');
        if (i + c < names.length) {
          var img = cell.appendImage(files[names[i + c]].getBlob());
          var w = img.getWidth(), h = img.getHeight(), k = Math.min(boxW / w, boxH / h, 1);
          img.setWidth(Math.round(w * k)).setHeight(Math.round(h * k));
        }
      }
    }
  });

  doc.saveAndClose();
  var df = DriveApp.getFileById(doc.getId());
  f.addFile(df); DriveApp.getRootFolder().removeFile(df);
  return { ok: true, docUrl: doc.getUrl() };
}

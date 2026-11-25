// Add this to your Google Apps Script project (bound to the Google Sheet that holds your liquidations).
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone. Paste the /exec URL in liquidation-config.js.
// If your project already has a doPost, call handleLiquidation_(data) from it when data.action === 'liquidation'.
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action === 'liquidation') return handleLiquidation_(data);
  return ContentService.createTextOutput('ignored');
}
function handleLiquidation_(data) {
  var sh = SpreadsheetApp.getActive().getSheetByName('LIQUIDATION') || SpreadsheetApp.getActive().insertSheet('LIQUIDATION');
  var rows = data.rows;                       // first row = headers, then one row per expense
  var all = sh.getDataRange().getValues();
  if (all.length < 1 || all[0][0] === '') { sh.getRange(1, 1, 1, rows[0].length).setValues([rows[0]]); all = sh.getDataRange().getValues(); }
  // remove the old rows of this liquidation (marked in the last column) then append the new ones
  var idCol = rows[0].length + 1;
  for (var i = all.length - 1; i >= 1; i--) if (all[i][idCol - 1] === data.id) sh.deleteRow(i + 1);
  var out = rows.slice(1).map(function (r) { return r.concat([data.id]); });
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  return ContentService.createTextOutput('ok');
}

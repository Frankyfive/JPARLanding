// =============================================
// SHEETSETUP.gs - Creates and initializes tabs
// =============================================

function setupSheetTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headers = ['ID', 'Title', 'Content', 'Image URL', 'Status', 'Date', 'Tags', 'Updated At'];

  Object.values(CONFIG.SHEETS).forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);

    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      Logger.log('Created tab: ' + tabName);
    }

    // Only write headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#00326D')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  });

  SpreadsheetApp.getUi().alert('Sheet tabs are ready.');
}

// Utility: get a sheet by CONFIG name key
function getSheet(sheetKey) {
  var name = CONFIG.SHEETS[sheetKey] || sheetKey;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

// Utility: get all data from a sheet as array of objects
function getSheetData(sheetName, statusFilter) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var result  = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[CONFIG.COLS.ID]) continue; // skip empty rows

    var status = row[CONFIG.COLS.STATUS];
    if (statusFilter && statusFilter !== 'all' && status !== statusFilter) continue;

    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    result.push(obj);
  }

  return result;
}

// --------------------------------------------
// DEBUG: Log the full deployed payload to console
// Run this from the Apps Script editor to test
// --------------------------------------------
function debugDeployedPayload() {
  var payload = buildDeployedPayload();
  Logger.log(JSON.stringify(payload, null, 2));
}
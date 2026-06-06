// =============================================
// SHEETSETUP.gs - Sheet tab management
// =============================================

// Standard headers for every tab
var SHEET_HEADERS = [
  'ID', 'Title', 'Event Type', 'Description',
  'Date', 'Time', 'Info', 'URL', 'Graphic', 'Status', 'UPDATED_AT'
];

// --------------------------------------------
// ADD: Create a new sheet tab with standard headers.
// Called from the CMS Manager → Add New Tab menu.
// --------------------------------------------
function addNewTab(tabName) {
  var ui = SpreadsheetApp.getUi();

  if (!tabName) {
    var response = ui.prompt(
      'Add New Tab',
      'Enter the name for the new tab (e.g. "Announcements"):',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    tabName = response.getResponseText().trim();
  }

  if (!tabName) {
    ui.alert('Tab name cannot be empty.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (ss.getSheetByName(tabName)) {
    ui.alert('A tab named "' + tabName + '" already exists.');
    return;
  }

  var sheet = ss.insertSheet(tabName);
  sheet.appendRow(SHEET_HEADERS);
  sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#00326D')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  ui.alert(
    'Tab Created',
    '"' + tabName + '" is ready.\n\nClose and reopen the spreadsheet so it appears in the Add Content menu.',
    ui.ButtonSet.OK
  );
}

// --------------------------------------------
// UTILITY: Get a sheet by CONFIG key or raw name
// --------------------------------------------
function getSheet(sheetKey) {
  var name  = CONFIG.SHEETS[sheetKey] || sheetKey;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

// --------------------------------------------
// UTILITY: Read all rows from a tab as objects.
// statusFilter: 'deployed' | 'testing' | 'all'
// --------------------------------------------
function getSheetData(sheetName, statusFilter) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var result  = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[CONFIG.COLS.ID]) continue; // skip blank rows

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
// --------------------------------------------
function debugDeployedPayload() {
  var payload = buildDeployedPayload();
  Logger.log(JSON.stringify(payload, null, 2));
}

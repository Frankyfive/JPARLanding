// =============================================
// SHEETSETUP.gs - Sheet tab management
// =============================================

// Standard headers for every tab
var SHEET_HEADERS = [
  'ID', 'Title', 'Event Type', 'Description',
  'Date', 'Time', 'Info', 'URL', 'Graphic', 'Status', 'UPDATED_AT'
];

// --------------------------------------------
// SETUP: Ensure all 3 required tabs exist with
// the correct headers. Safe to run on existing
// sheets — only adds missing columns, never
// deletes data. Run this first if preview is blank.
// --------------------------------------------
function setupSheets() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var log = [];

  Object.values(CONFIG.SHEETS).forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);

    if (!sheet) {
      // Create tab from scratch
      sheet = ss.insertSheet(tabName);
      sheet.appendRow(SHEET_HEADERS);
      sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#00326D')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      log.push('Created: ' + tabName);
    } else {
      // Tab exists — check if headers match
      var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var needsfix = false;

      SHEET_HEADERS.forEach(function(h, i) {
        if (existing[i] !== h) needsfix = true;
      });

      if (needsfix) {
        // Write correct headers into row 1 (preserves all data rows)
        sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
        sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
          .setFontWeight('bold')
          .setBackground('#00326D')
          .setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
        log.push('Fixed headers: ' + tabName);
      } else {
        log.push('OK: ' + tabName);
      }
    }
  });

  ui.alert(
    'Sheet Setup Complete',
    log.join('\n') + '\n\nReopen the spreadsheet to refresh the Add Content menu.',
    ui.ButtonSet.OK
  );
}

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
// BACKFILL: Add missing ID and UPDATED_AT to
// existing rows that were entered manually.
// Run from CMS Manager → Backfill Existing Rows.
// --------------------------------------------
function backfillExistingRows() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var ui       = SpreadsheetApp.getUi();
  var sheets   = ss.getSheets().filter(function(s) { return s.getName() !== 'Sheet1'; });
  var now      = new Date().toISOString();
  var patched  = 0;

  sheets.forEach(function(sheet) {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    for (var i = 1; i < data.length; i++) {
      var row     = data[i];
      var changed = false;

      // Skip completely empty rows (no title at all)
      if (!row[CONFIG.COLS.TITLE] && !row[CONFIG.COLS.ID]) continue;

      // Fill missing ID
      if (!row[CONFIG.COLS.ID]) {
        sheet.getRange(i + 1, CONFIG.COLS.ID + 1).setValue(Utilities.getUuid());
        changed = true;
      }

      // Fill missing Status
      if (!row[CONFIG.COLS.STATUS]) {
        sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(CONFIG.STATUS.TESTING);
        changed = true;
      }

      // Fill missing UPDATED_AT
      if (!row[CONFIG.COLS.UPDATED_AT]) {
        sheet.getRange(i + 1, CONFIG.COLS.UPDATED_AT + 1).setValue(now);
        changed = true;
      }

      if (changed) patched++;
    }
  });

  ui.alert(
    'Backfill Complete',
    'Rows updated with missing ID / Status / UPDATED_AT: ' + patched,
    ui.ButtonSet.OK
  );
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

    // Skip truly empty rows (no ID AND no Title)
    if (!row[CONFIG.COLS.ID] && !row[CONFIG.COLS.TITLE]) continue;

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

// =============================================
// SHEETSETUP.gs - Sheet tab management
// =============================================

// Standard headers for every tab
var SHEET_HEADERS = [
  'ID', 'Title', 'Event Type', 'Description',
  'Date', 'Time', 'Info', 'URL', 'Graphic', 'Status', 'UPDATED_AT'
];

// --------------------------------------------
// MIGRATE: Rename old-format sheet tabs to the
// new names and set all existing rows to 'deployed'
// so they appear in the preview and GitHub push.
//
// Old → New mapping:
//   Events  → JPAR Events
//   Home, Blog, Team, Gallery → skipped (no match)
//
// Safe: only renames if the old tab exists AND
// the new tab is empty (won't overwrite real data).
// --------------------------------------------
function migrateOldSheets() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var log = [];

  var MAP = {
    'Events': 'JPAR Events'
    // Add more mappings here if needed, e.g.:
    // 'TX Events': 'JPAR TX Events'
  };

  Object.keys(MAP).forEach(function(oldName) {
    var newName  = MAP[oldName];
    var oldSheet = ss.getSheetByName(oldName);
    var newSheet = ss.getSheetByName(newName);

    if (!oldSheet) {
      log.push('Skipped: "' + oldName + '" not found.');
      return;
    }

    if (newSheet) {
      var newRows = newSheet.getLastRow();
      if (newRows > 1) {
        log.push('Skipped: "' + newName + '" already has data (' + (newRows - 1) + ' rows).');
        return;
      }
      // New tab exists but is empty — delete it before renaming
      ss.deleteSheet(newSheet);
    }

    oldSheet.setName(newName);
    log.push('Renamed: "' + oldName + '" → "' + newName + '"');
  });

  // Fix headers on all required tabs
  Object.values(CONFIG.SHEETS).forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return;

    // Write correct headers (preserves data rows)
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#00326D')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  });

  // Backfill IDs, status (→ deployed), timestamps on all rows
  var now     = new Date().toISOString();
  var patched = 0;

  ss.getSheets().forEach(function(sheet) {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    for (var i = 1; i < data.length; i++) {
      var row     = data[i];
      var changed = false;

      if (!row[0] && !row[1]) continue; // fully empty row

      if (!row[CONFIG.COLS.ID]) {
        sheet.getRange(i + 1, CONFIG.COLS.ID + 1).setValue(Utilities.getUuid());
        changed = true;
      }
      // Normalise status to lowercase and default blanks to deployed
      var existingStatus = String(row[CONFIG.COLS.STATUS] || '').toLowerCase().trim();
      if (!existingStatus || existingStatus === CONFIG.STATUS.TESTING) {
        sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(CONFIG.STATUS.DEPLOYED);
        changed = true;
      }
      if (!row[CONFIG.COLS.UPDATED_AT]) {
        sheet.getRange(i + 1, CONFIG.COLS.UPDATED_AT + 1).setValue(now);
        changed = true;
      }

      if (changed) patched++;
    }
  });

  log.push('');
  log.push('Rows updated (IDs + deployed status): ' + patched);
  log.push('');
  log.push('Next: Deploy → Push Deployed to GitHub');

  ui.alert('Migration Complete', log.join('\n'), ui.ButtonSet.OK);
}

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

    var status = String(row[CONFIG.COLS.STATUS] || '').toLowerCase().trim();
    if (statusFilter && statusFilter !== 'all' && status !== statusFilter.toLowerCase()) continue;

    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    result.push(obj);
  }

  return result;
}

// --------------------------------------------
// DIAGNOSTIC: Shows exactly what the script
// sees in the spreadsheet. Run this to debug
// blank previews. Check the alert output.
// --------------------------------------------
function diagnoseCMS() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var out  = [];

  var sheets = ss.getSheets();
  out.push('=== ALL SHEET TABS (' + sheets.length + ') ===');
  sheets.forEach(function(s) {
    var rows = Math.max(0, s.getLastRow() - 1);
    out.push('  "' + s.getName() + '"  —  ' + rows + ' data row(s)');
  });

  out.push('');
  out.push('=== REQUIRED TABS ===');
  Object.values(CONFIG.SHEETS).forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      out.push('  MISSING: "' + tabName + '"');
      return;
    }

    var data    = sheet.getDataRange().getValues();
    var headers = data[0] || [];
    out.push('  FOUND: "' + tabName + '"');
    out.push('    Headers: ' + headers.join(' | '));
    out.push('    Total rows (incl header): ' + data.length);

    // Check first data row
    if (data.length > 1) {
      var r = data[1];
      out.push('    Row 2 ID: "' + r[CONFIG.COLS.ID] + '"');
      out.push('    Row 2 Title: "' + r[CONFIG.COLS.TITLE] + '"');
      out.push('    Row 2 Status: "' + r[CONFIG.COLS.STATUS] + '"');
    }

    // What getAllRows returns
    var allRows = getAllRows(tabName);
    out.push('    getAllRows() result: ' + allRows.length + ' row(s)');
  });

  out.push('');
  out.push('=== DEPLOYED PAYLOAD SUMMARY ===');
  var payload = buildDeployedPayload();
  Object.keys(payload.tabs).forEach(function(tab) {
    out.push('  "' + tab + '": ' + payload.tabs[tab].length + ' deployed row(s)');
  });

  var msg = out.join('\n');
  Logger.log(msg);
  ui.alert('CMS Diagnostic', msg, ui.ButtonSet.OK);
}

// --------------------------------------------
// DEBUG: Log the full deployed payload to console
// --------------------------------------------
function debugDeployedPayload() {
  var payload = buildDeployedPayload();
  Logger.log(JSON.stringify(payload, null, 2));
}

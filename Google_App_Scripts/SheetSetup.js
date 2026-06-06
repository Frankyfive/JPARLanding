// =============================================
// SHEETSETUP.gs - Sheet tab management
// =============================================

// Standard headers for every content tab
var SHEET_HEADERS = [
  'ID', 'Title', 'Event Type', 'Description',
  'Date', 'Time', 'Info', 'URL', 'Graphic', 'Status', 'UPDATED_AT'
];

// --------------------------------------------
// SORT: Sort a sheet's data rows by date column.
// Called from Datalayer.js after every append.
// Private helper (underscore suffix).
// --------------------------------------------
function sortSheetByDate_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  try {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .sort({ column: CONFIG.COLS.DATE + 1, ascending: true });
  } catch (e) {
    Logger.log('sortSheetByDate_ error on "' + sheet.getName() + '": ' + e.message);
  }
}

// --------------------------------------------
// EVENT TYPES TAB: Create or verify the
// "Event Types" tab with default values.
// Called from Setup Event Types menu item.
// --------------------------------------------
function setupEventTypesTab() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var name  = CONFIG.SHEETS.EVENT_TYPES;
  var sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1).setValue('Event Type');
    sheet.getRange(1, 1)
      .setFontWeight('bold')
      .setBackground('#00326D')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);

    var defaults = ['NETWORKING', 'TRAINING', 'CE', 'WEBINAR', 'WORKSHOP', 'ONLINE', 'IN-PERSON', 'OTHER'];
    defaults.forEach(function(t, i) { sheet.getRange(i + 2, 1).setValue(t); });
    sheet.setColumnWidth(1, 200);
    ui.alert('Event Types tab created with default types.\n\nAdd your own types directly to the sheet.');
  } else {
    ui.alert('"' + name + '" already exists with ' + (sheet.getLastRow() - 1) + ' type(s).');
  }
}

// --------------------------------------------
// ARCHIVE TRIGGER: Set up a daily time-based
// trigger that archives past deployed events.
// Run once from CMS Manager → Setup Auto-Archive.
// --------------------------------------------
function setupArchiveTrigger() {
  var ui = SpreadsheetApp.getUi();

  // Remove any existing archivePastEvents triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'archivePastEvents') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('archivePastEvents')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .inTimezone('America/Chicago')
    .create();

  ui.alert(
    'Auto-Archive Trigger Set',
    'Events will be automatically archived at 2:00 AM CT when their date has passed.\n\n' +
    'Digital Library entries are never auto-archived.',
    ui.ButtonSet.OK
  );
}

// --------------------------------------------
// ARCHIVE WORKER: Archives deployed events
// whose date has passed. Skip Digital Library
// and Event Types tabs. Called by the trigger.
// --------------------------------------------
function archivePastEvents() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var today    = new Date();
  today.setHours(0, 0, 0, 0);

  var skipSheets = [
    CONFIG.SHEETS.DIGITAL_LIBRARY,
    CONFIG.SHEETS.EVENT_TYPES,
    'Sheet1'
  ];

  var archived = 0;

  ss.getSheets().forEach(function(sheet) {
    if (skipSheets.indexOf(sheet.getName()) !== -1) return;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    for (var i = 1; i < data.length; i++) {
      var row     = data[i];
      var status  = String(row[CONFIG.COLS.STATUS] || '').toLowerCase().trim();
      var dateVal = row[CONFIG.COLS.DATE];

      if (status !== CONFIG.STATUS.DEPLOYED) continue;
      if (!dateVal) continue;

      var eventDate = new Date(dateVal);
      eventDate.setHours(0, 0, 0, 0);

      if (eventDate < today) {
        sheet.getRange(i + 1, CONFIG.COLS.STATUS     + 1).setValue(CONFIG.STATUS.ARCHIVE);
        sheet.getRange(i + 1, CONFIG.COLS.UPDATED_AT + 1).setValue(new Date().toISOString());
        archived++;
      }
    }
  });

  Logger.log('archivePastEvents: archived ' + archived + ' event(s).');
}

// --------------------------------------------
// MIGRATE: Rename old-format sheet tabs to the
// new names. Safe — only renames if old tab
// exists and new tab is empty.
// --------------------------------------------
function migrateOldSheets() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var log = [];

  var MAP = { 'Events': 'JPAR Events' };

  Object.keys(MAP).forEach(function(oldName) {
    var newName  = MAP[oldName];
    var oldSheet = ss.getSheetByName(oldName);
    var newSheet = ss.getSheetByName(newName);

    if (!oldSheet) { log.push('Skipped: "' + oldName + '" not found.'); return; }

    if (newSheet) {
      if (newSheet.getLastRow() > 1) {
        log.push('Skipped: "' + newName + '" already has data.'); return;
      }
      ss.deleteSheet(newSheet);
    }

    oldSheet.setName(newName);
    log.push('Renamed: "' + oldName + '" → "' + newName + '"');
  });

  // Fix headers on all required content tabs
  var contentTabs = [
    CONFIG.SHEETS.JPAR_EVENTS,
    CONFIG.SHEETS.JPAR_TX_EVENTS,
    CONFIG.SHEETS.DIGITAL_LIBRARY
  ];

  contentTabs.forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
      .setFontWeight('bold').setBackground('#00326D').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  });

  // Backfill IDs, lowercase status, timestamps
  var now     = new Date().toISOString();
  var patched = 0;

  ss.getSheets().forEach(function(sheet) {
    if (['Sheet1', CONFIG.SHEETS.EVENT_TYPES].indexOf(sheet.getName()) !== -1) return;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    for (var i = 1; i < data.length; i++) {
      var row     = data[i];
      var changed = false;
      if (!row[0] && !row[1]) continue;

      if (!row[CONFIG.COLS.ID]) {
        sheet.getRange(i + 1, CONFIG.COLS.ID + 1).setValue(Utilities.getUuid());
        changed = true;
      }
      var existingStatus = String(row[CONFIG.COLS.STATUS] || '').toLowerCase().trim();
      if (!existingStatus || existingStatus === CONFIG.STATUS.TESTING) {
        sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(CONFIG.STATUS.DEPLOYED);
        changed = true;
      } else if (existingStatus !== row[CONFIG.COLS.STATUS]) {
        sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(existingStatus);
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
  log.push('Rows updated: ' + patched);
  log.push('');
  log.push('Next: Deploy → Push Deployed to GitHub');
  ui.alert('Migration Complete', log.join('\n'), ui.ButtonSet.OK);
}

// --------------------------------------------
// SETUP: Ensure all required tabs exist with
// correct headers. Safe to run on existing data.
// --------------------------------------------
function setupSheets() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var log = [];

  // Content tabs
  var contentTabs = [
    CONFIG.SHEETS.JPAR_EVENTS,
    CONFIG.SHEETS.JPAR_TX_EVENTS,
    CONFIG.SHEETS.DIGITAL_LIBRARY
  ];

  contentTabs.forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);

    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      sheet.appendRow(SHEET_HEADERS);
      sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
        .setFontWeight('bold').setBackground('#00326D').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      log.push('Created: ' + tabName);
    } else {
      var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), SHEET_HEADERS.length)).getValues()[0];
      var needsFix = SHEET_HEADERS.some(function(h, i) { return existing[i] !== h; });

      if (needsFix) {
        sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
        sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
          .setFontWeight('bold').setBackground('#00326D').setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
        log.push('Fixed headers: ' + tabName);
      } else {
        log.push('OK: ' + tabName);
      }
    }
  });

  // Also ensure Event Types tab exists
  setupEventTypesTab();

  ui.alert(
    'Sheet Setup Complete',
    log.join('\n') + '\n\nReopen the spreadsheet to refresh the Add Content menu.',
    ui.ButtonSet.OK
  );
}

// --------------------------------------------
// ADD NEW TAB: Prompts for a name and creates
// a new sheet with standard content headers.
// --------------------------------------------
function addNewTab(tabName) {
  var ui = SpreadsheetApp.getUi();

  if (!tabName) {
    var response = ui.prompt('Add New Tab', 'Enter tab name:', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;
    tabName = response.getResponseText().trim();
  }

  if (!tabName) { ui.alert('Tab name cannot be empty.'); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(tabName)) { ui.alert('"' + tabName + '" already exists.'); return; }

  var sheet = ss.insertSheet(tabName);
  sheet.appendRow(SHEET_HEADERS);
  sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
    .setFontWeight('bold').setBackground('#00326D').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  ui.alert('Tab Created', '"' + tabName + '" is ready.\n\nReopen the spreadsheet so it appears in the Add Content menu.', ui.ButtonSet.OK);
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
// BACKFILL: Add missing ID, Status, UPDATED_AT
// to rows that were entered manually.
// --------------------------------------------
function backfillExistingRows() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var ui      = SpreadsheetApp.getUi();
  var now     = new Date().toISOString();
  var patched = 0;

  ss.getSheets().forEach(function(sheet) {
    if (['Sheet1', CONFIG.SHEETS.EVENT_TYPES].indexOf(sheet.getName()) !== -1) return;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    for (var i = 1; i < data.length; i++) {
      var row     = data[i];
      var changed = false;

      if (!row[CONFIG.COLS.TITLE] && !row[CONFIG.COLS.ID]) continue;

      if (!row[CONFIG.COLS.ID]) {
        sheet.getRange(i + 1, CONFIG.COLS.ID + 1).setValue(Utilities.getUuid());
        changed = true;
      }

      var rawStatus = String(row[CONFIG.COLS.STATUS] || '').trim();
      if (!rawStatus) {
        sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(CONFIG.STATUS.TESTING);
        changed = true;
      } else if (rawStatus !== rawStatus.toLowerCase()) {
        sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(rawStatus.toLowerCase());
        changed = true;
      }

      if (!row[CONFIG.COLS.UPDATED_AT]) {
        sheet.getRange(i + 1, CONFIG.COLS.UPDATED_AT + 1).setValue(now);
        changed = true;
      }

      if (changed) patched++;
    }
  });

  ui.alert('Backfill Complete', 'Rows updated: ' + patched, ui.ButtonSet.OK);
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
// DIAGNOSTIC: Shows what the script sees.
// Run to debug blank previews.
// --------------------------------------------
function diagnoseCMS() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var out = [];

  out.push('=== ALL SHEET TABS (' + ss.getSheets().length + ') ===');
  ss.getSheets().forEach(function(s) {
    out.push('  "' + s.getName() + '"  —  ' + Math.max(0, s.getLastRow() - 1) + ' data row(s)');
  });

  out.push('');
  out.push('=== REQUIRED CONTENT TABS ===');
  [CONFIG.SHEETS.JPAR_EVENTS, CONFIG.SHEETS.JPAR_TX_EVENTS, CONFIG.SHEETS.DIGITAL_LIBRARY].forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) { out.push('  MISSING: "' + tabName + '"'); return; }

    var data    = sheet.getDataRange().getValues();
    var headers = data[0] || [];
    out.push('  FOUND: "' + tabName + '"');
    out.push('    Headers: ' + headers.slice(0, 6).join(' | ') + '...');
    out.push('    Total rows: ' + data.length);
    if (data.length > 1) {
      out.push('    Row 2 Status: "' + data[1][CONFIG.COLS.STATUS] + '"');
    }
    out.push('    getAllRows(): ' + getAllRows(tabName).length + ' rows');
  });

  out.push('');
  out.push('=== DEPLOYED PAYLOAD SUMMARY ===');
  var payload = buildDeployedPayload();
  Object.keys(payload.tabs).forEach(function(tab) {
    out.push('  "' + tab + '": ' + payload.tabs[tab].length + ' deployed');
  });

  out.push('');
  out.push('=== EVENT TYPES ===');
  out.push('  Types: ' + getEventTypes().join(', '));

  var msg = out.join('\n');
  Logger.log(msg);
  ui.alert('CMS Diagnostic', msg, ui.ButtonSet.OK);
}

function debugDeployedPayload() {
  Logger.log(JSON.stringify(buildDeployedPayload(), null, 2));
}

// =============================================
// DATALAYER.gs - Read, write, filter, convert
// =============================================

// --------------------------------------------
// WRITE: Append a new row to a sheet tab
// --------------------------------------------
function appendRow(sheetName, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var now = new Date().toISOString();
  var row;

  if (sheetName === CONFIG.SHEETS.EVENTS) {
    row = [
      rowData.id          || Utilities.getUuid(),
      rowData.title       || '',
      rowData.eventType   || '',
      rowData.description || '',
      rowData.date        || '',
      rowData.time        || '',
      rowData.info        || '',
      rowData.url         || '',
      rowData.graphic     || '',
      rowData.status      || CONFIG.STATUS.TESTING,
      now
    ];
  } else {
    row = [
      rowData.id          || Utilities.getUuid(),
      rowData.title       || '',
      rowData.content     || '',
      rowData.imageUrl    || '',
      rowData.status      || CONFIG.STATUS.TESTING,
      rowData.date        || '',
      rowData.tags        || '',
      now
    ];
  }

  sheet.appendRow(row);
  return row[0];
}

// --------------------------------------------
// UPDATE: Update status of a row by ID
// --------------------------------------------
function updateRowStatus(sheetName, rowId, newStatus) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][CONFIG.COLS.ID] === rowId) {
      sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(newStatus);
      sheet.getRange(i + 1, CONFIG.COLS.UPDATED_AT + 1).setValue(new Date().toISOString());
      return true;
    }
  }

  throw new Error('Row not found with ID: ' + rowId);
}

// --------------------------------------------
// DELETE: Remove a row by ID (archive-safe)
// --------------------------------------------
function deleteRowById(sheetName, rowId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][CONFIG.COLS.ID] === rowId) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }

  throw new Error('Row not found with ID: ' + rowId);
}

// --------------------------------------------
// READ: Get rows from a tab filtered by status
// Reuses getSheetData() from SheetSetup.gs
// --------------------------------------------
function getDeployedRows(sheetName) {
  return getSheetData(sheetName, CONFIG.STATUS.DEPLOYED);
}

function getTestingRows(sheetName) {
  return getSheetData(sheetName, CONFIG.STATUS.TESTING);
}

function getAllRows(sheetName) {
  return getSheetData(sheetName, 'all');
}

// --------------------------------------------
// CONVERT: Build JSON payload from all tabs
// Only includes 'deployed' rows
// --------------------------------------------
function buildDeployedPayload() {
  var payload = {
    generatedAt: new Date().toISOString(),
    tabs: {}
  };

  Object.keys(CONFIG.SHEETS).forEach(function(key) {
    var sheetName = CONFIG.SHEETS[key];
    var rows      = getDeployedRows(sheetName);

    payload.tabs[sheetName] = rows.map(function(row) {
      return {
        id:        row['ID'],
        title:     row['Title'],
        content:   row['Content'],
        imageUrl:  row['Image URL'],
        date:      row['Date'],
        tags:      row['Tags'],
        updatedAt: row['Updated At']
      };
    });
  });

  return payload;
}

// --------------------------------------------
// VALIDATE: Check a row object has required fields
// --------------------------------------------
function validateRowData(rowData) {
  var errors = [];
  if (!rowData.title   || rowData.title.trim() === '')   errors.push('Title is required.');
  if (!rowData.content || rowData.content.trim() === '') errors.push('Content is required.');
  if (rowData.status && !Object.values(CONFIG.STATUS).includes(rowData.status)) {
    errors.push('Invalid status: ' + rowData.status);
  }
  return errors; // Empty array means valid
}
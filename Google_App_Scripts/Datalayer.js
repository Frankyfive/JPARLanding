// =============================================
// DATALAYER.gs - Read, write, filter, convert
// =============================================

// --------------------------------------------
// WRITE: Append a new row to any sheet tab.
// All tabs share the same column structure.
// --------------------------------------------
function appendRow(sheetName, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var now = new Date().toISOString();
  var row = [
    rowData.id          || Utilities.getUuid(),
    rowData.title       || '',
    rowData.eventType   || '',
    rowData.description || '',
    rowData.date        || '',
    rowData.time        || '',
    rowData.info        || '',
    rowData.url         || '',
    rowData.graphic     || '',
    String(rowData.status || CONFIG.STATUS.TESTING).toLowerCase(),
    now
  ];

  sheet.appendRow(row);
  return row[0]; // return the new row's ID
}

// --------------------------------------------
// UPDATE: Change the status of a row by ID
// --------------------------------------------
function updateRowStatus(sheetName, rowId, newStatus) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][CONFIG.COLS.ID] === rowId) {
      sheet.getRange(i + 1, CONFIG.COLS.STATUS     + 1).setValue(String(newStatus).toLowerCase());
      sheet.getRange(i + 1, CONFIG.COLS.UPDATED_AT + 1).setValue(new Date().toISOString());
      return true;
    }
  }

  throw new Error('Row not found with ID: ' + rowId);
}

// --------------------------------------------
// DELETE: Remove a row by ID
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
// READ: Convenience wrappers around getSheetData()
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
// CONVERT: Build JSON payload of all deployed rows
// across every sheet tab in the spreadsheet.
// --------------------------------------------
function buildDeployedPayload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetNames = ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return n !== 'Sheet1'; });

  var payload = {
    generatedAt: new Date().toISOString(),
    tabs: {}
  };

  sheetNames.forEach(function(sheetName) {
    var rows = getDeployedRows(sheetName);

    payload.tabs[sheetName] = rows.map(function(row) {
      return {
        id:          row['ID'],
        title:       row['Title'],
        eventType:   row['Event Type'],
        description: row['Description'],
        date:        row['Date'],
        time:        row['Time'],
        info:        row['Info'],
        url:         row['URL'],
        graphic:     row['Graphic'],
        updatedAt:   row['UPDATED_AT']
      };
    });
  });

  return payload;
}

// --------------------------------------------
// VALIDATE: Check required fields before saving
// --------------------------------------------
function validateRowData(rowData) {
  var errors = [];
  if (!rowData.title       || rowData.title.trim() === '')       errors.push('Title is required.');
  if (!rowData.description || rowData.description.trim() === '') errors.push('Description is required.');
  if (rowData.status && !Object.values(CONFIG.STATUS).includes(rowData.status)) {
    errors.push('Invalid status: ' + rowData.status);
  }
  return errors; // empty array = valid
}

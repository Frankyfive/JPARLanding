// =============================================
// DATALAYER.gs - Read, write, filter, convert
// =============================================

// Sheets excluded from content operations
var NON_CONTENT_SHEETS = ['Sheet1', CONFIG.SHEETS.EVENT_TYPES];

// --------------------------------------------
// WRITE: Append a new row to any content tab.
// Sorts the sheet by date after inserting.
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
  if (sheetName !== CONFIG.SHEETS.DIGITAL_LIBRARY) {
    sortSheetByDate_(sheet);
  }
  return row[0]; // return the new row's ID
}

// --------------------------------------------
// BATCH WRITE: Insert multiple rows at once.
// Used by createRecurringEvents to avoid slow
// single-row inserts. Sorts once at the end.
// --------------------------------------------
function appendRows_(sheetName, rowsData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var now  = new Date().toISOString();
  var ids  = [];
  var vals = rowsData.map(function(rowData) {
    var id = Utilities.getUuid();
    ids.push(id);
    return [
      id,
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
  });

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, vals.length, vals[0].length).setValues(vals);

  if (sheetName !== CONFIG.SHEETS.DIGITAL_LIBRARY) {
    sortSheetByDate_(sheet);
  }
  return ids;
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
// EVENT TYPES: Read/write the Event Types tab
// --------------------------------------------
function getEventTypes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.EVENT_TYPES);
  if (!sheet) return [];

  var data  = sheet.getDataRange().getValues();
  var types = [];
  for (var i = 1; i < data.length; i++) {
    var t = String(data[i][0] || '').trim().toUpperCase();
    if (t) types.push(t);
  }
  return types.sort();
}

function addEventType(typeName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.EVENT_TYPES);
  if (!sheet) {
    setupEventTypesTab();
    sheet = ss.getSheetByName(CONFIG.SHEETS.EVENT_TYPES);
  }

  typeName = String(typeName || '').trim().toUpperCase();
  if (!typeName) throw new Error('Event type name cannot be empty.');

  // Check for duplicate
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toUpperCase() === typeName) return typeName;
  }

  sheet.appendRow([typeName]);
  return typeName;
}

// --------------------------------------------
// RECURRING: Generate multiple dated events
// from a template rowData + recurrence config.
// recurrence = { type: 'weekly'|'biweekly'|'monthly',
//                endDate: 'YYYY-MM-DD',
//                days: [0-6] (0=Sun, for weekly) }
// --------------------------------------------
function createRecurringEvents(sheetName, rowData, recurrence) {
  if (!recurrence || !recurrence.type || recurrence.type === 'none') {
    var id = appendRow(sheetName, rowData);
    return [id];
  }

  var startDate = new Date(rowData.date + 'T00:00:00');
  var endDate   = new Date(recurrence.endDate + 'T23:59:59');

  if (isNaN(startDate.getTime())) throw new Error('Invalid start date.');
  if (isNaN(endDate.getTime()))   throw new Error('Invalid end date.');
  if (endDate < startDate)        throw new Error('End date must be after start date.');

  var dates = [];
  var d     = new Date(startDate);

  if (recurrence.type === 'weekly') {
    var days = recurrence.days || [];
    if (days.length === 0) throw new Error('Select at least one day for weekly recurrence.');
    while (d <= endDate) {
      if (days.indexOf(d.getDay()) !== -1) dates.push(new Date(d));
      d.setDate(d.getDate() + 1);
    }
  } else if (recurrence.type === 'biweekly') {
    while (d <= endDate) {
      dates.push(new Date(d));
      d.setDate(d.getDate() + 14);
    }
  } else if (recurrence.type === 'monthly') {
    while (d <= endDate) {
      dates.push(new Date(d));
      d.setMonth(d.getMonth() + 1);
    }
  }

  if (dates.length === 0) throw new Error('No occurrences found between start and end date.');
  if (dates.length > 52)  throw new Error('Too many occurrences (' + dates.length + '). Limit to 52 per batch.');

  var rowsData = dates.map(function(date) {
    var entry = {};
    Object.keys(rowData).forEach(function(k) { entry[k] = rowData[k]; });
    entry.id   = null;
    entry.date = Utilities.formatDate(date, 'America/Chicago', 'yyyy-MM-dd');
    return entry;
  });

  return appendRows_(sheetName, rowsData);
}

// --------------------------------------------
// CONVERT: Build JSON payload of all deployed
// rows across every content sheet tab.
// Sorted by date ascending (except Digital Library).
// --------------------------------------------
function buildDeployedPayload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetNames = ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return NON_CONTENT_SHEETS.indexOf(n) === -1; });

  var payload = {
    generatedAt: new Date().toISOString(),
    tabs: {}
  };

  sheetNames.forEach(function(sheetName) {
    var rows = getDeployedRows(sheetName);

    var mapped = rows.map(function(row) {
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

    // Sort by date ascending (Digital Library has no meaningful date order)
    if (sheetName !== CONFIG.SHEETS.DIGITAL_LIBRARY) {
      mapped.sort(function(a, b) {
        var da = a.date ? new Date(a.date).getTime() : Infinity;
        var db = b.date ? new Date(b.date).getTime() : Infinity;
        return da - db;
      });
    }

    payload.tabs[sheetName] = mapped;
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
  return errors;
}

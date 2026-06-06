// =============================================
// CALENDAR.gs - Google Calendar integration
// =============================================

// --------------------------------------------
// CREATE: Make a new public calendar for the CMS
// Run once from the menu then save the ID
// --------------------------------------------
function createPublicCalendar() {
  var ui = SpreadsheetApp.getUi();

  try {
    var existing = CONFIG.CALENDAR_ID
                 ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID)
                 : null;

    if (existing) {
      ui.alert(
        'Calendar Already Exists',
        'Calendar: ' + existing.getName()
        + '\nID: ' + existing.getId(),
        ui.ButtonSet.OK
      );
      return;
    }

    var calendar = CalendarApp.createCalendar(CONFIG.CALENDAR_NAME, {
      summary: 'Public events managed via the CMS.',
      timeZone: Session.getScriptTimeZone()
    });

    // Make it public so anyone can subscribe
    calendar.setSelected(true);

    var calId = calendar.getId();

    // Save to Script Properties for reuse
    PropertiesService.getScriptProperties().setProperty('CALENDAR_ID', calId);

    ui.alert(
      'Calendar Created',
      'Name: '    + calendar.getName()
      + '\nID: '  + calId
      + '\n\nSave this ID into CONFIG.CALENDAR_ID.\n'
      + 'Subscribe link:\nhttps://calendar.google.com/calendar?cid='
      + encodeURIComponent(calId),
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert('Error', err.message, ui.ButtonSet.OK);
    Logger.log('createPublicCalendar error: ' + err.message);
  }
}

// --------------------------------------------
// SYNC: Push all deployed Events rows to
// the calendar. Skips rows already synced.
// --------------------------------------------
function syncEventsToCalendar() {
  var ui = SpreadsheetApp.getUi();

  try {
    var calId = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID')
              || CONFIG.CALENDAR_ID;
    if (!calId) throw new Error('No Calendar ID found. Run "Create Public Calendar" first.');

    var calendar = CalendarApp.getCalendarById(calId);
    if (!calendar) throw new Error('Calendar not found for ID: ' + calId);

    var sheet    = SpreadsheetApp.getActiveSpreadsheet()
                   .getSheetByName(CONFIG.SHEETS.EVENTS);
    if (!sheet)  throw new Error('Events sheet not found.');

    var data     = sheet.getDataRange().getValues();
    var headers  = data[0];
    var synced   = 0;
    var skipped  = 0;

    for (var i = 1; i < data.length; i++) {
      var row    = data[i];
      var id     = row[CONFIG.COLS.ID];
      var status = row[CONFIG.COLS.STATUS];

      if (!id)                              { skipped++; continue; }
      if (status !== CONFIG.STATUS.DEPLOYED){ skipped++; continue; }

      // Build event object from row
      var title       = row[CONFIG.COLS.TITLE]   || 'Untitled Event';
      var description = row[CONFIG.COLS.CONTENT] || '';
      var imageUrl    = row[CONFIG.COLS.IMAGE_URL]|| '';
      var dateVal     = row[CONFIG.COLS.DATE];
      var tags        = row[CONFIG.COLS.TAGS]     || '';

      // Date column must be a valid date string or Date object
      var eventDate   = dateVal ? new Date(dateVal) : null;
      if (!eventDate || isNaN(eventDate.getTime())) { skipped++; continue; }

      // Check if this event ID was already synced
      // We store the GAS row ID in the event description as a marker
      var marker      = '[cms-id:' + id + ']';
      var existing    = calendar.getEventsForDay(eventDate).filter(function(e) {
        return e.getDescription().indexOf(marker) !== -1;
      });

      if (existing.length > 0) { skipped++; continue; } // Already synced

      // Build description with image and tags
      var fullDesc    = description;
      if (imageUrl)   fullDesc += '\n\nImage: ' + imageUrl;
      if (tags)       fullDesc += '\nTags: '    + tags;
      fullDesc       += '\n\n' + marker; // Append sync marker

      // Create an all-day event
      calendar.createAllDayEvent(title, eventDate, {
        description: fullDesc
      });

      synced++;
    }

    ui.alert(
      'Sync Complete',
      'Events synced: ' + synced + '\nSkipped: ' + skipped,
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert('Sync Error', err.message, ui.ButtonSet.OK);
    Logger.log('syncEventsToCalendar error: ' + err.message);
  }
}

// --------------------------------------------
// SUBSCRIBE: Show the public subscribe link
// --------------------------------------------
function showCalendarSubscribeLink() {
  var ui    = SpreadsheetApp.getUi();

  try {
    var calId = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID')
              || CONFIG.CALENDAR_ID;
    if (!calId) throw new Error('No Calendar ID found. Run "Create Public Calendar" first.');

    var calendar = CalendarApp.getCalendarById(calId);
    if (!calendar) throw new Error('Calendar not found for ID: ' + calId);

    var subscribeUrl = 'https://calendar.google.com/calendar?cid='
                     + encodeURIComponent(calId);
    var icalUrl      = 'https://calendar.google.com/calendar/ical/'
                     + encodeURIComponent(calId) + '/public/basic.ics';

    ui.alert(
      'Subscribe to ' + calendar.getName(),
      'Google Calendar link:\n' + subscribeUrl
      + '\n\niCal / Apple Calendar link:\n' + icalUrl
      + '\n\nShare either link with your team or embed it on your site.',
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert('Error', err.message, ui.ButtonSet.OK);
    Logger.log('showCalendarSubscribeLink error: ' + err.message);
  }
}

// --------------------------------------------
// DELETE: Remove a calendar event by CMS row ID
// Useful when a row is archived or deleted
// --------------------------------------------
function removeCalendarEventById(cmsRowId) {
  try {
    var calId    = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID')
                 || CONFIG.CALENDAR_ID;
    if (!calId)  return;

    var calendar = CalendarApp.getCalendarById(calId);
    var marker   = '[cms-id:' + cmsRowId + ']';
    var now      = new Date();
    var future   = new Date(now.getFullYear() + 5, 0, 1);
    var events   = calendar.getEvents(new Date(2000, 0, 1), future);

    events.forEach(function(e) {
      if (e.getDescription().indexOf(marker) !== -1) {
        e.deleteEvent();
      }
    });

  } catch (err) {
    Logger.log('removeCalendarEventById error: ' + err.message);
  }
}
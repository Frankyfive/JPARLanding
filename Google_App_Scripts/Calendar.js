// =============================================
// CALENDAR.gs - Google Calendar reference
// =============================================
// The public JPAR calendar is managed directly in Google Calendar.
// Events post automatically — no Apps Script sync required.
// The iCal URL is stored in CONFIG.CALENDAR_ICAL_URL and can be
// embedded or linked on any website.
// =============================================

function showCalendarInfo() {
  SpreadsheetApp.getUi().alert(
    'JPAR Public Calendar',
    'iCal / Subscribe URL:\n\n' + CONFIG.CALENDAR_ICAL_URL
    + '\n\nEmbed this URL in Google Calendar, Apple Calendar, or any website calendar widget.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

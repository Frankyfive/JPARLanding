// =============================================
// CONFIG.gs - Central configuration variables
// =============================================

var CONFIG = {

  // Google Drive
  DRIVE_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID'),

  // GitHub
  GITHUB_TOKEN:     PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN'),
  GITHUB_OWNER:     'Frankyfive',
  GITHUB_REPO:      'JPARLanding',
  GITHUB_BRANCH:    'main',
  GITHUB_FILE_PATH: 'data/cms.json',

  // Google Calendar (iCal subscription URL — events are managed directly in Google Calendar)
  CALENDAR_ICAL_URL: 'https://calendar.google.com/calendar/ical/c_feb05a8856474b2f2ba468abcadad5949329ae40bd70df2dfea770f7de1e51aa%40group.calendar.google.com/public/basic.ics',

  // Sheet tab names
  SHEETS: {
    JPAR_EVENTS:     'JPAR Events',
    JPAR_TX_EVENTS:  'JPAR TX Events',
    DIGITAL_LIBRARY: 'Digital Library'
  },

  // Status values
  STATUS: {
    ARCHIVE:  'archive',
    TESTING:  'testing',
    DEPLOYED: 'deployed'
  },

  // Column index map (0-based) — same for all tabs
  // Headers: ID | Title | Event Type | Description | Date | Time | Info | URL | Graphic | Status | UPDATED_AT
  COLS: {
    ID:          0,
    TITLE:       1,
    EVENT_TYPE:  2,
    DESCRIPTION: 3,
    DATE:        4,
    TIME:        5,
    INFO:        6,
    URL:         7,
    GRAPHIC:     8,
    STATUS:      9,
    UPDATED_AT:  10
  }
};

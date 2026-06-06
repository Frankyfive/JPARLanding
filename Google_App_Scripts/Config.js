// =============================================
// CONFIG.gs - Central configuration variables
// =============================================

var CONFIG = {

  // Google Drive
  DRIVE_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID'),

  // GitHub
  GITHUB_TOKEN: PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN'),
  GITHUB_OWNER:    'your-github-username',
  GITHUB_REPO:     'your-repo-name',
  GITHUB_BRANCH:   'main',
  GITHUB_FILE_PATH: 'data/cms.json',

  // Google Calendar
  CALENDAR_NAME:   'CMS Public Events',
  CALENDAR_ID:     '',  // Populated after first run of createPublicCalendar()

  // Sheet tab names (one per page/section)
  SHEETS: {
    HOME:     'Home',
    EVENTS:   'Events',
    BLOG:     'Blog',
    TEAM:     'Team',
    GALLERY:  'Gallery'
  },

  // Status values
  STATUS: {
    ARCHIVE:  'archive',
    TESTING:  'testing',
    DEPLOYED: 'deployed'
  },

  // Column index map (0-based) - shared across all tabs
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
// =============================================
// CODE.gs - Entry points and menu
// =============================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('CMS Manager')
    .addSubMenu(
      ui.createMenu('Add Content')
        .addItem('Home',    'showFormHome')
        .addItem('Events',  'showFormEvents')
        .addItem('Blog',    'showFormBlog')
        .addItem('Team',    'showFormTeam')
        .addItem('Gallery', 'showFormGallery')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Deploy')
        .addItem('Push Deployed to GitHub', 'deployToGithub')
        .addItem('Preview Testing Content', 'previewTesting')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Calendar')
        .addItem('Create Public Calendar',      'createPublicCalendar')
        .addItem('Sync Events Tab to Calendar', 'syncEventsToCalendar')
        .addItem('Get Subscribe Link',          'showCalendarSubscribeLink')
    )
    .addSeparator()
    .addItem('Setup Sheet Tabs', 'setupSheetTabs')
    .addToUi();
}

// Form launchers - one per tab
function showFormHome()    { showForm('Home');    }
function showFormEvents()  { showForm('Events');  }
function showFormBlog()    { showForm('Blog');    }
function showFormTeam()    { showForm('Team');    }
function showFormGallery() { showForm('Gallery'); }

// Generic form launcher - passes the target sheet name to the sidebar
function showForm(sheetName) {
  var template = HtmlService.createTemplateFromFile('Form');
  template.sheetName = sheetName;
  var html = template.evaluate()
    .setTitle('CMS: Add to ' + sheetName)
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Web app entry point for testing/preview
function doGet(e) {
  var page = (e && e.parameter && e.parameter.tab) ? e.parameter.tab : 'Home';
  var template = HtmlService.createTemplateFromFile('Preview');
  template.tab = page;
  return template.evaluate()
    .setTitle('CMS Preview - ' + page)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function previewTesting() {
  var template = HtmlService.createTemplateFromFile('Preview');
  template.tab = 'Events';
  var html = template.evaluate()
    .setTitle('CMS Preview - Testing')
    .setWidth(800)
    .setHeight(600)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, 'Preview: Testing Events');
}

// Include helper for HTML partials
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
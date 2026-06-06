// =============================================
// CODE.gs - Entry points and menu
// =============================================

function onOpen() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();

  // Build "Add Content" submenu dynamically from existing sheet tabs
  var sheetNames = ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return n !== 'Sheet1'; });

  // Cache names in Script Properties so numbered handlers can look them up
  var props = PropertiesService.getScriptProperties();
  sheetNames.forEach(function(name, i) {
    props.setProperty('_FORM_TAB_' + i, name);
  });
  props.setProperty('_FORM_TAB_COUNT', String(sheetNames.length));

  var addMenu = ui.createMenu('Add Content');
  sheetNames.slice(0, 10).forEach(function(name, i) {
    addMenu.addItem(name, 'showFormTab' + i);
  });

  ui.createMenu('CMS Manager')
    .addSubMenu(addMenu)
    .addSeparator()
    .addItem('Add New Tab',          'addNewTab')
    .addItem('Backfill Existing Rows', 'backfillExistingRows')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Deploy')
        .addItem('Push Deployed to GitHub', 'deployToGithub')
        .addItem('Preview Content',         'previewTesting')
    )
    .addToUi();
}

// ------------------------------------------------
// Dynamic tab form handlers (supports up to 10 tabs)
// Looks up the sheet name stored in onOpen()
// ------------------------------------------------
function showFormTab0() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_0')); }
function showFormTab1() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_1')); }
function showFormTab2() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_2')); }
function showFormTab3() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_3')); }
function showFormTab4() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_4')); }
function showFormTab5() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_5')); }
function showFormTab6() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_6')); }
function showFormTab7() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_7')); }
function showFormTab8() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_8')); }
function showFormTab9() { showForm(PropertiesService.getScriptProperties().getProperty('_FORM_TAB_9')); }

// Generic form launcher — opens the CMS sidebar for the given sheet tab
function showForm(sheetName) {
  if (!sheetName) return;
  var template = HtmlService.createTemplateFromFile('Form');
  template.sheetName = sheetName;
  var html = template.evaluate()
    .setTitle('CMS: Add to ' + sheetName)
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Web app entry point — used when deployed as a web app
function doGet(e) {
  var page = (e && e.parameter && e.parameter.tab) ? e.parameter.tab : 'JPAR Events';
  var template = HtmlService.createTemplateFromFile('Preview');
  template.tab = page;
  return template.evaluate()
    .setTitle('CMS Preview — ' + page)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Preview content in a modal dialog
function previewTesting() {
  var template = HtmlService.createTemplateFromFile('Preview');
  template.tab = 'JPAR Events';
  var html = template.evaluate()
    .setTitle('CMS Preview')
    .setWidth(1000)
    .setHeight(700)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, 'Preview: JPAR Events');
}

// Include helper for HTML partials
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

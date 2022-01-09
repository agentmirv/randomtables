/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Show sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  let t = HtmlService.createTemplateFromFile('Sidebar')

  let ui = t.evaluate()
    .setTitle('Random Tables 1.3');

  DocumentApp.getUi().showSidebar(ui);
}

function loadRandomTablesUrl() {
  let documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperty('randomTablesUrl') || "";
}

function saveRandomTablesUrl(url) {
  let documentProperties = PropertiesService.getDocumentProperties();
  if (url.trim() == "") {
    documentProperties.deleteProperty('randomTablesUrl');
  } else {
    documentProperties.setProperty('randomTablesUrl', url);
  }
}

function handleLoadButton(url) {
  let sections = [];

  try {
    let spreadsheet = SpreadsheetApp.openByUrl(url);
    saveRandomTablesUrl(url);
    sections = getSheetUrls(spreadsheet);
  } catch (error) {
    throw new Error(`Url [${url}] Error: ${error}`);
  }

  return sections;
}

function getSheetUrls(spreadsheet) {
  let sections = [];
  let urls = [];
  
  urls.push(spreadsheet.getUrl());

  let sheet = spreadsheet.getSheetByName('Links');

  if (sheet != null && sheet.getLastRow() > 1) {
    // Starting at Row 2, load URL string at Column 2
    let sheetUrls = sheet.getRange(2, 2, sheet.getLastRow() - 1).getValues();
    // Append each URL to a list of URLs
    urls = sheetUrls.reduce((accumulator, currentValue) => { 
      accumulator.push(currentValue[0]); 
      return accumulator; 
    }, urls);    
  }

  sections = urls.reduce((accumulator, currentValue) => { 
    accumulator.push({ url: currentValue, name: "", buttons: [], isLoaded: false, isMinimized: false }); 
    return accumulator; 
  }, sections);

  return sections;
}

function loadSection(url) {
  let section = { url: url, name: "", buttons: [], isLoaded: false, isMinimized: false };

  try {
    let spreadsheet = SpreadsheetApp.openByUrl(url);
    let sheet = spreadsheet.getSheetByName('Index');
    let name = spreadsheet.getName();
    section.name = name;
    
    if (sheet != null && sheet.getLastRow() > 1) {
      // Starting at row 2, load Button text at Column 1
      let values = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
      // Append each Button Text to the list of section buttons
      section.buttons = values.reduce((accumulator, currentValue) => { accumulator.push(currentValue[0]); return accumulator; }, section.buttons);
    }  
  } catch (error) {
    throw new Error(`Url [${url}] Error: ${error}`);
  }

  return section;
}


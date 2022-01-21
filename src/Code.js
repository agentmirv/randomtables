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

//=========================================================

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
      const startRow = 2;
      const startColumn = 1;
      const numRows = sheet.getLastRow() - 1;
      const numColumns = 3;
      let values = sheet.getRange(startRow, startColumn, numRows, numColumns).getValues();

      // Append each Button Text to the list of section buttons
      section.buttons = values.reduce((accumulator, currentValue) => { 
        accumulator.push({ 
          name: currentValue[0], 
          inputCount:  getInputCount(spreadsheet, currentValue[2]),
          inputRange: currentValue[2],
        }); 
        return accumulator; 
      }, section.buttons);
    }  
  } catch (error) {
    throw new Error(`Url [${url}] Error: ${error}`);
  }

  return section;
}

function getInputCount(spreadsheet, inputRangeA1) {
  var count = 0;

  if (inputRangeA1 != "") {
    const inputRange = spreadsheet.getRange(inputRangeA1);
    count = inputRange.getNumColumns();
  }

  return count;
}

//=========================================================

function showDialog(url, name, inputCount, inputRange) {
  if (inputCount > 0) {
    let t = HtmlService.createTemplateFromFile('Dialog');
    t.url = url;
    t.inputCount = inputCount;
    t.inputRange = inputRange;
    let rowHeight = 34;
    let buttonAreaHeight = 3 * rowHeight;
    let ui = t.evaluate()
      .setWidth(400)
      .setHeight((inputCount * rowHeight) + buttonAreaHeight);

    DocumentApp.getUi().showModalDialog(ui, name);
  } else {
    console.info({ name: "showDialog", inputCount: inputCount });
  }
}

function getDialogInput(url, inputRangeA1, index) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  var inputRange = null;

  try {
    inputRange = spreadsheet.getRange(inputRangeA1);
  } catch (error) {
    throw new Error(`Input [${inputRangeA1}]: ${error}`);
  }

  const row = 2;
  const column = index + 1;
  let inputCell = inputRange.getCell(row, column).getA1Notation();

  let inputRangeValues = inputRange.getValues();
  let description = inputRangeValues[0][index];
  let value = inputRangeValues[1][index];

  let hasOptions = false;
  let options = [];
  let inputRangeValidations = inputRange.getDataValidations();
  let valueValidations = inputRangeValidations[1][index];

  if (valueValidations != null) {
    let type = valueValidations.getCriteriaType();
    if (type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      hasOptions = true;
      let criteria = valueValidations.getCriteriaValues();
      options = criteria[0].filter(option => option.toString().trim().length > 0);
    } else if (type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      hasOptions = true;
      let criteria = valueValidations.getCriteriaValues();
      let criteriaRange = criteria[0];
      options = criteriaRange.getValues().flat().filter(option => option.toString().trim().length > 0);
    }
  }

  let input = {
    index: index,
    isLoaded: true,
    description: description,
    value: value,
    inputCell: inputCell,
    hasOptions: hasOptions,
    options: options, 
  };

  return input;
}

//=========================================================

 function handleSectionButtonClick(url, buttonText) {
  let selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    let ui = DocumentApp.getUi();
    ui.alert('Selection Detected', 'Deselect the selected text and try again.', ui.ButtonSet.OK);

  } else {
    console.info({ name: "handleSectionButtonClick", url: url, buttonText: buttonText });
    spreadsheetFunction(url, buttonText);
  }
}

function getIndexRow(spreadsheet, buttonText) {
  let indexSheet = spreadsheet.getSheetByName('Index');
  let indexValues = indexSheet.getRange(1, 1, indexSheet.getLastRow(), 3).getValues();
  let indexRow = {};
  
  [indexRow.buttonText, indexRow.outputRange, indexRow.inputRange] = indexValues.find(indexEntry => buttonText == indexEntry[0]);
  
  return indexRow; 
}

function spreadsheetFunction(url, buttonText) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let indexRow = getIndexRow(spreadsheet, buttonText);
  
  console.info({ name: "spreadsheetFunction", indexRow: indexRow });

  if (indexRow) {
    if (indexRow.inputRange != "") {
      // Show Input Dialog
      let dialogInputs = getDialogInputs(spreadsheet, indexRow.inputRange);
      showDialog(url, indexRow.buttonText, dialogInputs);
      
    } else {
      // Cycle random sheet functions
      let temp = spreadsheet.getRange('A1').getValue();
      spreadsheet.getRange('A1').setValue(temp);

      // Write Output to Doc
      addContentAtCursor(spreadsheet, indexRow.outputRange);
    }
  }
}

function getDialogInputs(spreadsheet, inputCell) {
  let dialogInputs = [];
  try {
    var inputRange = spreadsheet.getRange(inputCell);
  } catch (error) {
    throw new Error(`Input [${inputCell}]: ${error}`);
  }
  let inputRangeValues = inputRange.getValues();
  let inputRangeValidations = inputRange.getDataValidations();

  for (let j = 0; j < inputRange.getNumColumns(); j++) {
    let input = {
      input_help: inputRangeValues[0][j],
      default_value: inputRangeValues[1][j],
      // getCell column index starts at 1
      // This is the A1 Notation of the default_value cell
      input_range: inputRange.getCell(2, j + 1).getA1Notation()
    };

    // Process the input range valudations for the default_value cell
    // Set input.input_options
    if (inputRangeValidations[1][j] != null) {
      input.input_type = inputRangeValidations[1][j].getCriteriaType();
      if (input.input_type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        let criteria = inputRangeValidations[1][j].getCriteriaValues();
        input.input_options = criteria[0].filter(option => option.toString().trim().length > 0);
      } else if (input.input_type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
        let criteria = inputRangeValidations[1][j].getCriteriaValues();
        let criteriaRange = criteria[0];
        input.input_options = criteriaRange.getValues().flat().filter(option => option.toString().trim().length > 0);
      }
    }

    dialogInputs.push(input);
  }

  return dialogInputs;
}

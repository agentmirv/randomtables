/**
 * Random Tables
 * Marvin Sevilla
 */

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
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  let documentProperties = PropertiesService.getDocumentProperties();
  let randomTablesUrl = documentProperties.getProperty('randomTablesUrl') || "";

  let t = HtmlService.createTemplateFromFile('Sidebar')
  t.randomTablesUrl = randomTablesUrl;

  let ui = t.evaluate()
    .setTitle('Random Tables 1.1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Handles the load button click.
 *
 * @param {string} url The event of the spreadsheet url.
 * @returns {Object} The data used to draw the sidebar sections.
 */
function handleLoadButton(url) {
  let documentProperties = PropertiesService.getDocumentProperties();

  if (url.trim() == "") {
    documentProperties.deleteProperty('randomTablesUrl');
  } else {
    documentProperties.setProperty('randomTablesUrl', url);
  }

  return loadSpreadsheetUrl(url);
}

/**
 * Loads the spreadsheet from the url (Buttons sheet)
 * Also processes the Links sheet.
 *
 * @param {string} url The event of the spreadsheet url.
 * @returns {Object} The data used to draw the sidebar sections.
 */
function loadSpreadsheetUrl(url) {
  let data = { sections: [] };

  if (url.trim() != "") {
    let section = getButtonData(url);
    if (section != null) {
      data.sections.push(getButtonData(url));
    }

    let spreadsheet = SpreadsheetApp.openByUrl(url);
    let name = spreadsheet.getName();
    let sheet = spreadsheet.getSheetByName('Links');
    if (sheet != null && sheet.getLastRow() > 0) {
      let range = sheet.getRange(1, 2, sheet.getLastRow());
      let values = range.getValues();
      for (let i = 1; i < values.length; i++) {
        data.sections.push(getButtonData(values[i][0]));
      }
    }
  }

  return data;
}

/**
 * Loads the spreadsheet url and processes the Button sheet.
 *
 * @param {string} url The event of the spreadsheet url.
 * @returns {Object} The data used to draw a single sidebar section. 
 */
function getButtonData(url) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let name = spreadsheet.getName();
  let sheet = spreadsheet.getSheetByName('Index');
  let buttons = [];
  
  if (sheet != null && sheet.getLastRow() > 0) {
    let range = sheet.getRange(1, 1, sheet.getLastRow());
    let values = range.getValues();
    for (let i = 1; i < values.length; i++) {
      buttons.push(values[i][0]);
    }
  }

  return { url: url, name: name, buttons: buttons };
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 * 
 * @param {string} url The event of the spreadsheet url.
 * @param {string} function_name The function name, the first column of the Index sheet.
 * @param {Object} inputs The object representing the inputs to be rendered on the Dialog.
 */
function showDialog(url, function_name, inputs) {
  let t = HtmlService.createTemplateFromFile('Dialog');
  t.url = url;
  t.function_name = function_name;
  t.inputs = inputs;

  let rowHeight = 34;
  let ui = t.evaluate()
    .setWidth(400)
    .setHeight((inputs.length * rowHeight) + (3 * rowHeight))
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  DocumentApp.getUi().showModalDialog(ui, function_name);
}

/**
 * Adds content at the cursor position. 
 * 
 * @param {string} content The text to be written at the cursor position.
 */
function addAtCursor(content) {
  let cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    let element = cursor.insertText(content);
    let parent = element.getParent();
    let elementIndex = parent.getChildIndex(element);
    let cursorNew = DocumentApp.getActiveDocument().newPosition(parent, elementIndex + 1);
    DocumentApp.getActiveDocument().setCursor(cursorNew);
  }
}

/**
 * Handles the sidebar button click.
 */
function handleButtonClick(url, function_name) {
  let selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    let ui = DocumentApp.getUi();
    ui.alert('Selection Detected', 'Deselect the selected text and try again.', ui.ButtonSet.OK);
  } else {
    spreadsheetFunction(url, function_name);
  }
}

/**
 * Processes the function in the spreadsheet and 
 * shows the Dialog or processes the output.
 * 
 * @param {string} url The event of the spreadsheet url.
 * @param {string} function_name The function name, the first column of the Index sheet.
 */
function spreadsheetFunction(url, function_name) {
  let ui = DocumentApp.getUi();
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let indexSheet = spreadsheet.getSheetByName('Index');
  let indexRange = indexSheet.getRange(1, 1, indexSheet.getLastRow(), 3);
  let indexValues = indexRange.getValues();

  for (let i = 0; i < indexValues.length; i++) {
    if (function_name == indexValues[i][0]) {
      let output_cell = indexValues[i][1];
      let input_cell = indexValues[i][2];

      if (input_cell != "") {
        // Get inputs 
        let inputs = [];
        let inputRange = spreadsheet.getRange(input_cell);
        let inputRangeValues = inputRange.getValues();
        let inputRangeValidations = inputRange.getDataValidations();

        for (let j = 0; j < inputRangeValues[0].length; j++) {
          let input = {
            input_help: inputRangeValues[0][j],
            default_value: inputRangeValues[1][j],
            input_range: inputRange.getCell(2, j + 1).getA1Notation()
          };

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

          inputs.push(input);
        }

        showDialog(url, function_name, inputs);
      } else {
        // Cycle random sheet functions
        let temp = spreadsheet.getRange('A1').getValue();
        spreadsheet.getRange('A1').setValue(temp);

        let output = spreadsheet.getRange(output_cell).getValue();
        addAtCursor(`${output}\n`);
      }

      break;
    }
  }
}

/**
 * Processes the submit of the Dialog, calling the spreadsheet function
 * and processes the output.
 * 
 * @param {Object} data The object representing data submitted from the Dialog form. 
 */
function submitDialog(data) {
  let ui = DocumentApp.getUi();
  let spreadsheet = SpreadsheetApp.openByUrl(data.url);
  let indexSheet = spreadsheet.getSheetByName('Index');
  let indexRange = indexSheet.getRange(1, 1, indexSheet.getLastRow(), 3);
  let indexValues = indexRange.getValues();

  for (let i = 0; i < indexValues.length; i++) {
    if (data.function_name == indexValues[i][0]) {
      let output_cell = indexValues[i][1];
      let input_cell = indexValues[i][2];
      let inputRange = spreadsheet.getRange(input_cell);
      let inputRangeValues = inputRange.getValues();

      for (let j = 0; j < inputRangeValues[0].length; j++) {
        let cell = inputRange.getCell(2, j + 1);
        let key = cell.getA1Notation();
        if (key in data.input) {
          cell.setValue(data.input[key])
        }
      }

      let output = spreadsheet.getRange(output_cell).getValue();
      addAtCursor(`${output}\n`);
    }
  }
}

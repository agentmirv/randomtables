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
 * 
 */
function showSidebar() {
  let documentProperties = PropertiesService.getDocumentProperties();
  let randomTablesUrl = documentProperties.getProperty('randomTablesUrl') || "";

  let t = HtmlService.createTemplateFromFile('Sidebar')
  t.randomTablesUrl = randomTablesUrl;

  let ui = t.evaluate()
    .setTitle('Random Tables')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * 
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
 * 
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
    if (sheet != null) {
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
 * 
 */
function getButtonData(url) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let name = spreadsheet.getName();
  let sheet = spreadsheet.getSheetByName('Index');
  if (sheet != null) {
    let range = sheet.getRange(1, 1, sheet.getLastRow());
    let values = range.getValues();
    let buttons = [];
    for (let i = 1; i < values.length; i++) {
      buttons.push(values[i][0]);
    }
    return { url: url, name: name, buttons: buttons };
  }
}

/**
 * 
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
 * Adds content at the cursor location. 
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
 * 
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
              input.input_options = criteria[0];
            } else if (input.input_type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
              let criteria = inputRangeValidations[1][j].getCriteriaValues();
              let criteriaRange = criteria[0];
              input.input_options = criteriaRange.getValues().flat();
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
 * 
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

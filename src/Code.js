/**
 * @REMOnlyCurrentDoc  Limits the script to only accessing the current document.
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
 * 
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Random Tables')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * 
 */
function loadSpreadsheetUrl(url) {
  var data = { sections: [] };
  var section = getButtonData(url);
  if (section != null) {
  data.sections.push(getButtonData(url));
  }

  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var name = spreadsheet.getName();
  var sheet = spreadsheet.getSheetByName('Links');
  if (sheet != null) {
  var range = sheet.getRange(1, 2, sheet.getLastRow());
  var values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    data.sections.push(getButtonData(values[i][0]));
  }
  }  
  return data;
}

/**
 * 
 */
function getButtonData(url) {
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var name = spreadsheet.getName();
  var sheet = spreadsheet.getSheetByName('Index');
  if (sheet != null)
  {
  var range = sheet.getRange(1, 1, sheet.getLastRow());
  var values = range.getValues();
  var buttons = [];
  for (var i = 1; i < values.length; i++) {
    buttons.push(values[i][0]);
  }
  return { url: url, name: name, buttons: buttons };
  }
}

/**
 * 
 */
function showDialog(url, function_name, inputs) {
  var t = HtmlService.createTemplateFromFile('Dialog');
  t.url = url;
  t.function_name = function_name;
  t.inputs = inputs;

  var rowHeight = 34;
  var ui = t.evaluate()
    .setWidth(400)
    .setHeight((inputs.length * rowHeight) + (3 * rowHeight))
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  DocumentApp.getUi().showModalDialog(ui, function_name);
}

/**
 * Adds content at the cursor location. 
 */
function addAtCursor(content) {
  var cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    var element = cursor.insertText(content);
    var parent = element.getParent();
    var elementIndex = parent.getChildIndex(element);
    var cursorNew = DocumentApp.getActiveDocument().newPosition(parent, elementIndex + 1);
    DocumentApp.getActiveDocument().setCursor(cursorNew);
  }
}

/**
 * 
 */
function spreadsheetFunction(url, function_name) {
  var ui = DocumentApp.getUi();
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var indexSheet = spreadsheet.getSheetByName('Index');
  var indexRange = indexSheet.getRange(1, 1, indexSheet.getLastRow(), 3);
  var indexValues = indexRange.getValues();

  for (var i = 0; i < indexValues.length; i++) {
    if (function_name == indexValues[i][0]) {
      var output_cell = indexValues[i][1];
      var input_cell = indexValues[i][2];

      if (input_cell != "") {
        // Get inputs 
        var inputs = [];
        var inputRange = spreadsheet.getRange(input_cell);
        var inputRangeValues = inputRange.getValues();
        var inputRangeValidations = inputRange.getDataValidations();

        for (var j = 0; j < inputRangeValues[0].length; j++) {
          var input = {
            input_help: inputRangeValues[0][j],
            default_value: inputRangeValues[1][j],
            input_range: inputRange.getCell(2, j + 1).getA1Notation()
          };

          if (inputRangeValidations[1][j] != null) {
            input.input_type = inputRangeValidations[1][j].getCriteriaType();
            if (input.input_type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
              var criteria = inputRangeValidations[1][j].getCriteriaValues();
              input.input_options = criteria[0];
            } else if (input.input_type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
              var criteria = inputRangeValidations[1][j].getCriteriaValues();
              var criteriaRange = criteria[0];
              input.input_options = criteriaRange.getValues().flat();
            }
          }
          
          inputs.push(input);
        }

        showDialog(url, function_name, inputs);
      } else {
        // Cycle random sheet functions
        var temp = spreadsheet.getRange('A1').getValue();
        spreadsheet.getRange('A1').setValue(temp);

        var output = spreadsheet.getRange(output_cell).getValue();
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
  var ui = DocumentApp.getUi();
  var spreadsheet = SpreadsheetApp.openByUrl(data.url);
  var indexSheet = spreadsheet.getSheetByName('Index');
  var indexRange = indexSheet.getRange(1, 1, indexSheet.getLastRow(), 3);
  var indexValues = indexRange.getValues();

  for (var i = 0; i < indexValues.length; i++) {
    if (data.function_name == indexValues[i][0]) {
      var output_cell = indexValues[i][1];
      var input_cell = indexValues[i][2];
      var inputRange = spreadsheet.getRange(input_cell);
      var inputRangeValues = inputRange.getValues();

      for (var j = 0; j < inputRangeValues[0].length; j++) {
        var cell = inputRange.getCell(2, j + 1);
        var key = cell.getA1Notation();
        if (key in data.input) {
          cell.setValue(data.input[key])
        }
      }

      var output = spreadsheet.getRange(output_cell).getValue();
      addAtCursor(`${output}\n`);
    }
  }
}

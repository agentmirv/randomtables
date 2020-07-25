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
  // Get randomTablesUrl from DocumentProperties
  let documentProperties = PropertiesService.getDocumentProperties();
  let randomTablesUrl = documentProperties.getProperty('randomTablesUrl') || "";

  // Load Sidebar
  let t = HtmlService.createTemplateFromFile('Sidebar')
  t.randomTablesUrl = randomTablesUrl;

  let ui = t.evaluate()
    .setTitle('Random Tables 2.0')
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
  // Update randomTablesUrl DocumentProperties
  let documentProperties = PropertiesService.getDocumentProperties();

  if (url.trim() == "") {
    documentProperties.deleteProperty('randomTablesUrl');
  } else {
    documentProperties.setProperty('randomTablesUrl', url);
  }

  // Return the sidebar data
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
  // Build the sidebar data
  let data = { sections: [] };

  if (url.trim() != "") {
    // Load spreadsheet URLs from the Links sheet
    let loadUrls = loadLinksSheet(url);
    // Prepend the original spreadsheet URL
    loadUrls.unshift(url);

    // Load each spreadsheet URL and process Index sheet into a sidebar section
    loadUrls.forEach(loadUrl => {
      let section = loadIndexSheet(loadUrl);
      data.sections.push(section);
    });
  }

  return data;
}

/**
 * Loads the spreadsheet url and processes the Links sheet.
 *
 * @param {string} url The event of the spreadsheet url.
 * @returns {Array} The additional spreadsheet urls to load. 
 */
function loadLinksSheet(url) {
  let urls = [];
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let sheet = spreadsheet.getSheetByName('Links');

  if (sheet != null && sheet.getLastRow() > 0) {
    // Load URL string at Column 2
    let range = sheet.getRange(1, 2, sheet.getLastRow());
    let values = range.getValues();
    for (let i = 1; i < values.length; i++) {
      let linksUrl = values[i][0];
      urls.push(linksUrl);
    }
  }

  return urls;
}

/**
 * Loads the spreadsheet url and processes the Index sheet.
 *
 * @param {string} url The event of the spreadsheet url.
 * @returns {Object} The data used to draw a single sidebar section. 
 */
function loadIndexSheet(url) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let name = spreadsheet.getName();
  let section = { url: url, name: name, buttons: [] };

  let sheet = spreadsheet.getSheetByName('Index');
  if (sheet != null && sheet.getLastRow() > 0) {
    // Load Button text at Column 1
    let range = sheet.getRange(1, 1, sheet.getLastRow());
    let values = range.getValues();
    for (let i = 1; i < values.length; i++) {
      let buttonText = values[i][0];
      section.buttons.push(buttonText);
    }
  }

  return section;
}

/**
 * Handles the section button click.
 * 
 * @param {string} url The event of the spreadsheet url. 
 * @param {string} functionName The button text in the Index sheet. 
 */
function handleSectionButtonClick(url, functionName) {
  let selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    let ui = DocumentApp.getUi();
    ui.alert('Selection Detected', 'Deselect the selected text and try again.', ui.ButtonSet.OK);
  } else {
    spreadsheetFunction(url, functionName);
  }
}

/**
 * Processes the function in the spreadsheet and 
 * shows the Dialog or processes the output.
 * 
 * @param {string} url The event of the spreadsheet url.
 * @param {string} functionName The function name, the first column of the Index sheet.
 */
function spreadsheetFunction(url, functionName) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let indexSheet = spreadsheet.getSheetByName('Index');
  let indexValues = indexSheet.getRange(1, 1, indexSheet.getLastRow(), 3).getValues();

  let selectedFunction = indexValues.filter(indexEntry => functionName == indexEntry[0]);

  if (selectedFunction.length > 0) {
    let [, output_cell, input_cell] = selectedFunction[0];
    
    if (input_cell != "") {
      // Show Input Dialog
      let dialogInputs = getDialogInputs(spreadsheet, input_cell);
      showDialog(url, functionName, dialogInputs);

    } else {
      // Cycle random sheet functions
      let temp = spreadsheet.getRange('A1').getValue();
      spreadsheet.getRange('A1').setValue(temp);

      // Write Output to Doc
      addContentAtCursor(spreadsheet, output_cell);
    }
  }
}

/**
]* Gets the dialog input data for the function input range
 * 
 * @returns {Array} The data used to draw a list of dialog inputs.  
 */
function getDialogInputs(spreadsheet, inputCell){
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

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 * 
 * @param {string} url The event of the spreadsheet url.
 * @param {string} functionName The function name, the first column of the Index sheet.
 * @param {Object} inputs The object representing the inputs to be rendered on the Dialog.
 */
function showDialog(url, functionName, inputs) {
  let t = HtmlService.createTemplateFromFile('Dialog');
  t.url = url;
  t.function_name = functionName;
  t.inputs = inputs;

  let rowHeight = 34;
  let ui = t.evaluate()
    .setWidth(400)
    .setHeight((inputs.length * rowHeight) + (3 * rowHeight))
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  DocumentApp.getUi().showModalDialog(ui, functionName);
}

/**
 * Processes the submit of the Dialog, calling the spreadsheet function
 * and processes the output.
 * 
 * @param {Object} data The object representing data submitted from the Dialog form. 
 */
function submitDialog(data) {
  let spreadsheet = SpreadsheetApp.openByUrl(data.url);
  let indexSheet = spreadsheet.getSheetByName('Index');
  let indexValues = indexSheet.getRange(1, 1, indexSheet.getLastRow(), 3).getValues();

  let selectedFunction = indexValues.filter(indexEntry => data.function_name == indexEntry[0]);

  if (selectedFunction.length > 0) {
    let [, output_cell, input_cell] = selectedFunction[0];  

    try {
      var inputRange = spreadsheet.getRange(input_cell);
    } catch (error) {
      throw new Error(`Input [${inputCell}]: ${error}`);
    }

    // Set the Inputs
    for (let j = 0; j < inputRange.getNumColumns(); j++) {
        // This is the cell containing the default_value
        let cell = inputRange.getCell(2, j + 1);
        let key = cell.getA1Notation();
        if (key in data.input) {
          cell.setValue(data.input[key])
        }
    }

    // Write Output to Doc
    addContentAtCursor(spreadsheet, output_cell);
  }
}

/**
 * Writes content at the cursor position. 
 * 
 * @param {Object} spreadsheet The spreadsheet containing the output range. 
 * @param {string} outputCell The output cell in A1 notation.
 */
function addContentAtCursor(spreadsheet, outputCell) {
  let document = DocumentApp.getActiveDocument();
  let cursor = document.getCursor();
  let elementInserted = null;

  if (cursor) {
    try {
      let outputRange = spreadsheet.getRange(outputCell);

      // Output Range detected
      if (outputRange.getNumRows() == 2) {
        let [outputType, outputValue ] = outputRange.getValues();
        if (outputType == 'imageurl') {
          try {
            let outputImageBlob = UrlFetchApp.fetch(outputValue).getBlob();
            elementInserted = cursor.insertInlineImage(outputImageBlob);
          } catch (error) {
            throw new Error(`ImageURL [${outputValue}]: ${error}`);
          }
        }      
      }
      
      // Output Cell detected or invalid Output Range type
      if (!elementInserted) {
        let outputValue = outputRange.getValue();
        elementInserted = cursor.insertText(outputValue);
      }
      
      // Move cursor accordingly
      let parent = elementInserted.getParent();
      let elementIndex = parent.getChildIndex(elementInserted);
      let cursorNew = document.newPosition(parent, elementIndex + 1);
      document.setCursor(cursorNew);

    } catch (error) {
      throw new Error(`Output [${outputCell}]: ${error}`);
    }  
  }
}


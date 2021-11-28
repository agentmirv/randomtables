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
  let randomTablesUrl = loadRandomTablesUrl();

  console.info({ name: "showSidebar", randomTablesUrl: randomTablesUrl });

  // Load Sidebar
  let t = HtmlService.createTemplateFromFile('Sidebar')
  t.randomTablesUrl = randomTablesUrl;

  let ui = t.evaluate()
    .setTitle('Random Tables 1.2');

  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Handles the load button click.
 *
 * @param {string} url The spreadsheet url.
 * @returns {Object} The data used to draw the sidebar sections.
 */
function handleLoadButton(url) {
  // Update randomTablesUrl DocumentProperties
  saveRandomTablesUrl(url)

  // Return the sidebar data
  return loadSpreadsheetUrl(url);
}

/**
 * Saves the url to the DocumentProperties
 *
 * @param {string} url The spreadsheet url.
 */
function saveRandomTablesUrl(url) {
  let documentProperties = PropertiesService.getDocumentProperties();
  if (url.trim() == "") {
    documentProperties.deleteProperty('randomTablesUrl');
  } else {
    documentProperties.setProperty('randomTablesUrl', url);
  }
}

/**
 * Saves the url to the DocumentProperties
 *
 * @returns {Object} The spreadsheet url.
 */
function loadRandomTablesUrl() {
  let documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperty('randomTablesUrl') || "";
}

/**
 * Loads the spreadsheet from the url (Buttons sheet)
 * Also processes the Links sheet.
 *
 * @param {string} url The spreadsheet url.
 * @returns {Object} The data used to draw the sidebar sections.
 */
function loadSpreadsheetUrl(url) {
  console.info({ name: "loadSpreadsheetUrl", url: url });
  
  // Build the sidebar data
  let data = { sections: [] };

  if (url.trim() != "") {
    // Starting with the original url, append the urls on the Links sheet
    let loadUrls = loadLinksSheet(url);

    console.info({ name: "loadSpreadsheetUrl", loadUrls: loadUrls });

    // Load each spreadsheet URL and process Index sheet into a sidebar section
    data.sections = loadUrls.reduce((accumulator, currentValue) => { accumulator.push(loadIndexSheet(currentValue)); return accumulator; }, data.sections);
  }

  return data;
}

/**
 * Loads the spreadsheet url and processes the Links sheet.
 *
 * @param {string} url The spreadsheet url.
 * @returns {Array} The additional spreadsheet urls to load. 
 */
function loadLinksSheet(url) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let sheet = spreadsheet.getSheetByName('Links');
  let urls = [url];

  if (sheet != null && sheet.getLastRow() > 1) {
    // Starting at Row 2, load URL string at Column 2
    let newValues = sheet.getRange(2, 2, sheet.getLastRow() - 1).getValues();
    // Append each URL to a list of URLs
    urls = newValues.reduce((accumulator, currentValue) => { accumulator.push(currentValue[0]); return accumulator; }, urls);

    console.info({ name: "loadLinksSheet", urls: urls });
  }

  return urls;
}

/**
 * Loads the spreadsheet url and processes the Index sheet.
 *
 * @param {string} url The spreadsheet url.
 * @returns {Object} The data used to draw a single sidebar section. 
 */
function loadIndexSheet(url) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  let sheet = spreadsheet.getSheetByName('Index');
  let name = spreadsheet.getName();
  let section = { url: url, name: name, buttons: [] };
  
  if (sheet != null && sheet.getLastRow() > 1) {
    // Starting at row 2, load Button text at Column 1
    let values = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
    // Append each Button Text to the list of section buttons
    section.buttons = values.reduce((accumulator, currentValue) => { accumulator.push(currentValue[0]); return accumulator; }, section.buttons);

    console.info({ name: "loadIndexSheet", url: url, buttons: section.buttons });
  }

  return section;
}

/**
 * Handles the section button click.
 * 
 * @param {string} url The spreadsheet url. 
 * @param {string} buttonText The button text in the Index sheet. 
 */
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

/**
 * Gets data representing the Index sheet row with a match on Button Text
 * 
 * @param {Object} spreadsheet The spreadsheet containing the input range.
 * @param {string} buttonText The button text in the Index sheet. 
 * @returns {Object} The data representing the Index sheet row.  
 */
function getIndexRow(spreadsheet, buttonText) {
  let indexSheet = spreadsheet.getSheetByName('Index');
  let indexValues = indexSheet.getRange(1, 1, indexSheet.getLastRow(), 3).getValues();
  let indexRow = {};
  
  [indexRow.buttonText, indexRow.outputRange, indexRow.inputRange] = indexValues.find(indexEntry => buttonText == indexEntry[0]);
  
  return indexRow; 
}

/**
 * Processes the function in the spreadsheet and 
 * shows the Dialog or processes the output.
 * 
 * @param {string} url The spreadsheet url.
 * @param {string} buttonText The button text in the Index sheet. 
 */
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

/**
 * Gets the dialog input data for the function input range
 * 
 * @param {Object} spreadsheet The spreadsheet containing the input range.
 * @param {string} inputCell The input cell in A1 notation.
 * @returns {Array} The list of dialog input data.  
 */
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

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 * 
 * @param {string} url The spreadsheet url.
 * @param {string} buttonText The button text in the Index sheet. 
 * @param {Object} inputs The object representing the inputs to be rendered on the Dialog.
 */
function showDialog(url, buttonText, inputs) {
  let t = HtmlService.createTemplateFromFile('Dialog');
  t.url = url;
  t.function_name = buttonText;
  t.inputs = inputs;

  console.info({ name: "showDialog", url: t.url, function_name: t.function_name, inputs: t.inputs });

  let rowHeight = 34;
  let ui = t.evaluate()
    .setWidth(400)
    .setHeight((inputs.length * rowHeight) + (3 * rowHeight));

  DocumentApp.getUi().showModalDialog(ui, buttonText);
}

/**
 * Processes the submit of the Dialog, calling the spreadsheet function
 * and processes the output.
 * 
 * @param {Object} data The object representing data submitted from the Dialog form. 
 */
function submitDialog(data) {
  let spreadsheet = SpreadsheetApp.openByUrl(data.url);
  let indexRow = getIndexRow(spreadsheet, data.function_name);
  
  console.info({ name: "submitDialog", data: data });
  
  if (indexRow) {
    try {
      var inputRange = spreadsheet.getRange(indexRow.inputRange);
    } catch (error) {
      throw new Error(`Input [${indexRow.inputRange}]: ${error}`);
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
    addContentAtCursor(spreadsheet, indexRow.outputRange);
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
        // Support only one output column
        outputRange = outputRange.offset(0, 0, 2, 1);
        let [outputType, outputValue] = outputRange.getValues();
        if (outputType == 'imageurl') {
          try {
            let outputImageBlob = UrlFetchApp.fetch(outputValue).getBlob();
            elementInserted = cursor.insertInlineImage(outputImageBlob);
            console.info({ name: "addContentAtCursor", outputValue: outputValue });
          } catch (error) {
            throw new Error(`ImageURL [${outputValue}]: ${error}`);
          }
        }
      }

      // Output Cell detected or invalid Output Range type
      if (!elementInserted) {
        let outputValue = outputRange.getValue();
        elementInserted = cursor.insertText(outputValue);
        console.info({ name: "addContentAtCursor", outputValue: outputValue });
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


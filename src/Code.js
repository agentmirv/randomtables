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

function handleInitialize() {
  return loadRandomTablesUrl_();
}

function handleLoadButton(url) {
  let sections = [];

  try {
    let spreadsheet = SpreadsheetApp.openByUrl(url);
    saveRandomTablesUrl_(url);
    sections = getSheetUrls_(spreadsheet);
  } catch (error) {
    throw new Error(`Url [${url}] Error: ${error}`);
  }

  return sections;
}

function handleLoadSection(url) {
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
          outputRange: currentValue[1], 
          inputCount:  getInputCount_(spreadsheet, currentValue[2]),
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

function handleAction(url, name, outputRange, inputCount, inputRange) {
  console.info(url, name, outputRange, inputCount, inputRange);
  let documentHasSelection = documentHasSelection_();
  if (!documentHasSelection) {
    if (inputCount > 0) {
      showDialog_(url, name, outputRange, inputCount, inputRange);
    } else {
      let spreadsheet = SpreadsheetApp.openByUrl(url);
      // Cycle random sheet functions
      let temp = spreadsheet.getRange('A1').getValue();
      spreadsheet.getRange('A1').setValue(temp);
      addContentAtCursor_(spreadsheet, outputRange);
    }
  }
}

function loadRandomTablesUrl_() {
  let documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperty('randomTablesUrl') || "";
}

function saveRandomTablesUrl_(url) {
  let documentProperties = PropertiesService.getDocumentProperties();
  if (url.trim() == "") {
    documentProperties.deleteProperty('randomTablesUrl');
  } else {
    documentProperties.setProperty('randomTablesUrl', url);
  }
}

function getSheetUrls_(spreadsheet) {
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

function getInputCount_(spreadsheet, inputRangeA1) {
  var count = 0;

  if (inputRangeA1 != "") {
    const inputRange = spreadsheet.getRange(inputRangeA1);
    count = inputRange.getNumColumns();
  }

  return count;
}

function documentHasSelection_() {
  let hasSelection = false;

  let selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    hasSelection = true;
    let ui = DocumentApp.getUi();
    ui.alert('Selection Detected', 'Deselect the selected text and try again.', ui.ButtonSet.OK);
  } 

  return hasSelection;
}

function showDialog_(url, name, outputRange, inputCount, inputRange) {
  if (inputCount > 0) {
    let t = HtmlService.createTemplateFromFile('Dialog');
    t.url = url;
    t.outputRange = outputRange;
    t.inputCount = inputCount;
    t.inputRange = inputRange;
    let rowHeight = 34;
    let buttonAreaHeight = 3 * rowHeight;
    let ui = t.evaluate()
      .setWidth(400)
      .setHeight((inputCount * rowHeight) + buttonAreaHeight);

    DocumentApp.getUi().showModalDialog(ui, name);
  } else {
    console.info({ name: "showDialog_", inputCount: inputCount });
  }
}

//=========================================================

function handleLoadDialogInput(url, inputRangeA1, index) {
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  var inputRange = null;

  try {
    inputRange = spreadsheet.getRange(inputRangeA1);
  } catch (error) {
    throw new Error(`Input [${inputRangeA1}]: ${error}`);
  }

  let inputRangeValues = inputRange.getValues();
  let description = inputRangeValues[0][index];
  let value = inputRangeValues[1][index];

  let hasOptions = false;
  let options = [];
  let inputRangeValidations = inputRange.getDataValidations();
  let valueValidations = inputRangeValidations[1][index];

  if (valueValidations != null) {
    let type = valueValidations.getCriteriaType();
    let isInList = type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST;
    let isInRange = type == SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE;
    let criteriaOptions = [];
    hasOptions = isInList || isInRange;

    if (hasOptions) {
      let criteriaValues = valueValidations.getCriteriaValues();

      if (isInList) {
        criteriaOptions = criteriaValues[0].filter(option => option.toString().trim().length > 0);
      } else if (isInRange) {
        let criteriaRange = criteriaValues[0];
        criteriaOptions = criteriaRange.getValues().flat().filter(option => option.toString().trim().length > 0);  
      }

      options = criteriaOptions.reduce((accumulator, currentValue) => { 
        accumulator.push({ value: currentValue }); 
        return accumulator; 
      }, options);
    }  
  }

  let input = {
    index: index,
    isLoaded: true,
    description: description,
    value: value,
    hasOptions: hasOptions,
    options: options, 
  };

  return input;
}

function handleSubmitDialog(data) {
  let spreadsheet = SpreadsheetApp.openByUrl(data.url);
  var inputRange = null;

  try {
    inputRange = spreadsheet.getRange(data.inputRange);
  } catch (error) {
    throw new Error(`Input [${data.inputRange}]: ${error}`);
  }

  console.info(data.inputRange);

  let inputSheet = inputRange.getSheet(); 

  let valueRange = inputSheet.getRange(
    inputRange.getLastRow(), // row, 
    inputRange.getColumn(), //column, 
    1, //numRows, 
    inputRange.getNumColumns(), //numColumns
  ); 

  console.info({ 
    row: valueRange.getRow(), 
    col: valueRange.getColumn(), 
    numRows: valueRange.getNumRows(), 
    numCols: valueRange.getNumColumns(),
  });

  let rangeValues = [];
  let rowValues = [];

  rowValues = data.inputs.sort((a, b) => a.index - b.index).reduce((accumulator, currentValue) => { 
    accumulator.push( currentValue.value ); 
    return accumulator; 
  }, rowValues);

  rangeValues.push(rowValues);
  console.info(rangeValues);

  valueRange.setValues(rangeValues);

  // Write Output to Doc
  addContentAtCursor_(spreadsheet, data.outputRange);
}

//=========================================================

function addContentAtCursor_(spreadsheet, outputCell) {
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
            console.info({ name: "addContentAtCursor_", outputValue: outputValue });
          } catch (error) {
            throw new Error(`ImageURL [${outputValue}]: ${error}`);
          }
        }
      }

      // Output Cell detected or invalid Output Range type
      if (!elementInserted) {
        let outputValue = outputRange.getValue();
        elementInserted = cursor.insertText(outputValue);
        console.info({ name: "addContentAtCursor_", outputValue: outputValue });
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


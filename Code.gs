// Function to handle user input for Excel file path through CLI
function getUserInputCLI() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the Excel file path:');
  var filePath = response.getResponseText();
  
  if (!filePath) {
    throw new Error("Error: File path cannot be empty.");
  }

  // Additional validation for a valid file path
  var isValidFilePath = validateFilePath(filePath);
  if (!isValidFilePath) {
    throw new Error("Error: Invalid file path entered.");
  }
  
  return filePath;
}

// Function to validate the file path
function validateFilePath(filePath) {
  
  
  // Check if filePath is not undefined or null
  if (filePath && filePath.startsWith) {
    // check if the file path starts with "https://drive.google.com/"
    if (filePath.startsWith("https://drive.google.com/")) {
      return true;
    }
  }
  return false;
}

// Function to handle user input for Excel file path through dialog window
function selectFilePathDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('SelectFilePage')
                      .setWidth(600)
                      .setHeight(425);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Excel File');
}

// Function to create custom menu in Google Sheets
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Write File Path', 'getUserInputCLI')
      .addItem('Select File Path', 'selectFilePathDialog')
      .addToUi();
}

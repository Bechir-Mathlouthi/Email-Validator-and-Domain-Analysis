// Function to initialize the script
function initScript() {
  // Create custom menu
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Show Email List', 'showEmailList')
    .addItem('Search Email', 'searchEmail')
    .addItem('Add Email', 'addEmail')
    .addItem('Update Email', 'updateEmail')
    .addItem('Delete Email', 'deleteEmail')
    .addItem('Show Keyword List', 'showKeywordList')
    .addItem('Search Keyword', 'searchKeyword')
    .addItem('Add Keyword', 'addKeyword')
    .addItem('Update Keyword', 'updateKeyword')
    .addItem('Delete Keyword', 'deleteKeyword')
    .addToUi();
}

// Function to show the email list
function showEmailList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ALL EMAILS');
  var range = sheet.getRange('A2:A' + sheet.getLastRow());
  var values = range.getValues();
  var emails = [];
  for (var i = 0; i < values.length; i++) {
    emails.push(values[i][0]);
  }
  var ui = SpreadsheetApp.getUi();
  ui.alert('Email List', emails.join('\n'), ui.ButtonSet.OK);
}

// Function to search for an email
function searchEmail() {
  var email = SpreadsheetApp.getUi().prompt('Search Email', 'Enter the email address:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (email.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ALL EMAILS');
    var range = sheet.getRange('A:A');
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] === email.getResponseText()) {
        sheet.setActiveRange(range.getCell(i + 1, 1));
        return;
      }
    }
    SpreadsheetApp.getUi().alert('Email Not Found', 'The email address was not found in the list.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Function to add an email
function addEmail() {
  var email = SpreadsheetApp.getUi().prompt('Add Email', 'Enter the new email address:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (email.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ALL EMAILS');
    sheet.appendRow([email.getResponseText()]);
    SpreadsheetApp.getUi().alert('Success', 'The email address has been added to the list.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Function to update an email
function updateEmail() {
  var email = SpreadsheetApp.getUi().prompt('Update Email', 'Enter the email address to update:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (email.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    var newEmail = SpreadsheetApp.getUi().prompt('Update Email', 'Enter the new email address:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    if (newEmail.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ALL EMAILS');
      var range = sheet.getRange('A:A');
      var values = range.getValues();
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] === email.getResponseText()) {
          sheet.getRange(i + 1, 1).setValue(newEmail.getResponseText());
          SpreadsheetApp.getUi().alert('Success', 'The email address has been updated.', SpreadsheetApp.getUi().ButtonSet.OK);
          return;
        }
      }
      SpreadsheetApp.getUi().alert('Email Not Found', 'The email address was not found in the list.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

// Function to delete an email
function deleteEmail() {
  var email = SpreadsheetApp.getUi().prompt('Delete Email', 'Enter the email address to delete:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (email.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ALL EMAILS');
    var range = sheet.getRange('A:A');
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] === email.getResponseText()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.getUi().alert('Success', 'The email address has been deleted.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
    }
    SpreadsheetApp.getUi().alert('Email Not Found', 'The email address was not found in the list.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Function to show the keyword list
function showKeywordList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Keywords for matching');
  var range = sheet.getRange('A2:A' + sheet.getLastRow());
  var values = range.getValues();
  var keywords = [];
  for (var i = 0; i < values.length; i++) {
    keywords.push(values[i][0]);
  }
  var ui = SpreadsheetApp.getUi();
  ui.alert('Keyword List', keywords.join('\n'), ui.ButtonSet.OK);
}

// Function to search for a keyword
function searchKeyword() {
  var keyword = SpreadsheetApp.getUi().prompt('Search Keyword', 'Enter the keyword:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (keyword.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Keywords for matching');
    var range = sheet.getRange('A:A');
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] === keyword.getResponseText()) {
        sheet.setActiveRange(range.getCell(i + 1, 1));
        return;
      }
    }
    SpreadsheetApp.getUi().alert('Keyword Not Found', 'The keyword was not found in the list.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Function to add a
// Function to add a keyword
function addKeyword() {
  var keyword = SpreadsheetApp.getUi().prompt('Add Keyword', 'Enter the new keyword:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (keyword.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Keywords for matching');
    sheet.appendRow([keyword.getResponseText()]);
    SpreadsheetApp.getUi().alert('Success', 'The keyword has been added to the list.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Function to update a keyword
function updateKeyword() {
  var keyword = SpreadsheetApp.getUi().prompt('Update Keyword', 'Enter the keyword to update:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (keyword.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    var newKeyword = SpreadsheetApp.getUi().prompt('Update Keyword', 'Enter the new keyword:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    if (newKeyword.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Keywords for matching');
      var range = sheet.getRange('A:A');
      var values = range.getValues();
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] === keyword.getResponseText()) {
          sheet.getRange(i + 1, 1).setValue(newKeyword.getResponseText());
          SpreadsheetApp.getUi().alert('Success', 'The keyword has been updated.', SpreadsheetApp.getUi().ButtonSet.OK);
          return;
        }
      }
      SpreadsheetApp.getUi().alert('Keyword Not Found', 'The keyword was not found in the list.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

// Function to delete a keyword
function deleteKeyword() {
  var keyword = SpreadsheetApp.getUi().prompt('Delete Keyword', 'Enter the keyword to delete:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (keyword.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Keywords for matching');
    var range = sheet.getRange('A:A');
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] === keyword.getResponseText()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.getUi().alert('Success', 'The keyword has been deleted.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
    }
    SpreadsheetApp.getUi().alert('Keyword Not Found', 'The keyword was not found in the list.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


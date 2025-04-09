function clearCells() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TranslationChecking');
  sheet.getRange('B4:E11001').setValue("");
  sheet.getRange('B4:E11001').setBackground("#dbead4");
  sheet.getRange('K1').setValue("");
  sheet.getRange('K4').setValue("");
  sheet.getRange('M4').setValue("");
  sheet.getRange('I4').setValue("");
  sheet.getRange('I4').setBackground("#dbead4");
  
  SpreadsheetApp.getUi().alert('Sheet Cleared.\n\nYou can always undo this.');
  Logger.log('Sheet cleared.');
}

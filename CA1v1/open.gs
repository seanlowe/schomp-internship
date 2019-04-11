function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Utilities').addItem("Rollover Data for New Month",'rollover').addToUi();
  ss.setActiveSheet(ss.getSheetByName('Quick Access')).setActiveSelection('B3');
  ss.getSheetByName('List').hideSheet();
  ss.getSheetByName('Master').hideSheet();
}
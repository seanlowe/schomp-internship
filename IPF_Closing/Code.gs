function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Utilities").addItem("Scrape Data from 30-Day", "scrape").addToUi();
}


function scrape() {
  // created by Sean Lowe, 8/14/2018
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var day = ss.getSheetByName("30 Day").getRange(3, 1, 26, 22).getValues();
  //Logger.log(day);
  var target;
  var arr = [];
  var q = ui.prompt("Date", "Enter the date to be associated with this data.", ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() == ui.Button.OK) {
    for (var i = 0; i < 7; i++) {
      // 0 (name) 4 (internet) 8 (phone) 12 (fresh) 17 (all)
      // add 10 to i to get next value for current client advisor
      target = ss.getSheetByName(day[i][0]); 
      
      arr[i] = [];
      arr[i][0] = day[i][0];           // name
      arr[i][1] = q.getResponseText(); // date
      arr[i][2] = day[i][4];    // i new
      arr[i][3] = day[i+10][4]; // i used
      arr[i][4] = day[i+20][4]; // i total
      arr[i][5] = day[i][8];    // p new
      arr[i][6] = day[i+10][8]; // p used
      arr[i][7] = day[i+20][8]; // p total
      arr[i][8] = day[i][12];    // f new
      arr[i][9] = day[i+10][12]; // f used
      arr[i][10] = day[i+20][12]; // f total
      arr[i][11] = day[i][17];    // all new
      arr[i][12] = day[i+10][17]; // all used
      arr[i][13] = day[i+20][17]; // all total
      
      target.getRange(1,target.getLastColumn()-1,arr.length).setValues(arr[i]); 
    }
  }
  return;
}
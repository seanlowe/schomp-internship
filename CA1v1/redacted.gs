/*
------------------------------------------------------------------------------------------------
function refreshCA() { 
  // created by Sean Lowe, 7/28/2018
  var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List").getRange("A15:B").getDisplayValues();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List").getRange(15,1,range.length, range[0].length).setValues(range);
  range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List").getRange("A15:B").sort( { column: 1, ascending: true } );
  return;
}
------------------------------------------------------------------------------------------------
function resetRefresh() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List").getRange(15, 1).setValue("=IMPORTRANGE(C3,\"1v1!A13:B\")");
  refreshCA();
}
------------------------------------------------------------------------------------------------
function newSheet() {
  // created by Sean Lowe, 7/26/2018 // version 2.0
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var link = ss.getSheetByName("List").getRange(3, 3).getValue();
  var sheet; var range; var arr;
  var nu, uu, npvr, upvr, cr, npro, upro, acc; nu = uu = npvr = upvr = cr = npro = upro = acc = 0;
  var CAname = ui.prompt("Name of New Sheet", "Enter the name of the Client Advisor who's sheet you are creating (ex. First Last).", ui.ButtonSet.OK_CANCEL);
  if (CAname.getSelectedButton() == ui.Button.OK) {
    CAname = CAname.getResponseText();
    var tname = ui.prompt(CAname+"\'s Team", "Enter the Team Name that " + CAname + " can be found on (ex. Jeff).", ui.ButtonSet.OK);
    tname = tname.getResponseText();
    var CArow = ui.prompt("Row Number", "Enter the Row Number that " + CAname + " appears on their team sheet (ex. 3).", ui.ButtonSet.OK);
    CArow = parseInt(CArow.getResponseText());
    // add to CArow to get: // 16 for closing ratio // 21 for new units // 22 for used units // 23 for new product // 24 for New PVR // 25 for used product // 26 for Used PVR // 27 for accessories
    cr = CArow + 16; 
    nu = CArow + 21; uu = CArow + 22; 
    npro = CArow + 23; npvr = CArow + 24; 
    upro = CArow + 25; upvr = CArow + 26; 
    acc = CArow + 27;
    sheet = ss.getSheetByName("Master").copyTo(ss).setName(CAname);
    arr = ss.getSheetByName(CAname).getRange(8, 7, 7).getValues();
    arr[0][0] = "=IMPORTRANGE('List'!C3, \"'Team " + tname + "'!B" + nu + "\") + IMPORTRANGE('List'!C3, \"'Team " + tname + "'!B" + uu + "\")"; // new units + used units
    arr[1][0] = "=G10+G11";
    arr[2][0] = "=IMPORTRANGE('List'!C3, \"'Team " + tname + "'!B" + npvr + "\")"; // new PVR
    arr[3][0] = "=IMPORTRANGE('List'!C3, \"'Team " + tname + "'!B" + upvr + "\")"; // used PVR
    arr[4][0] = "=IFERROR(IMPORTRANGE('List'!C3, \"'Team " + tname + "'!B" + acc + "\"), 0)"; // accessories
    arr[5][0] = "=IMPORTRANGE('List'!C3, \"'Team " + tname + "'!B" + cr + "\")"; // closing ratio
    arr[6][0] = "=((IMPORTRANGE('List'!C3, \"'Team " + tname + "'!B" + npro + "\") + IMPORTRANGE('List'!C3, \"'Team " + tname + "'!B" + upro + "\"))-G12)/G8"; 
    // arr[6][0] is product/deal --> (new product + used product) - accessories / (new units + used units)
    ss.getSheetByName(CAname).getRange(8, 7, arr.length, 1).setValues(arr);
    range = ss.getSheetByName("List").getRange(15, 1, ss.getSheetByName("List").getLastRow()-1,  1).getValues();
    range[range.length] = [];
    range[range.length-1][0] = CAname;
    ss.getSheetByName("List").getRange(15, 1, range.length, 1).setValues(range);
    refreshCA();
    //ss.setActiveSheet(sheet);
  }
}
------------------------------------------------------------------------------------------------
function replaceCA() {
  // created by Sean Lowe, 7/28/2018
  // version 1.0
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var CAname = ss.getActiveSheet().getSheetName();
  var link = ss.getSheetByName("List").getRange(3, 3).getValue();
  var arr; var range;
  var nu, uu, npvr, upvr, cr, npro, upro, acc; nu = uu = npvr = upvr = cr = npro = upro = acc = 0;
  var comp = false; var replace = false;
  var newCA = ui.prompt("Replace CA", "Enter the name of the CA that is going to replace " + CAname + ".", ui.ButtonSet.OK_CANCEL);
  if (newCA.getSelectedButton() == ui.Button.OK) {
    newCA = newCA.getResponseText();
    var nteam = ui.prompt("Team of New CA", "Please enter the Team Name that " + newCA + " is assigned to. If assigned "
                          + "to the same team that " + CAname + " was, just enter 'Same'.", ui.ButtonSet.OK);
    nteam = nteam.getResponseText();
    var nrow = ui.prompt("Row of New CA", "Please enter the Row Number that " + newCA + " is assigned to. If assigned "
                         + "to the same row that " + CAname + " was, just enter 'Same'.", ui.ButtonSet.OK);
    if (nrow.getResponseText().toLowerCase() != "same" || nteam.toLowerCase() != "same") { // if either the row or the team is different, just re-do them both.
      nrow = parseInt(nrow.getResponseText());
      cr = nrow + 16; 
      nu = nrow + 21; uu = nrow + 22; 
      npro = nrow + 23; npvr = nrow + 24; 
      upro = nrow + 25; upvr = nrow + 26; 
      acc = nrow + 27;
      arr = ss.getSheetByName(CAname).getRange(8, 7, 7).getValues();
      arr[0][0] = "=IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + nu + "\") + IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + uu + "\")"; // new units + used units
      arr[1][0] = "=G10+G11";
      arr[2][0] = "=IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + npvr + "\")"; // new PVR
      arr[3][0] = "=IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + upvr + "\")"; // used PVR
      arr[4][0] = "=IFERROR(IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + acc + "\"), 0)"; // accessories
      arr[5][0] = "=IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + cr + "\")"; // closing ratio
      arr[6][0] = "=((IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + npro + "\") + IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + upro + "\"))-G12)/G8"; 
      // range[6][0] is product/deal --> (new product + used product) - accessories / (new units + used units)
      ss.getSheetByName(CAname).getRange(8, 7, arr.length, 1).setValues(arr);
    }
    // find name on List and correct it, then change the old CA's sheetName to the new name
    range = ss.getSheetByName("List").getRange(15, 1, ss.getSheetByName("List").getLastRow()-1,  1).getValues();
    for (var i = 0; i < range.length && !comp; i++) {
      if (range[i][0] == CAname) {
        range[i][0] = newCA;
        ss.getSheetByName(CAname).setName(newCA);
        comp = true;
      }
    }
    ss.getSheetByName("List").getRange(15, 1, range.length, 1).setValues(range);
    //refreshCA();
  }
}
------------------------------------------------------------------------------------------------
function reassignCA() {
  // created by Sean Lowe, 7/28/2018
  // version 1.0
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var CAname = ss.getActiveSheet().getSheetName();
  var link = ss.getSheetByName("List").getRange(3, 3).getValue();
  var arr; var range;
  var nu, uu, npvr, upvr, cr, npro, upro, acc; nu = uu = npvr = upvr = cr = npro = upro = acc = 0;
  var comp = false; var replace = false;
  var nteam = ui.prompt("Assign CA to different Team", "Enter the Team Name that " + CAname + " is getting reassigned to.", ui.ButtonSet.OK_CANCEL);
  if (nteam.getSelectedButton() == ui.Button.OK) {
    nteam = nteam.getResponseText();
    var nrow = ui.prompt("New Row Number", "Please enter the Row Number that " + CAname + " is assigned to on their new team.", ui.ButtonSet.OK);
    nrow = parseInt(nrow.getResponseText());
    cr = nrow + 16; 
    nu = nrow + 21; uu = nrow + 22; 
    npro = nrow + 23; npvr = nrow + 24; 
    upro = nrow + 25; upvr = nrow + 26; 
    acc = nrow + 27;
    arr = ss.getSheetByName(CAname).getRange(8, 7, 7).getValues();
    arr[0][0] = "=IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + nu + "\") + IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + uu + "\")"; // new units + used units
    arr[1][0] = "=G10+G11";
    arr[2][0] = "=IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + npvr + "\")"; // new PVR
    arr[3][0] = "=IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + upvr + "\")"; // used PVR
    arr[4][0] = "=IFERROR(IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + acc + "\"), 0)"; // accessories
    arr[5][0] = "=IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + cr + "\")"; // closing ratio
    arr[6][0] = "=((IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + npro + "\") + IMPORTRANGE('List'!C3, \"'Team " + nteam + "'!B" + upro + "\"))-G12)/G8"; 
    // arr[6][0] is product/deal --> (new product + used product) - accessories / (new units + used units)
    ss.getSheetByName(CAname).getRange(8, 7, arr.length, 1).setValues(arr);
  }
}
------------------------------------------------------------------------------------------------
// =IMPORTRANGE(List!C3,VLOOKUP(I7,List!A15:B,2))
units - I12+I13
total PVR - G10+G11
new PVR -  I15
used PVR - I17
Acc - I18
CR - (I8+I9+I10)/3
PPD - ((I14+I16)-I18)/G8
------------------------------------------------------------------------------------------------
function myFunction() {
  // created by Kennen Lawrence, 8/1/18
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getActiveSheet();
  var range=sheet.getRange('I8:I18');
  var range2=sheet.getRange('G8:G14');
  var editors='kennen.lawrence@a2zsync.com';
  range.protect().addEditor(editors);
  range2.protect().addEditor(editors);
  
}
------------------------------------------------------------------------------------------------
function removeCA() {
  // version 1.0
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheets = ss.getSheets();
  var range = ss.getSheetByName("List").getRange(2, 1, ss.getSheetByName("List").getLastRow()-1,  1).getValues();
  var comp = false;
  var name = ui.prompt("Remove an Existing Sheet", "Enter the name of the Client Advisor exactly how it appears as the Sheet Name.", ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() == ui.Button.CANCEL) { return; } //This has to be added because without it if the user presses cancel the script would've deleted them anyway -Kennen
  for (var i = 0; i < sheets.length && !comp; i++) {
    //Logger.log("sheetname="+sheets[i].getSheetName()+"       name="+name.getResponseText());
    if (sheets[i].getSheetName() == name.getResponseText()) {
      ss.deleteSheet(sheets[i]);
      comp = true;
    }
  }
  //for (i = 0; i < range.length && comp; i++) {
    //Logger.log("range["+i+"][0]="+range[i][0]+"       name="+name.getResponseText());
    //if (range[i][0] == name.getResponseText()) {
      //ss.getSheetByName("List").deleteRow(i+2);
      //ss.getSheetByName("List").insertRows(i+2, 1);
      //comp = false;
    //}
  //}
  //resetDV();
}
------------------------------------------------------------------------------------------------







*/
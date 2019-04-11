// scripts created by Sean Lowe
// 7/27/18 through 8/1/2018

function sweep() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() != "Master" && sheets[i].getSheetName() != "List" && sheets[i].getSheetName() != "Quick Access") {
      ss.deleteSheet(sheets[i]);
    } } /*resetDV(); */ }

function redoProtect() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var name = "";
  var pros = [];
  var editors = ["alexa.gerner@schomp.com","kennen.lawrence@a2zsync.com"];
  for (var i = 0; i < sheets.length; i++) {
    name = sheets[i].getSheetName();
    if (name != "List" && name != "Quick Access" && name != "Master") {
      pros = sheets[i].getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var j = 0; j < pros.length; j++) { pros[j].remove(); }
      var prot = ss.getSheetByName(name).getRange("I8:I18").protect(); // axcessa information
      prot.removeEditors(prot.getEditors()).addEditors(editors);
      prot = ss.getSheetByName(name).getRange("G8:G12").protect(); // all current
      prot.removeEditors(prot.getEditors()).addEditors(editors);   // month data
      prot = ss.getSheetByName(name).getRange("G14").protect();    // except
      prot.removeEditors(prot.getEditors()).addEditors(editors);   // cell G13 (closing ratio)
    }
  }
}

function resetDV() {
  var qa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quick Access");
  var list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List");
  var dv = SpreadsheetApp.newDataValidation();
  dv.setAllowInvalid(false);
  dv.requireValueInRange(list.getRange(15, 1, list.getLastRow()-15));
  qa.getRange("B3").setDataValidation(dv);
}

function hide(sheet) { if (sheet != null && sheet != "Quick Access" && sheet.isSheetHidden() == false) { sheet.hideSheet(); } }
function caName() { /* created by Kennen Lawrence, 8/1/18 */ return SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getSheet().getSheetName(); }

function babyboom() {
  // version 1.0
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheets = ss.getSheets();
  var list = ss.getSheetByName("List").getRange(15, 1, ss.getSheetByName("List").getLastRow()-15).getDisplayValues();
  var found = false;
  for (var i = 0; i < sheets.length; i++) {
    found = false;
    for (var j = 0; j < list.length && !found; j++) {
      if (list[j][0] == sheets[i].getSheetName()) { list.splice(j, 1); j--; found = true; }
    }
  }
  for (i = 0; i < list.length; i++) {
    newSheet(list[i][0]);
    hide(ss.getSheetByName(list[i][0]));
  }
}

function rollover() {
  // version 1.1
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi(); var name = "";
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    name = sheets[i].getSheetName();
    if (name != "List" && name != "Quick Access" && name != "Master") {
      sheets[i].getRange(8, 3, 7, 1).setValues(sheets[i].getRange(8, 7, 7, 1).getDisplayValues());
    }
  }
  // replace current month SAD link with new month link and blank out new month
  ss.getSheetByName("List").getRange(4, 3).moveTo(ss.getSheetByName("List").getRange(3, 3));
  resetDV();
}

function viewQA() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  ss.setActiveSheet(ss.getSheetByName("Quick Access"));
  try { sheet.hideSheet(); } // prevents protection errors
  catch (err) { return; }
}

function onEdit(e) { 
  //Logger.log(e.range.getA1Notation());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var found = false;
  if (e.source.getSheetName() == "Quick Access" && e.value != "" && e.range.getA1Notation() == "B3") {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetName() == e.value) {
        ss.setActiveSheet(ss.getSheetByName(e.value));
        found = true; break;
      }
    }
    if (!found) {
      var alert = ui.alert('Not found', 'This CA does not have a sheet. Would you like to make one?', ui.ButtonSet.YES_NO);
      if (alert == ui.Button.YES) { newSheet(e.value); ss.setActiveSheet(ss.getSheetByName(e.value)); }
      else { ss.toast("No Sheet will be created for "+ e.value +" at this time.","Cancelled"); }
    }
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quick Access").getRange(3, 2).setValue("");
  }
}

function newSheet(name) {
  // version 5.0.1 (Spring Cleaning)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var editors = ["alexa.gerner@schomp.com","kennen.lawrence@a2zsync.com"]
  if (name == "" || name == undefined || name == null) { return; }
  ss.getSheetByName("Master").copyTo(ss).setName(name).getRange('D31').setValue(name);
  
  // make sure no one can screw up the sheet
  var prot = ss.getSheetByName(name).getRange("I8:I18").protect(); // axcessa information
  prot.removeEditors(prot.getEditors()).addEditors(editors);
  prot = ss.getSheetByName(name).getRange("G8:G14").protect(); // current month data
  prot.removeEditors(prot.getEditors()).addEditors(editors);
}

// script file for all redacted functions

/*

//function onEdit(e) {
//  // created by Sean Lowe, 7/1/2018
//  // version 0.1
//  Logger.log(e.source.getSheetName());
//  if (e.source.getSheetName() != "Form Responses 1") { return; }
//  Logger.log("passed sheet check");
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var range; //= ss.getActiveSheet().getRange(12, 1, 34, 11).getValues();
//  var form = ss.getSheetByName("Form Responses 1");
//  var fr = form.getRange(2, 2, form.getLastRow()-1, form.getLastColumn()-1).getValues();
//  //Logger.log(fr);
//  var arr = [];
//  var sheets = ss.getSheets();
//  var sheetNames = [];
//  var temp;
//  var assigned = false;
  
//  for (var i = 0; i < sheets.length; i++) {
//    sheetNames[i] = sheets[i].getSheetName();
//  }
  
//  for (i = 0; i < fr.length; i++) {
//    arr[i] = [];
//    arr[i][0] = fr[i][0];   // name
//    arr[i][1] = fr[i][2];   // department
//    arr[i][2] = fr[i][3];   // day
//    for (var j = 0; j < sheetNames.length; j++) {
//      if (arr[i][2] == sheetNames[j]) {  // check what day it is
//        for (var k = 4; k < fr[0].length; k++) {
//          if (fr[i][k] != "") {
//            arr[i][3] = fr[i][k]; // time
//          }  }  }  }  }
  
//  for (i = 0; i < arr.length; i++) {  
//    range = ss.getSheetByName(arr[i][2]).getRange(12, 1, 34, 11).getValues();
      // rows 1-10 1st time /// 11-12 gap /// 13-23 2nd time /// 24-25 gap /// 26-36 3rd time
//    if (arr[i][1] == "Sales / Genius") {
//      temp = range[i][4].toLocaleTimeString().split(" MST");
//      temp = temp[0].split(":")[0] + ":" + temp[0].split(":")[1] + " " + temp[0].split(" ")[1];
      //Logger.log(temp); Logger.log(arr[i][3]);
//      if (temp == arr[i][3]) { // 10:00 AM
        //Logger.log("reached inside range[0][4] if");
//        assigned = false;
//        for (k = 0; k < 11; k += 2) {
//          for (j = 0; j < 4; j += 2) {
            //Logger.log("Reached first time slot");  Logger.log("k="+k + "   " + "j="+j);  Logger.log(range[k][j] + "   ---/--- " + arr[i][0]);
//            if (range[k][j] == undefined || range[k][j] == "") { 
//              Logger.log("assigned something"); 
//              range[k][j] = arr[i][0]; 
//              assigned = true;
//              break; 
//            }
//          }
//          if (assigned) { break; }
//        }
//      }
//      else if (temp == arr[i][3]) { // 12:00 PM
//        assigned = false;
//        for (k = 12; k < 23; k += 2) {
//          for (j = 0; j < 4; j += 2) {
//            Logger.log("Reached second time slot");
//            if (range[k][j] == undefined || range[k][j] == "") { 
//              Logger.log("assigned something"); 
//              range[k][j] = arr[i][0]; 
//              assigned = true;
//              break; 
//            }
//          }
//          if (assigned) { break; }
//        }
//      }
//      else if (temp == arr[i][3]) { // 2:30 PM
//        assigned = false;
//        for (k = 24; k < 35; k += 2) {
//          for (j = 0; j < 4; j += 2) {
//            Logger.log("Reached third time slot");
//            if (range[k][j] == undefined || range[k][j] == "") { 
//              Logger.log("assigned something"); 
//              range[k][j] = arr[i][0]; 
//              assigned = true;
//              break; 
//            }
//          } // end j
//          if (assigned) { break; }
//        } // end k 
//      } // end else if
//    } // end 'sales / genius' if
//    else { // entry is on aftersales team
//      if (temp == arr[i][3]) {        }
//      else if (range[12][4] == arr[i][3]) {        }
//      else if (range[25][4] == arr[i][3]) {        }
//    }
//    ss.getSheetByName(arr[i][2]).getRange(12, 1, 34, 11).setValues(range);
//  } // end i for
//  Logger.log("end of onEdit function");  //Logger.log(range);
//}


function onEdit(e) {
  // created by Sean Lowe, 7/12/2018
  // version 0.2
  Logger.log(e.source.getSheetName());
  if (e.source.getSheetName() != "Form Responses 1") { Logger.log("edit received on the wrong sheet .. exiting"); return; }
  Logger.log("passed sheet check");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range;
  var form = ss.getSheetByName("Form Responses 1");
  var fr = form.getRange(2, 2, form.getLastRow()-1, form.getLastColumn()-1).getValues();
  var arr = []; var sheetNames = [];
  var sheets = ss.getSheets();
  var assigned = false;
  var temp;  var i = 0;  var j = 0;  var k = 0;  var l = 0; var m = 0;
  
  for (i = 0; i < sheets.length; i++) { sheetNames[i] = sheets[i].getSheetName(); }
  
  for (i = 0; i < fr.length; i++) {
    arr[i] = [];
    arr[i][0] = fr[i][0];   // name
    arr[i][1] = fr[i][2];   // department
    arr[i][2] = fr[i][3];   // day
    for (var j = 0; j < sheetNames.length; j++) {
      if (arr[i][2] == sheetNames[j]) {  // check what day it is
        for (var k = 4; k < fr[0].length; k++) {
          if (fr[i][k] != "") {
            arr[i][3] = fr[i][k]; // time
          }  }  }  }  }
  
  for (i = 0; i < arr.length; i++) {
    Logger.log("ITERATION i= "+i+" arr["+i+"][2]="+arr[i][2]);
    range = ss.getSheetByName(arr[i][2]).getRange(12, 1, 34, 11).getValues();
    // rows 0-9 1st time /// 10-11 gap /// 12-22 2nd time /// 23-24 gap /// 25-35 3rd time
    if (arr[i][1] == "Sales / Genius") {
      //Logger.log("inside 'sales/genius' if");
      if (i == 0 || i == 12 || i == 25) {
        temp = range[i][4].toLocaleTimeString().split(" MST");
        temp = temp[0].split(":")[0] + ":" + temp[0].split(":")[1] + " " + temp[0].split(" ")[1];
      }
      switch (arr[i][3]) { // set upper bound for time for loop
        case 0:
          l = 10; m = 0;
          break;
        case 12:
          l = 23; m = 12;
          break;
        case 25:
          l = 36; m = 25;
          break;
        default:
          // check what area it lays in between
          if (i < 12) { l = 10; m = 0; }
          else if (i < 25) { l = 23; m = 12; }
          else if (i > 25) { l = 36; m = 25; }
          break;
      }
      if (arr[i][3] == temp) { // check times
        Logger.log(arr[i]);
        Logger.log("inside time check of temp="+temp+"       "+"arr["+i+"][3]="+arr[i][3]);
        assigned = false;
        for (k = m; k < l; k += 2) { // delete everything
          for (j = 0; j < 4; j += 2) {
            if (range[k][j] != "") { range[k][j] = ""; }
          } }
        for (k = m; k < l; k += 2) {
          for (j = 0; j < 4; j += 2) {
            Logger.log("i="+i+"   "+"j="+j+"   "+"k="+k+"   "+"l="+l+"   ");
            if (range[k][j] == undefined || range[k][j] == "") { 
              Logger.log("assigned a time slot");
              range[k][j] = arr[i][0]; 
              assigned = true;
              break;
            }  }
          if (assigned) { break; }
        }  }
    } else { return; }
    ss.getSheetByName(arr[i][2]).getRange(12, 1, 34, 11).setValues(range);
  } // end arr.length for loop
} // end of function


*/

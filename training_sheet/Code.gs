function trigger() { onEdit("Trigger"); }
function nightBefore() { remind("Trigger"); }

function month(mon) {
  switch(mon) {
    case 1: return "January";
    case 2: return "February";
    case 3: return "March";
    case 4: return "April";
    case 5: return "May";
    case 6: return "June";
    case 7: return "July";
    case 8: return "August";
    case 9: return "September";
    case 10: return "October";
    case 11: return "November";
    case 12: return "December";
  }
}

function modTrig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var remind; var type; var interval;
  remind = ui.prompt("title", "prompt", ui.ButtonSet.OK_CANCEL);
  if (remind.getSelectedButton() == ui.Button.OK) {
    for (var i = 0; i < triggers.length; i++) {
      Logger.log(triggers[i].getHandlerFunction());
      if (triggers[i].getHandlerFunction() == "remind") {
        Logger.log("found remind function trigger");
        remind = triggers[i];
      }
    }
    ScriptApp.deleteTrigger(remind);
    type = ui.prompt("How often would you like to run the trigger? Please enter one of the following: 'hours', 'minutes', or 'days'.", ui.ButtonSet.OK);
    remind = ScriptApp.newTrigger("remind").timeBased().everyHours(type).create();
  }  
}

function tomor(today) {
  var tomorrow;
  if (today.getMonth()+1 == 1 && today.getDate() == 31) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 2 && today.getDate() == 28) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 3 && today.getDate() == 31) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 4 && today.getDate() == 30) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 5 && today.getDate() == 31) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 6 && today.getDate() == 30) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 7 && today.getDate() == 31) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 8 && today.getDate() == 31) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 9 && today.getDate() == 30) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 10 && today.getDate() == 31) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 11 && today.getDate() == 30) { tomorrow = month(today.getMonth()+2) + " 1"; }
  else if (today.getMonth()+1 == 12 && today.getDate() == 31) { tomorrow = month(1) + " 1"; }
  else { tomorrow = month(today.getMonth()+1) + " " + (today.getDate()+1); }
  return tomorrow;
}

function remind(s) {
  Logger.log(s);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = []; var sheets = []; var timeRows = []; var timeCols = [];
  var arr = []; var final = [];
  var now = new Date(); var today = new Date(); var tomorrow;
  tomorrow = tomor(today);
  Logger.log(tomorrow)
  today = month(today.getMonth()+1) + " " + today.getDate();
  Logger.log(today);
  var time; var temp = ""; var nowh, nowm, timeh, timem;
  var dif; var gen = false;
  var form = ss.getSheetByName("Form Responses 1");
  var fr = form.getRange(2, 2, form.getLastRow()-1, form.getLastColumn()-1).getValues();
  // MST: UTC−07:00 (now)
  // MDT: UTC−06:00 (time)
  nowh = now.getUTCHours()-6;
  nowm = now.getMinutes();
  //time = time.getUTCHours()-7;
  //Logger.log("now="+now+"      time="+time);
  
  sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    sheetNames[i] = sheets[i].getSheetName();
  }
  
  if (s == undefined) {
    //Logger.log("before timeRows");
    timeRows = [12,18,24]; timeCols = [3, 5];
    for (i = 0; i < sheetNames.length; i++) {
      //Logger.log("i="+i);
      if (sheetNames[i] == today) {
        for (var j = 0; j < timeRows.length && !gen; j++) {
          for (var k = 0; k < timeCols.length; k++) {
            //Logger.log("on today's sheet");
            gen = false;
            time = ss.getSheetByName(sheetNames[i]).getRange(timeRows[j], timeCols[k]).getValue();
            timeh = time.getUTCHours() - 7; timem = time.getMinutes();
            //Logger.log("hour="+timeh+"       min="+timem);
            //Logger.log("nhour="+nowh+"       nmin="+nowm);
            dif = timeh-nowh;
            Logger.log("dif="+dif);
            if (dif < 0) {
              Logger.log("time slot at "+timeh+":"+timem+" is already passed");
            } else if (dif > 1) {
              Logger.log("more than one hour remains before training slot at "+timeh+":"+timem);
            } else if (dif <= 1) {
              if (dif == 1) {
                Logger.log("in the same hour as time slot at "+timeh+":"+timem+". send email.");
                gen = true;
              } else if (dif == 0 || timeh == nowh) {
                // check if time is passed -> if not, loop through section and grab names
                if (nowm > timem) {
                  Logger.log("time slot at "+timeh+":"+timem+" has already passed");
                } else {
                  gen = true;
                  Logger.log("time slot at "+timeh+":"+timem+" not reached. send email");
                }
              }
            }
            if (gen) {
              var col = 0;
              switch (timeCols[k]) { case 3: col = 1; break; case 5: col = 6; break; }
              arr = ss.getSheetByName(sheetNames[i]).getRange(timeRows[j], col, 5, 2).getValues();
              Logger.log(arr);
              break;
            }
          } // k for
        } // j for
      } // sheetNames check
    } // i for
    Logger.log("after time checks");
    for (i = 0; i < fr.length; i++) {
    gen = false;
      for (j = 0; j < arr.length && !gen; j++) {
        for (k = 0; k < arr[0].length && !gen; k++) {
          Logger.log("fr[i][0]="+fr[i][0]+"      arr[j][k]="+arr[j][k]);
          if (fr[i][0] == arr[j][k]) {
            Logger.log("matched someone on upcoming time slot in form responses. send email");
            gen = true;
            if (fr[i][1] != undefined && fr[i][1] != "") {
              MailApp.sendEmail(fr[i][1],'Upcoming Training Time Slot',"Hello!\n\nThis is a friendly reminder that you have signed up to attend an upcoming training at time slot: "+timeh+":"+timem+". \n\nThanks!");
            }
          }
        }
      }
    }
  } // end of if s == undefined (normal run trigger)
  else if (s != undefined && s == "Trigger") {
    
  }
} // end of function
// MailApp.sendEmail('kennen.lawrence@schomp.com','HELP Sales Daily_March',input.getResponseText());

function onEdit(e, x) {
  // created by Sean Lowe and Kennen Lawrence, 7/16/2018
  // version 1.4
  //Logger.log(e); //Logger.log(e.source.getSheetName());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (e != "Trigger") { var ui = SpreadsheetApp.getUi(); }
  var form = ss.getSheetByName("Form Responses 1");
  var fr = form.getRange(2, 2, form.getLastRow()-1, form.getLastColumn()-1).getValues();
  var arr = []; var sheetNames = []; var times = []; var timeRows = []; var sheets = ss.getSheets();
  var i, j, k, l, m; i = j = k = l = m = 0;
  var range, temp, col, check, full, row, col;
  var del = false; var add = false; var comp = false; // variables to tell us if we're deleting or adding a name to the form responses page
  
  // delete any and all empty rows from Form Responses before doing anything.
  var count = 2;
  for (i = 0; i < fr.length; i++) {
    if (fr[i][0] == "") { form.deleteRow(i + count); count--; }
  }
  
  if (e != "Open" && e != "Trigger" && e.source.getSheetName() != "Form Responses 1") { 
    //if (true) { Logger.log("wrong sheet edited .. exiting"); return; } // uncomment this to make it ignore non-form response edits
    //Logger.log(e.range.getA1Notation());
    var esh = e.source.getSheetName();
    var etime; var ename = ""; var oldname = "";
    
    // section for re-writing form responses to form response sheet
    for (i = 0; i < fr.length; i++) {
      if (fr[i][3] != undefined && fr[i][3] != "") { fr[i][3] = "'" + fr[i][3]; }
      else { fr[i][7] = "'" + fr[i][7]; }
      if (fr[i][3] != undefined && fr[i][3] != "") {
        for (k = 4; k < fr[0].length; k++) {
          if (fr[i][k] != "") {  fr[i][k] = "'" + fr[i][k];  }
        }
      } else {
        for (k = 8; k < fr[0].length; k++) {
          if (fr[i][k] != "") {  fr[i][k] = "'" + fr[i][k];  }
        }
      }
    }   
    
    if (e.range.getA1Notation() != undefined) {
      var etemp = e.range.getValues();
    }
    
    if (etemp.length > 1) { // multi-cell edit
      Logger.log("multi cell edit");
    } else {
      //Logger.log("single cell edit");
      if (e.range.getColumn() == 1 || e.range.getColumn() == 2) {
        etime = ss.getSheetByName(esh).getRange(e.range.getRow(), 3).getValue().toLocaleTimeString();
      } else if (e.range.getColumn() == 6 || e.range.getColumn() == 7) {
        etime = ss.getSheetByName(esh).getRange(e.range.getRow(), 5).getValue().toLocaleTimeString();
      }
      temp = etime.split(" MST");
      etime = temp[0].split(":")[0] + ":" + temp[0].split(":")[1] + " " + temp[0].split(" ")[1];
      //Logger.log(esh); Logger.log(etime); Logger.log("e.value="+e.value+"       e.oldvalue="+e.oldValue); Logger.log(typeof (e.value));
      if ((typeof (e.value) != "object") && (e.value != "" && e.value != undefined) && (e.oldValue != "" && e.oldValue != undefined)) { oldname = e.oldValue; ename = e.value; add = true; del = true; }
      else if (e.oldValue != ""  && e.oldValue != undefined) { oldname = e.oldValue; add = false; del = true; }
      else if (e.oldValue == undefined) { ename = e.value; add = true; del = false; } 
      //Logger.log("The string ename contains: " + ename); Logger.log("The string oldname contains: " + oldname); Logger.log(add);
      
      // adding, editing, or deleting an entry from the FR sheet
      comp = false;
      for (i = 0; i < fr.length && !comp; i++) {
        //Logger.log("i="+i+"   fr["+i+"][0]="+fr[i][0]+"     oldname="+oldname+"        ename="+ename+"      add="+add+"     del="+del);
        //if (fr[i][0] == ename && add == true) { break; }
        if (fr[i][0] == "" && add && !del) {
          fr[i][0] = ename;
          fr[i][1] = "";
          row = e.range.getRow(); 
          if (e.range.getColumn() == 1 || e.range.getColumn() == 2) {
            col = 3; fr[i][2] = "Sales / Genius";
          } else {
            col = 7; fr[i][2] = "Aftersales (Parts & Service)";
          } // department
          fr[i][col] = "'" + esh;
          fr[i][col+1] = "'" + etime;
          comp = true;
          for (j = 0; j < fr[0].length; j++) { if (fr[i][j] == undefined || fr[i][j] == null) fr[i][j] = ""; }
          //Logger.log("added something");
        }
        else if (fr[i][0] == oldname && del && !add) { 
          form.deleteRow(i+2);
          //fr.splice(i, 1); i--;
          comp = true;
          //Logger.log("deleted something");
          //Logger.log("calling onEdit again to make sure everything is dandy");
          onEdit("Open");
          //Logger.log("finished recursive onEdit call");
          return;
        }
        else if (add && del && fr[i][0] == oldname) {
          fr[i][0] = ename; 
          comp = true;
          //Logger.log("switched name on existing entry");
        }
      } // add or deleting for loop
      
      //Logger.log("reaching the !comp if");
      if (!comp) {
        fr[fr.length] = [];
        fr[fr.length-1][0] = ename; 
        fr[fr.length-1][1] = "";
        row = e.range.getRow(); 
        if (e.range.getColumn() == 1 || e.range.getColumn() == 2) {
          col = 3; fr[fr.length-1][2] = "Sales / Genius";
        } else {
          col = 7; fr[fr.length-1][2] = "Aftersales (Parts & Service)";
        } // department
        fr[fr.length-1][col] = "'" + e.source.getSheetName();
        fr[fr.length-1][col+1] = "'" + etime
        comp = true;
        for (i = 0; i < fr[0].length; i++) { if (fr[fr.length-1][i] == undefined || fr[fr.length-1][i] == null) fr[fr.length-1][i] = ""; }
        //Logger.log("added something"); //Logger.log(fr[fr.length-1]); //Logger.log(fr);
      }
      
      form.getRange(2, 2, fr.length, fr[0].length).setValues(fr);
      //Logger.log("calling onEdit again to make sure everything is dandy");
      onEdit("Open");
      //Logger.log("finished recursive onEdit call");
      return;
    }
    
  } else {
    //Logger.log("passed sheet check");
    fr = form.getRange(2, 2, form.getLastRow()-1, form.getLastColumn()-1).getValues();
    
    // grab all sheetnames and put them into an array
    for (i = 0; i < sheets.length; i++) { sheetNames[i] = sheets[i].getSheetName(); }
    
    // pull all form responses into an array
    for (i = 0; i < fr.length; i++) {
      arr[i] = [];
      arr[i][0] = fr[i][0];   // name
      arr[i][1] = fr[i][2];   // department
      if (fr[i][3] != undefined && fr[i][3] != "") { arr[i][2] = fr[i][3]; } // day
      else { arr[i][2] = fr[i][7]; }
      for (j = 0; j < sheetNames.length; j++) {
        if (arr[i][2] == sheetNames[j]) {  // check what day it is
          arr[i][3] == "Sheet";
          if (fr[i][3] != undefined && fr[i][3] != "") {
            for (k = 4; k < fr[0].length; k++) {
              if (fr[i][k] != "") {  arr[i][3] = fr[i][k];  } // time
            }
          } else {
            for (k = 8; k < fr[0].length; k++) {
              if (fr[i][k] != "") {  arr[i][3] = fr[i][k];  } // time
            }
          }
        }
      }
      if (arr[i][3] == "Sheet" && e == "Trigger") { return; }
      else if (arr[i][3] == "Sheet") { ui.alert('Error!', 'No time was assigned to ' + arr[i][0] + ' on the Form Responses sheet. Please manually enter a time to continue!', ui.ButtonSet.OK); return; }
      if ((arr[i][3] == undefined || arr[i][3] == null || arr[i][3] == "") && e == "Trigger") { return; }
      else if (arr[i][3] == undefined || arr[i][3] == null || arr[i][3] == "") { ui.alert('Error!', 'The sheet ' + arr[i][2] + ' was not found. Please correct the date for ' + arr[i][0] + ' to continue!', ui.ButtonSet.OK); return; }
    }
    //Logger.log(arr);
    
    // go through each element in the name array
    for (i = 0; i < sheetNames.length; i++) {
      // erase everything on the current range (current day)
      // should erase all sheets once before reassignment
      if (sheetNames[i] != "Master" && sheetNames[i] != "Form Responses 1") {  // delete everything but master and F.R. 1
        range = ss.getSheetByName(sheetNames[i]).getRange(12, 1, 18, 7).getValues(); // grab entire section of times on whatever day it is
        for (j = 0; j < range.length; j++) {
          range[j][0] = "";  range[j][1] = "";  range[j][5] = "";  range[j][6] = "";
        }
        //Logger.log("ITERATION i = "+i+" sheetNames["+i+"] = " + sheetNames[i]);
        
        // Pull in this sheet. Check date pull in date sheet and find all times, 
        //for each time search through form responses sheet for date/time matches and add them
        for (k = 0; k < arr.length; k++) {
          if (arr[k][2] == sheetNames[i]) {
            // grab all specific times for that day
            if (arr[k][1] == "Sales / Genius") { times = [range[0][2], range[6][2], range[12][2]]; col = 0; }
            else { times = [range[0][4], range[6][4], range[12][4]]; col = 5; }
            timeRows = [0, 6, 12, range.length];
            //Logger.log("times before format = " + times);
            for (j = 0; j < times.length; j++) { // make sure times are in correct format
              //Logger.log("make sure times are in correct format"); Logger.log("times[j]="+times[j]);
              temp = times[j].toLocaleTimeString().split(" MST");
              temp = temp[0].split(":")[0] + ":" + temp[0].split(":")[1] + " " + temp[0].split(" ")[1];
              times[j] = temp;
            }
            //Logger.log("times after format = " + times);
            //Logger.log("outside of times & arr for loop");
            check = false; full = false;
            for (j = 0; j < times.length && !check && !full; j++) {
              //Logger.log("arr[k][2]="+arr[k][2] + "     sheetNames[i]="+sheetNames[i]);
              //Logger.log("arr[k][3]="+arr[k][3] + "     times[j]="+times[j]);
              if (arr[k][3] == times[j]) {
                //Logger.log("date and time matched"); Logger.log("range["+j+"]["+k+"]="+range[j][k]+"   arr["+k+"][0]="+arr[k][0]);
                for (m = timeRows[j]; m < timeRows[j+1]-1 && !check; m++) {
                  if (!check && (range[m][col] == undefined || range[m][col] == null || range[m][col] == "")) { range[m][col] = arr[k][0]; check = true; }
                  else if (!check && (range[m][col+1] == undefined || range[m][col+1] == null || range[m][col+1] == "")) { range[m][col+1] = arr[k][0]; check = true; }
                }
                if (!check) { full = true; }
              }
            }
            if (full && e == "Trigger") { return; }
            else if (full) { ui.alert('Error!', 'The time "' + arr[k][3] + '" for ' + arr[k][1] + ' in sheet "' + sheetNames[i] + '" is full. Please asign a new time for ' + arr[k][0] + ' in "Form Responses 1" sheet or have them modify their form response.', ui.ButtonSet.OK); return; }
            else if (!check && !full && e == "Trigger") { return; }
            else if (!check && !full) { ui.alert('Error!', 'The time "' + arr[k][3] + '" for ' + arr[k][0] + ' could not be found for ' + arr[k][1] + ' in sheet ' + sheetNames[i] + '. Please correct the time in "Form Responses 1" sheet.', ui.ButtonSet.OK); return; }
          }
        }
        //Logger.log("times="+times);
        ss.getSheetByName(sheetNames[i]).getRange(12, 1, 18, 7).setValues(range); // push values back to the page
      } 
    } // end name for
  }
  //Logger.log("finished function");
}

// make a copy of master - for new training days
function newSheet() {
  // created by Sean Lowe, 7/1/18
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet;
  var name = ui.prompt("Name of New Sheet", "Enter the date of the sheet you are creating (ex. January 1)", ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() == ui.Button.OK) {
    //var name = "test"; // uncomment this and the line under it for testing purposes
    //sheet.setName(name);
    sheet = ss.getSheetByName("Master").copyTo(ss).setName(name.getResponseText());
    var stuffs = sheet.getRange(10, 1, 1, 6).getValues();
    stuffs[0][0] = name.getResponseText(); stuffs[0][5] = name.getResponseText();
    sheet.getRange(10, 1, 1, 6).setValues(stuffs);
    ss.setActiveSheet(sheet);
  }
}
// a b c d e f g h i j  k  l  m  n  o  p  q  r  s  t  u  v  w  x  y  z 
// 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26
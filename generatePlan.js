var rowsPerDay = 7;

function getNoClassDays() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("No Class Days");
  var rangeString = "A" + (sheet.getFrozenRows() + 1) + ":D" + sheet.getLastRow();
  var data = sheet.getRange(rangeString).getDisplayValues();
  var noClassDays = [];
  for (var x=0;x<data.length;x++) {
    var noClassDay = {};
    noClassDay.date = new Date(data[x][3],data[x][2]-1,data[x][1]);
    noClassDay.name = data[x][0];
    noClassDay.type = "no class"
    noClassDays.push(noClassDay);
  }
  return noClassDays;
}


function getCourseDates(sheet) {
  var data = sheet.getRange("B2:G2").getDisplayValues();
  var courseDates = {};
  courseDates["start"] = new Date(data[0][0],data[0][1]-1,data[0][2]);
  courseDates["end"] = new Date(data[0][3],data[0][4]-1,data[0][5]);
  return courseDates;
}


function getCourseName(sheet) {
  var courseCode = sheet.getRange("A2:A2").getDisplayValue();
  return courseCode;
}
  

function getMilestones(courseDates) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Milestones");
  var dataRange = sheet.getRange("A" + (sheet.getFrozenRows() + 1) + ":G" + sheet.getLastRow());
  var data = dataRange.getDisplayValues();
  var milestones = [];
  for (var x=0;x<data.length;x++) {
    if (data[x][1] != "Note") {
      var milestone = {};
      var milestoneDate = new Date(courseDates.start);
      var daysToAdd = (Number(data[x][2]-1) * 7) + Number(data[x][3]-1);
      milestoneDate.setDate(milestoneDate.getDate() + daysToAdd);
      milestone.date = milestoneDate;
      milestone.name = data[x][0];
      milestone.type = data[x][1];
      milestone.notes = data[x][5];
      if (data[x][6]) {
        milestone.materials = data[x][6];
      }
      if (data[x][1] === "Unit Start") {
        milestone.colors = {"foreground": dataRange.getCell(x+1,2).getFontColor(), "background": dataRange.getCell(x+1,2).getBackground()};
      }
      milestones.push(milestone);
    }
  }
  return milestones;
}


function setColors(sheet, colors) {
  for (i=0;i<colors.length;i++) {
    var currentRange = sheet.getRange("A" + colors[i].row + ":E" + colors[i].row);
      currentRange.setBackground(colors[i].background);
      currentRange.setFontColor(colors[i].foreground);
    if (colors[i].unit) {
      currentRange = sheet.getRange("A" + colors[i].row + ":A" + sheet.getMaxRows());
      currentRange.setBackground(colors[i].background);
      currentRange.setFontColor(colors[i].foreground);
    }
    else if (colors[i].day) {
      var currentRange = sheet.getRange("B" + colors[i].row + ":E" + colors[i].row);
      currentRange.setBackground(colors[i].background);
      currentRange.setFontColor(colors[i].foreground);
    }
  }
}


function main() {
  var infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Course Information");
  var courseName = getCourseName(infoSheet);
  var courseDates = getCourseDates(infoSheet);
  var dates = getNoClassDays().concat(getMilestones(courseDates));
  dates.sort(function(a, b){
    var keyA = new Date(a.date),
        keyB = new Date(b.date);
    if(keyA < keyB) return -1;
    if(keyA > keyB) return 1;
    return 0;
  });

  var numberOfDays = Math.round((courseDates.end - courseDates.start)/(1000*60*60*24)) + 1;
  var dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var monthNames = ["January", "February", "March", "April", "May", "June",  "July", "August", "September", "October", "November", "December"];

  var planValues = [];
  planValues.push(["Date", "Time", "Activity", "Notes", "Materials"]);
  var planColors = [{"row": 1, "foreground": "#006CD9", "background": "#ffa040"}];

  var currentFG = null;
  var currentBG = null;
  for (x=0; x<numberOfDays; x++) {
    var tempDate = new Date(courseDates.start);
    tempDate.setDate(tempDate.getDate() + x);

    if (tempDate.getDay() > 0 && tempDate.getDay() < 6) {
      while (dates[0].date < tempDate) {
        dates.shift();
      }

      planColors.push({"row": planValues.length + 1, "day": true, "foreground": currentFG, "background": currentBG});
  
      var currentEvent = null;
      var headerDetails = [tempDate.toLocaleDateString('en-CA',{ year: 'numeric', month: 'long', day: 'numeric' }),
                           '=SUM(B' + (planValues.length+2) + ':B' + (planValues.length+rowsPerDay) + ')&" minutes"',
                           "","",""]
      //Logger.log(tempDate + " " + x + " " + tempDate.getDay() + " " + headerDetails);
      var eventsThisDay = [headerDetails,];

      while (dates[0] && dates[0].date.getDate() === tempDate.getDate()) {
        currentEvent = dates.shift();
        if (currentEvent.type === "no class"){
          eventsThisDay[0][1] = "";
          eventsThisDay[0][2] = currentEvent.name;
          planColors.push({"row": planValues.length + 1, "unit": false, "foreground": "#000000", "background": "#FF5050"})          
        }
        else if (currentEvent.type === "Unit Start") {
          eventsThisDay[0][2] = currentEvent.name + " Unit Start";
          currentFG = currentEvent.colors.foreground;
          currentBG = currentEvent.colors.background;
          planColors.push({"row": planValues.length + 1, "unit": true, "foreground": currentFG, "background": currentBG});
        } 
        else {
          var eventDetails = [];
          eventDetails[0] = "";
          eventDetails[1] = "";
          if (currentEvent.type === "Formative" || currentEvent.type === "Summative") {
            eventDetails[2] = currentEvent.name + " due";
          }
          else {
            eventDetails[2] = currentEvent.name;
          }
          eventDetails[3] = currentEvent.notes;
          if ("materials" in currentEvent) {
            eventDetails[4] = currentEvent.materials;
          }
          else {
            eventDetails[4] = "";
          }
          eventsThisDay.push(eventDetails);
        }
      }
  
      if (currentEvent === null || currentEvent.type !== "no class") {
        while (eventsThisDay.length < rowsPerDay) {
            eventsThisDay.push(["","","","",""]);
          }
      }
      planValues = planValues.concat(eventsThisDay);
    }
  }

  var planName = courseName + " Forward Plan (" + monthNames[courseDates.start.getMonth()] + " " + courseDates.start.getFullYear() + "-" + monthNames[courseDates.end.getMonth()] + " " + courseDates.end.getFullYear() + ")";
  var forwardPlan = SpreadsheetApp.create(planName, planValues.length, 5);
  var fpSheet = forwardPlan.getSheets()[0];
  fpSheet.setName("Forward Plan");
  fpSheet.setFrozenRows(1);
  setColors(fpSheet, planColors);
  var wholePlanRange = fpSheet.getRange("A1:E" + planValues.length);
  wholePlanRange.setValues(planValues);
  wholePlanRange.setVerticalAlignment("middle");
  var dateRange = fpSheet.getRange("A:A"); 
  dateRange.setNumberFormat("dddd, mmmm d");
  var notesRange = fpSheet.getRange("D:D");
  notesRange.setWrap(true);
  fpSheet.autoResizeColumn(1);
  fpSheet.autoResizeColumn(3);
  fpSheet.setColumnWidth(4,500);
  fpSheet.setColumnWidth(5,400);
}

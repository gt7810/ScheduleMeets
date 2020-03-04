function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Create Schedule')
      .addItem('Normal Blocks', 'createSchedule')
      .addSeparator()
      .addToUi();
  ui.createMenu('Create Events')
      .addItem('Tomorrow', 'createEvents')
      .addSeparator()
      .addToUi();
}

/* Step 1: Create the list of subject events to be created */
/* Current setting: create for tomorrow */
function createSchedule(){
  /* delete previous schedule */
  clearRangeExisting('Daily Schedule');
  
  /* Set date manually start 
  var cdate = new Date('2020-02-08');
  var day = getDayMapping(cdate); 
  Set date manually end */
  
  /* Current setting: tomorrow start */
  var today = new Date();
  var tomorrow = new Date();
  tomorrow.setDate(today.getDate()+1);
  var day = getDayMapping(tomorrow);
  /* Current setting: tomorrow end */
  
  if(day > 0){
    var connection = Jdbc.getConnection("jdbc:mysql://8.8.8.8.80:3306/DatabaseName", "DatabaseUser", "DatabasePassword"); /* connect to database holding timetable data */
    var SQLstatement = connection.createStatement();
    var classlist = SQLstatement.executeQuery("SELECT distinct  subj_class,  teacher_key,  teacher_email, day, period  FROM Timtable table");
    
    /* Create spreadsheet for events to be created */
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Daily Schedule");
    var cell = sheet.getRange('A2');
    
    /* put data to array */
    var scheduleSet = []; 
    var row = 0;
    while(classlist.next()){
      var subject = classlist.getString(1);
      var teacher = classlist.getString(2);
      var email = classlist.getString(3);
      var period = classlist.getString(5);
      var start = getBlockStart(period);
      var end = getBlockEnd(period);
      var requestID = generateRequestID();
      
      /* Temporarily store data for merge */
      scheduleSet.push({'subject': subject, 'start': start, 'end': end, 'teacher':teacher, 'email': email, 'requestID': requestID, 'period': period});
      row++;
    }
    SQLstatement.close();
    connection.close();
    
    /* Merge subject together for same block and period */ 
    var removedSet = [];
    var arr= scheduleSet, scheduleSetUniq = arr.filter(function (a) {
        var key = a.email + '|' + a.start + '|' + a.end + '|' + a.period; 
        if (!this[key]) {
            this[key] = true;
            return true;
        }else{
          removedSet.push(a);
        }
    }, Object.create(null));
   
    for(i in removedSet){
      var row = removedSet[i];
      var result = scheduleSetUniq.filter(function(v, i) {
        return ((v["teacher"] == row['teacher'] && v["start"] == row['start'] && v["end"] == row['end'] && v["period"] == row['period']));
      })
      /* update scheduleSetUniq */
      result[0]['subject'] = result[0]['subject'] + "|" + row['subject'];
    } 

    var ctr = 0;
    for(i in scheduleSetUniq){
      var row = scheduleSetUniq[i];
      var subject = row['subject'];
      var teacher = row['teacher'];
      var email = row['email'];
      var period = row['period'];
      var start = row['start'];
      var end = row['end'];
      var requestID = row['requestID'];
      /* put data to sheet */
      cell.offset(i, 0).setValue(subject);
      cell.offset(i, 1).setValue(start);
      cell.offset(i, 2).setValue(end);
      cell.offset(i, 3).setValue(teacher);
      cell.offset(i, 4).setValue(email);
      /* request ID must be unique daily */
      cell.offset(i, 5).setValue(requestID);
    }
  }
}

/* Step 2: Create the events based on the Daily Schedule Sheet */
function createEvents(){
  /* retrieve events to be created */
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Schedule");
  
  var startRow = 2; 
  var lastrow = sheet.getLastRow() + 1;
  if(lastrow > 2){
    var dataRange = sheet.getRange(startRow, 1, lastrow-2, 6); 
    var data = dataRange.getValues();
    
    var today = new Date();
    var todaystr = Utilities.formatDate(today, 'Asia/Hong_Kong', 'yyyy-MM-dd');
    var tomorrow = new Date();
    tomorrow.setDate(today.getDate()+1);
    var tomorrowstr = Utilities.formatDate(tomorrow, 'Asia/Hong_Kong', 'yyyy-MM-dd');
    
    for (i in data) {
      var row = data[i]; 
      var event = {
        "summary": row[0],
        "guestsCanInviteOthers": false,
        "description": 'To view all Hangouts for the day, please visit https://meet.google.com/_meet',
        "start": {
          'dateTime': tomorrowstr+"T"+row[1],
          'timeZone': 'Asia/Hong_Kong'
        },
        "end": {
          'dateTime': tomorrowstr+"T"+row[2],
          'timeZone': 'Asia/Hong_Kong'
        },
        "organizer": {
          'email': 'yourcalendaraccount
        },
        "conferenceData": {
          "createRequest": {
            "conferenceSolutionKey": {
              "type": "hangoutsMeet"
            },
            "requestId": row[5]
          }
        },
        'attendees': getAttendees(row[0], row[4])
      };
      try{
        Calendar.Events.insert(event, 'primary', {
          "conferenceDataVersion": 1, "sendUpdates": 'all'}); 
      }catch(e){
      }
      Utilities.sleep(1000); 
    }
  }
}

function getAttendees(course, teacher){
  var courseQry = '';
  if(course.indexOf("|")>-1){
    var res = course.split("|");
    courseQry = "('" + res[0] + "','" + res[1] + "')";
  }else{
    courseQry = "('" + course + "')";
  }
  /* This query pulls the student emails matched by the class codes that they attend for setting up the invites */
  var query = "SELECT distinct studentemail  FROM studentTable join Course table where  coursecode in " + courseQry;
  var resultSet = Jdbc.getConnection("jdbc:mysql://8.8.8.8.80:3306/DatabaseName", "DatabbaseUser", "DatabasePassword").createStatement().executeQuery(query);  
  var userSet = [{'email':teacher}]
  while(resultSet.next()){
    if(resultSet.getString(1) != ''){
      userSet.push({'email': resultSet.getString(1)});
    }
  }
  return userSet;
}


function getDayMapping(cdate){
  /* Retrieve Day Mapping */
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("School Day Calendar Mapping");
 
  var cdatestr = Utilities.formatDate(cdate, 'Asia/Hong_Kong', 'yyyy-MM-dd');
  
  var startRow = 2; 
  var lastrow = sheet.getLastRow() + 1; 
  var dataRange = sheet.getRange(startRow, 1, lastrow-2, 2); 
  var data = dataRange.getValues();
  
  var col=0;
  for (i in data) {
    var row = data[i];
    /* get for tomorrow */
    if(row[0] == cdatestr){
      col = row[1];
      break;
    }    
  }
  return col;
}

function clearRangeExisting(sheet) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheet);
  sheet.getRange('A2:G1000').clearContent();
}

/*
Block 1	08.35 - 09.20
Block 2	09.30 - 10.15
Break
Block 3	10.45 - 11.30
Block 4	11.40 - 12.25
Lunch
Block 5	13.25 - 14:10
Block 6	14.20 - 15.05
*/
function getBlockStart(period){
  var start = '';
  switch(period){
    /* block 1 */
    case '2': 
      start = "08:35:00";
      break;
    /* block 2 */
    case '3': 
      start = "09:30:00";
      break;
    /* block 3 */
    case '5': 
      start = "10:45:00";
      break;
    /* block 4 */
    case '6': 
      start = "11:40:00";
      break;
    /* block 5 */
    case '7': 
      start = "13:25:00";
      break;
    /* block 6 */
    case '9': 
      start = "14:20:00";
      break;
  }
  return start;
}

/*
Block 1	08.35 - 09.20
Block 2	09.30 - 10.15
Break
Block 3	10.45 - 11.30
Block 4	11.40 - 12.25
Lunch
Block 5	13.25 - 14:10
Block 6	14.20 - 15.05
*/
function getBlockEnd(period){
  var end = '';
  switch(period){
    /* block 1 */
    case '2': 
      end = "09:20:00";
      break;
    /* block 2 */
    case '3': 
      end = "10:15:00";
      break;
    /* block 3 */
    case '5': 
      end = "11:30:00";
      break; 
    /* block 4 */
    case '6': 
      end = "12:25:00";
      break; 
    /* block 5 */
    case '7': 
      end = "14:10:00";
      break;
    /* block 6 */
    case '9': 
      end = "15:05:00";
      break;
  }
  return end;
}

function generateRequestID() {
  var requestID = "";
  var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  for(var i=0;i<8;i++){
    requestID += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return requestID;
}


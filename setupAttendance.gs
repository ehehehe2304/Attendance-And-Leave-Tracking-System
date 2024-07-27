function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function createATDform() {
  try {
    // Create the Google Form
    var form = FormApp.create('Attendance Form');
    form.addTextItem().setTitle('Employee ID').setRequired(true);
    form.addTextItem().setTitle('Employee Name').setRequired(true);

    return form;
  } catch (error) {
    Logger.log('Error in create ATD Form: ' + error.message);
    throw new Error('Form creation failed: ' + error.message);
  }
}

function formatATDSheet(sheetId) {
  try {
    // Create the Google Sheet
    var sheet = SpreadsheetApp.openById(sheetId);
    var formSheet = sheet.getSheets()[0];
    formSheet.setName("All Attendance Records");
    var dailyPage = sheet.insertSheet('Daily Attendance');
    
    // Define headers for specific columns
    var headers = ['Employee ID', 'Employee Name'];
    var statusHeader = 'Status';
    var sampleData = [
    ['1001', 'employee01'],
    ['1002', 'employee02'],
    ['1003', 'employee03']
  ];
    // Set headers
    dailyPage.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Insert sample data
    dailyPage.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

    // Determine the last column number based on existing columns
    var dailylastColumn = dailyPage.getLastColumn();
    var dailyStatusColumn = dailylastColumn + 1;

    // Set headers only for the specific columns
    dailyPage.getRange(1, dailyStatusColumn).setValue(statusHeader); // Last column

    // Set the default status as "Absent" for all rows with sample data
    var dailyLastRow = dailyPage.getLastRow();
    // Set the default status as "Absent"
    dailyPage.getRange(2, dailyStatusColumn,sampleData.length,1).setValue("Absent"); 

    return sheet;
  } catch (error) {
    Logger.log('Error in format ATD sheet: ' + error.message);
    throw new Error('ATD Sheet format failed: ' + error.message);
  }
}


function createMonthlyPage(sheet, monthlyPageName) {

  // Check if the sheet already exists
  var monthlyPage = sheet.getSheetByName(monthlyPageName);
  if (!monthlyPage) {
    Logger.log('Creating new sheet: ' + monthlyPageName);
    var headers = ['Employee ID', 'Employee Name'];
    var sampleData = [
      ['1001', 'employee01'],
      ['1002', 'employee02'],
      ['1003', 'employee03']
    ];
    monthlyPage = sheet.insertSheet(monthlyPageName);
    monthlyPage.appendRow(headers);
    
    // Insert sample data
    monthlyPage.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Determine the last column number based on existing columns
    var monthlyLastRow = monthlyPage.getLastRow();
    var monthlyLastColumn = monthlyPage.getLastColumn();
    var presentColumn = monthlyLastColumn + 1;
    var lateColumn = monthlyLastColumn + 2;
    var absentColumn = monthlyLastColumn + 3;
    var leaveColumn = monthlyLastColumn + 4;

    var presentHeader = 'Present';
    var lateHeader = 'Late';
    var absentHeader = 'Absent';
    var leaveHeader = 'Leave';

    // Set headers only for the specific columns
    monthlyPage.getRange(1, presentColumn).setValue(presentHeader);
    monthlyPage.getRange(1, lateColumn).setValue(lateHeader); 
    monthlyPage.getRange(1, absentColumn).setValue(absentHeader);
    monthlyPage.getRange(1, leaveColumn).setValue(leaveHeader);

    // Function to create a 2D array filled with a specific value
    function createZeroArray(rows, cols) {
      var array = [];
      for (var i = 0; i < rows; i++) {
        array.push(new Array(cols).fill(0));
      }
      return array;
    }

    // Get number of rows to update
    var numRows = sampleData.length; // Number of rows
    var numCols = 1; // Only one column per range

    // Create a 2D array with all values set to 0
    var zeroArray = createZeroArray(numRows, numCols);

    // Update the ranges with 0 values
    monthlyPage.getRange(2, presentColumn, numRows, 1).setValues(zeroArray);
    monthlyPage.getRange(2, lateColumn, numRows, 1).setValues(zeroArray);
    monthlyPage.getRange(2, absentColumn, numRows, 1).setValues(zeroArray);
    monthlyPage.getRange(2, leaveColumn, numRows, 1).setValues(zeroArray);
    
    Logger.log('Monthly page created with name: ' + monthlyPageName);
  } else {
    Logger.log('Sheet already exists: ' + monthlyPageName);
  }

  // Verify if the sheet is correctly referenced
  Logger.log('Returning sheet: ' + monthlyPage.getName());
  return monthlyPage;
}






function onATDFormSubmit(e) {
  try {
    Logger.log('Event object: ' + JSON.stringify(e));
    var ss = e.source;
    var sheet = e.source.getSheets()[0];

    var email = Session.getActiveUser().getEmail();

    // Get the range of the edited cell.
    var range = e.range;
    var lastRow = sheet.getLastRow();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('Headers: ' + headers);

    var dailyPage = ss.getSheetByName('Daily Attendance');
    var dailyLastRow = dailyPage.getLastRow();
    var dailyHeaders = dailyPage.getRange(1, 1, 1, dailyPage.getLastColumn()).getValues()[0];
    var dailyIdColumn = dailyHeaders.indexOf('Employee ID') + 1;
    var dailyStatusColumn = dailyHeaders.indexOf('Status') + 1;


     // Find the column indices based on the header names
    var submitTimeColumn = headers.indexOf('Timestamp') + 1;
    var idColumn = headers.indexOf('Employee ID') + 1;
    var id = sheet.getRange(lastRow, idColumn).getValue();

    // check monthly report
    var today = new Date();
    var monthYear = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM");
    var monthlyPageName = "MonthlyReport_" + monthYear;
    
   // Check if the report sheet already exists
    var monthlyPage = createMonthlyPage(ss, monthlyPageName);
    Logger.log('Monthly Page Object: ' + monthlyPage);
    var monthlyLastRow = monthlyPage.getLastRow();
    var monthlyHeaders = monthlyPage.getRange(1, 1, 1, monthlyPage.getLastColumn()).getValues()[0];
    var monthlyIdColumn = monthlyHeaders.indexOf('Employee ID') + 1;
    var monthlyPresentColumn = monthlyHeaders.indexOf('Present') + 1;
    var monthlyLateColumn = monthlyHeaders.indexOf('Late') + 1;
  
    var submitTime = new Date(sheet.getRange(lastRow, submitTimeColumn).getValue());
    // Define the "late" time
    var lateTime = new Date();
    lateTime.setHours(9, 5, 0, 0); //  TIMEEEE

    // Extract time components from submitTime
    var submitTimeHours = submitTime.getHours();
    var submitTimeMinutes = submitTime.getMinutes();
    var submitTimeSeconds = submitTime.getSeconds();

    // Extract time components from lateTime
    var lateHours = lateTime.getHours();
    var lateMinutes = lateTime.getMinutes();
    var lateSeconds = lateTime.getSeconds();

    // Define isLate
    var isLate = 
      (submitTimeHours > lateHours) || 
      (submitTimeHours === lateHours && submitTimeMinutes > lateMinutes) || 
      (submitTimeHours === lateHours && submitTimeMinutes === lateMinutes && submitTimeSeconds > lateSeconds);


    if (isLate) {
      Logger.log('Form submission is late.');
      // Update daily
      var dailyUpdated = false;
      for (var i = 2; i <= dailyLastRow; i++) {
        var dailyId = dailyPage.getRange(i, dailyIdColumn).getValue();
        Logger.log('Daily ID: ' + dailyId);
    
        if (id === dailyId) {
          dailyPage.getRange(i, dailyStatusColumn).setValue('Late');
          Logger.log('Daily Changed at Row ' + i);
          dailyUpdated = true;
          break;
        } 
      }
      if (!dailyUpdated) {
        Logger.log('Late ID not found in Daily Attendance.');
      }

      // Update monthly
      var monthlyUpdated = false;
      for (var i = 2; i <= monthlyLastRow; i++) {
        var monthlyId = monthlyPage.getRange(i, monthlyIdColumn).getValue();
        Logger.log('Monthly ID: ' + monthlyId);
        var monthlyLate = monthlyPage.getRange(i, monthlyLateColumn).getValue();

        if (id === monthlyId) {
          var updateLate = monthlyLate + 1;
          monthlyPage.getRange(i, monthlyLateColumn).setValue(updateLate);
          Logger.log('Monthly Changed at Row ' + i);
          monthlyUpdated = true;
          break;
        }
      }
      if (!monthlyUpdated) {
        Logger.log('Late ID not found in Monthly Report.');
      }
      
      MailApp.sendEmail({
        to: email,
        subject: today + ' Late Notification',
        body: id + ' is LATE.\n\n Just Signed In At : ' + submitTime
      });

    } else {
      Logger.log('Form submission is on time.');
      // Update daily
      var dailyUpdated = false;
      for (var i = 2; i <= dailyLastRow; i++) {
        var dailyId = dailyPage.getRange(i, dailyIdColumn).getValue();
        Logger.log('Daily ID: ' + dailyId);

        if (id === dailyId) {
          dailyPage.getRange(i, dailyStatusColumn).setValue('Present');
          Logger.log('Daily Changed at Row ' + i);
          dailyUpdated = true;
          break;
        } 
      }
      if (!dailyUpdated) {
        Logger.log('IT ID not found in Daily Attendance.');
      }

      // Update monthly
      var monthlyUpdated = false;
      for (var i = 2; i <= monthlyLastRow; i++) {
        var monthlyId = monthlyPage.getRange(i, monthlyIdColumn).getValue();
        Logger.log('Monthly ID: ' + monthlyId);
        var monthlyPresent = monthlyPage.getRange(i, monthlyPresentColumn).getValue();

        if (id === monthlyId) {
          var updatePresent = monthlyPresent + 1;
          monthlyPage.getRange(i, monthlyPresentColumn).setValue(updatePresent);
          Logger.log('Monthly Changed at Row ' + i);
          monthlyUpdated = true;
          break;
        }
      }
      if (!monthlyUpdated) {
        Logger.log('OT ID not found in Monthly Report.');
      } 
    }
    Logger.log('Form submission processed successfully.');
  } catch (error) {
    Logger.log('Error in onFormSubmit: ' + error.message);
  }
}



function resetDailySheetToDefault() {
  try {
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var dailyPage = sheet.getSheetByName('Daily Attendance');
    var dailyLastRow = dailyPage.getLastRow();
    var dailyHeaders = dailyPage.getRange(1, 1, 1, dailyPage.getLastColumn()).getValues()[0];
    var dailyIdColumn = dailyHeaders.indexOf('Employee ID') + 1;
    var dailyStatusColumn = dailyHeaders.indexOf('Status') + 1;

    var today = new Date();
    var monthYear = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM");
    var monthlyPageName = "MonthlyReport_" + monthYear;

    var monthlyPage = sheet.getSheetByName(monthlyPageName);
    var monthlyLastRow = monthlyPage.getLastRow();
    var monthlyHeaders = monthlyPage.getRange(1, 1, 1, monthlyPage.getLastColumn()).getValues()[0];
    var monthlyIdColumn = monthlyHeaders.indexOf('Employee ID') + 1;
    var monthlyLeaveColumn = monthlyHeaders.indexOf('Leave') + 1;
    var monthlyAbsentColumn = monthlyHeaders.indexOf('Absent') + 1;



    for (var i = 2; i <= dailyLastRow; i++) {
      var dailyID = dailyPage.getRange(i, dailyIdColumn).getValue();
      dailyPage.getRange(i, dailyStatusColumn).setValue('Absent');

      for (var j = 2; j <= monthlyLastRow; j++) {
        var monthlyId = monthlyPage.getRange(j, monthlyIdColumn).getValue();
        if (dailyID === monthlyId) {
          var monthlyAbsent = monthlyPage.getRange(j, monthlyAbsentColumn).getValue();
          monthlyPage.getRange(j, monthlyAbsentColumn).setValue(monthlyAbsent + 1);
          break;
        }
      }
    }

    checkLeave();

    for (var j = 2; j <= monthlyLastRow; j++) {
    var monthlyId = monthlyPage.getRange(j, monthlyIdColumn).getValue();
    
      // Loop through the leaveIds to check for a match
      for (var i = 0; i < leaveIds.length; i++) {
        if (leaveIds[i] === monthlyId) {
          var monthlyLeave = monthlyPage.getRange(j, monthlyLeaveColumn).getValue();
          monthlyPage.getRange(j, monthlyLeaveColumn).setValue(monthlyLeave + 1);
          break; // Exit the inner loop once a match is found
        }
      }
    } 

    remindAbsent();
    Logger.log('Sheet reset to default state.');
  } catch (error) {
    Logger.log('Error in resetDailySheetToDefault: ' + error.message);
  }
}

function createTimeDrivenTrigger(atdSheetId) {
  var spreadsheet = SpreadsheetApp.openById(atdSheetId);

  ScriptApp.newTrigger('onATDFormSubmit')
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();
  Logger.log('Form Trigger created successfully.');

  // Create reset daily trigger
  ScriptApp.newTrigger('resetDailySheetToDefault')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(5)
    .create();

  // Create absent notification trigger
  ScriptApp.newTrigger('remindAbsent')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();

  Logger.log('Time-driven trigger created successfully.');
}


function remindAbsent() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var email = Session.getActiveUser().getEmail();
  var dailyPage = sheet.getSheetByName('Daily Attendance');
  var dailyLastRow = dailyPage.getLastRow();
  var dailyHeaders = dailyPage.getRange(1, 1, 1, dailyPage.getLastColumn()).getValues()[0];
  var dailyIdColumn = dailyHeaders.indexOf('Employee ID') + 1;
  var dailyNameColumn = dailyHeaders.indexOf('Employee Name') + 1;
  var dailyStatusColumn = dailyHeaders.indexOf('Status') + 1;
  var today = new Date();

  var absentEntries = [];

  for (var i = 2; i <= dailyLastRow; i++) {
    var dailyID = dailyPage.getRange(i, dailyIdColumn).getValue();
    var dailyName = dailyPage.getRange(i, dailyNameColumn).getValue();
    var status = dailyPage.getRange(i, dailyStatusColumn).getValue();

    if (status === 'Absent') {
      absentEntries.push({
        id: dailyID,
        name: dailyName,
        row: i
      });
    }
  }

  // Prepare HTML table for email body
  var htmlTable = '<table border="1" style="border-collapse: collapse; width: 100%;">' + '<tr><th>Employee ID</th><th>Employee Name</th></tr>';
  
  if (absentEntries.length > 0) {
    absentEntries.forEach(function(entry) {
      htmlTable += '<tr>' +
                    '<td>' + entry.id + '</td>' +
                    '<td>' + entry.name + '</td>' +
                    '</tr>';
    });
    htmlTable += '</table>';
  }

  MailApp.sendEmail({
    to: email,
    subject: today +' Absent Notification',
    htmlBody: 'Dear User,Here is the Employee Absent List for today on '+ today + ': <br><br>' + htmlTable 
  });
}

function checkLeave(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var leaveSheet = sheet.getSheetByName('All Leave Requests');
  var leaveLastRow = leaveSheet.getLastRow();
  var leaveHeaders = leaveSheet.getRange(1, 1, 1, leaveSheet.getLastColumn()).getValues()[0];
  var leaveIdColumn = leaveHeaders.indexOf('Employee ID') + 1;
  var leaveStatusColumn = leaveHeaders.indexOf('Status') + 1;
  var leaveStartDateColumn = leaveHeaders.indexOf('Start Date') + 1;
  var leaveEndDateColumn = leaveHeaders.indexOf('End Date') + 1;

  var leaveIds = [];

  for (var k = 2; k <= leaveLastRow; k++) {
  var leaveId = leaveSheet.getRange(k, leaveIdColumn).getValue();
  var leaveStatus = leaveSheet.getRange(k, leaveStatusColumn).getValue();
  var leaveStartDate = new Date(leaveSheet.getRange(k, leaveStartDateColumn).getValue());
  var leaveEndDate = new Date(leaveSheet.getRange(k, leaveEndDateColumn).getValue());

    // Check if the leave is approved and if today's date falls within the leave period
    if (leaveStatus === 'approved' && (today >= leaveStartDate && today <= leaveEndDate)) {
      // Add the leave ID to the list
      leaveIds.push(leaveId);
    }
  }
  Logger.log('List of IDs who applied for leave: ' + leaveIds);
  return leaveIds;
}

      





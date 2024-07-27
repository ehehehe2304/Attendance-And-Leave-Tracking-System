


function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function setupFolder() {
  try {
    Logger.log('Starting setupFolder');
    var folder = DriveApp.createFolder('Leave Application');
    var folderurl = folder.getUrl();
    Logger.log('Folder created: ' + folder.getUrl());
    
    // create form and sheet for leave
    var leaveForm = createLeaveForm();
    Logger.log('Form created: ' + leaveForm.getId());
    var leaveSheet = SpreadsheetApp.create('Leave Application Record');
    var leaveSheetId = leaveSheet.getId();
    Logger.log('Sheet created: ' + leaveSheet.getId());
    
    // Link the leave form to the leave sheet
    leaveForm.setDestination(FormApp.DestinationType.SPREADSHEET, leaveSheet.getId());
    Logger.log('Leave Form linked to sheet');


    // create form and sheet for attendance
    var atdForm = createATDform();
    Logger.log('Form created: ' + atdForm.getId());
    
    var atdSheet = SpreadsheetApp.create('Attendance Record');
    var atdSheetId = atdSheet.getId();
    Logger.log('Sheet created: ' + atdSheet.getId());

    // Link the AT form to the AT sheet
    atdForm.setDestination(FormApp.DestinationType.SPREADSHEET, atdSheet.getId());
    Logger.log('ATForm linked to sheet');


    DriveApp.getFileById(leaveForm.getId()).moveTo(folder);
    DriveApp.getFileById(leaveSheet.getId()).moveTo(folder);
    DriveApp.getFileById(atdForm.getId()).moveTo(folder);
    DriveApp.getFileById(atdSheet.getId()).moveTo(folder);

    Logger.log('Files moved');
    
    Logger.log('Folder URL: ' + folderurl);
    Logger.log('Leave Sheet ID: ' + leaveSheetId);
    Logger.log('ATD Sheet ID: ' + atdSheetId);
    return {
      folderurl: folderurl,
      leaveSheetId: leaveSheetId,
      atdSheetId: atdSheetId
    };

  } catch (error) {
    Logger.log('Error in setupFolder: ' + error.message);
    throw new Error('Setup failed: ' + error.message);
  }
}



function createLeaveForm() {
  try {
    // Create the Google Form
    var form = FormApp.create('Leave Application Form');
    form.setCollectEmail(true); 
    form.addTextItem().setTitle('Employee ID').setRequired(true);
    form.addTextItem().setTitle('Employee Name').setRequired(true);

    var leaveTypeItem = form.addMultipleChoiceItem()
      .setTitle('Type of Leave')
      .setChoiceValues(['Annual Leave', 'Sick Leave', 'Maternity Leave', 'Personal Leave', 'Unpaid Leave'])
      .showOtherOption(true) // This line enables the 'Other' option with text input
      .setRequired(true);

    form.addDateItem().setTitle('Start Date').setRequired(true);
    form.addDateItem().setTitle('End Date').setHelpText('If only 1 day, can leave it blank or select the same date.').setRequired(false);

    return form;
  } catch (error) {
    Logger.log('Error in createForm: ' + error.message);
    throw new Error('Form creation failed: ' + error.message);
  }
}

function formatLeaveSheet(sheetId) {
  try {
    // Create the Google Sheet
    var sheet = SpreadsheetApp.openById(sheetId);
    var formSheet = sheet.getSheets()[0];
    formSheet.setName("All Leave Requests");

    // Define headers for specific columns
    var statusHeader = 'Status';
    var daysHeader = 'Days';

    // Determine the last column number based on existing columns
    var lastColumn = sheet.getLastColumn();
    var statusColumn = lastColumn + 2;
    var durationDayColumn = lastColumn + 1; 

    // Set headers only for the specific columns
    formSheet.getRange(1, durationDayColumn).setValue(daysHeader); // Second to last column
    formSheet.getRange(1, statusColumn).setValue(statusHeader); // Last column


    var range = formSheet.getRange(2,statusColumn,formSheet.getMaxRows() - 1,1);
    Logger.log('Status Column Range: ' + range.getA1Notation());
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Approved', 'Rejected'])
      .setAllowInvalid(false)
      .build();
    range.setDataValidation(rule);

    return sheet;
  } catch (error) {
    Logger.log('Error in createSheet: ' + error.message);
    throw new Error('Sheet creation failed: ' + error.message);
  }
}

function applyConditionalFormatting(sheetId) {
  try {
    var sheet = SpreadsheetApp.openById(sheetId);
    var sheetPage = sheet.getSheets()[0];
    var headers = sheetPage.getRange(1, 1, 1, sheetPage.getLastColumn()).getValues()[0];
    var statusColumn = headers.indexOf('Status') + 1;

    // Check if the 'Status' column is valid
    if (statusColumn <= 0) {
      throw new Error('Status column not found or invalid column index.');
    }

    var statusColumnRange = sheetPage.getRange(2, statusColumn, sheetPage.getMaxRows() - 1, 1);
    
    // Define conditional formatting rules
    var rulePending = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending')
      .setBackground('#FFFFCC') // Light yellow
      .setRanges([statusColumnRange])
      .build();
    
    var ruleRejected = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Rejected')
      .setBackground('#FFCCCC') // Light red
      .setRanges([statusColumnRange])
      .build();

    var ruleApproved = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Approved')
      .setBackground('#CCFFCC') // Light green
      .setRanges([statusColumnRange])
      .build();

    // Apply the rules
    var rules = sheetPage.getConditionalFormatRules();
    rules.push(rulePending, ruleRejected, ruleApproved);
    sheetPage.setConditionalFormatRules(rules);

    

  } catch (error) {
    Logger.log('Error in applyConditionalFormatting: ' + error.message);
    throw new Error('Conditional formatting failed: ' + error.message);
  }
}


function onLeaveFormSubmit(e) {
  try {
    Logger.log('Event object: ' + JSON.stringify(e));
    
    var sheet = e.source.getSheets()[0]; 
    Logger.log('Sheet obtained: ' + sheet.getName());
    
    var lastRow = sheet.getLastRow();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('Headers: ' + headers);
 
    // Find the column indices based on the header names
    var date1Column = headers.indexOf('Start Date') + 1;
    var date2Column = headers.indexOf('End Date') + 1;

    var date1 = new Date(sheet.getRange(lastRow, date1Column).getValue());
    var date2 = sheet.getRange(lastRow, date2Column).getValue();

    // Calculate the number of days
    var days = 1;
    if (date2) {
      date2 = new Date(date2);
      var diffTime = Math.abs(date2 - date1);
      days = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // Adding 1 to include the start date
    }

    // Set the day count
    var daysColumn = headers.indexOf('Days') + 1;
    sheet.getRange(lastRow, daysColumn).setValue(days);

    // Set the default status as "Pending"
    var statusColumn = headers.indexOf('Status') + 1;
    sheet.getRange(lastRow, statusColumn).setValue("Pending"); 


    // seperate month
      var month = date1.getMonth() + 1; // JavaScript months are 0-11, so add 1
      var year = date1.getFullYear();
    
      // Create the sheet name in the format "YYYY-MM"
      var sheetName = year + '-' + (month < 10 ? '0' + month : month);
    
      // Get or create the monthly sheet
      var monthlySheet = e.source.getSheetByName(sheetName);
      if (!monthlySheet) {
        // If the sheet does not exist, create it and add the header
        monthlySheet = e.source.insertSheet(sheetName);
        monthlySheet.appendRow(headers);
      }

      var row = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      // Append the new row to the monthly sheet
      monthlySheet.appendRow(row);

    Logger.log('Form submission processed successfully.');
  } catch (error) {
    Logger.log('Error in onFormSubmit: ' + error.message);
  }
}

function createTrigger(sheetId) {
  try {
    // Check if sheetId is provided and is a valid string
    if (!sheetId || typeof sheetId !== 'string') {
      throw new Error('Invalid spreadsheet ID.');
    }

    // Get existing triggers and delete them
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    Logger.log('All triggers deleted.');
    
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    
    if (!spreadsheet) {
      throw new Error('No active spreadsheet found.');
    }
    
    Logger.log('Active Spreadsheet: ' + spreadsheet.getId());
    
    ScriptApp.newTrigger('onLeaveFormSubmit')
      .forSpreadsheet(spreadsheet)
      .onFormSubmit()
      .create();
    Logger.log('Form Trigger created successfully.');

    ScriptApp.newTrigger('onLeaveSheetEdit')
      .forSpreadsheet(spreadsheet)
      .onEdit()
      .create();
    Logger.log('Sheet Edit Trigger created successfully.');

  } catch (error) {
    Logger.log('Error in createFormSubmitTrigger: ' + error.message);
    throw new Error('Trigger creation failed: ' + error.message);
  }
}


function onLeaveSheetEdit(e) {
  try {
    // Ensure the event object is valid
    if (!e || !e.source || !e.range) {
      throw new Error('Invalid event object.');
    }

    // Get the sheet where the edit was made.
    var sheet = e.source.getActiveSheet(); ;
    // Get the range of the edited cell.
    var range = e.range;
    
    // Get the headers of the sheet.
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Find the index of the 'Status' column.
    var statusColumn = headers.indexOf('Status') + 1;
    // Find the index of the 'Email' column.
    var emailColumn = headers.indexOf('Email Address') + 1;
    // Find the index of the 'Start Date' column.
    var startDateColumn = headers.indexOf('Start Date') + 1;
    var idColumn = headers.indexOf('Employee ID') + 1;
    var nameColumn = headers.indexOf('Employee Name') + 1;


    if (statusColumn < 1 || emailColumn < 1 || startDateColumn < 1) {
      throw new Error('One or more required columns are missing.');
    }


    // Check if the edited cell is in the 'Status' column.
    if (range.getColumn() == statusColumn) {

      // Get the row number of the edited cell.
      var row = range.getRow();
      // Get the updated status value.
      var status = sheet.getRange(row, statusColumn).getValue();
      // Get the email address from the 'Email' column.
      var email = sheet.getRange(row, emailColumn).getValue();
      // Get the date from the 'Start Date' column.
      var date1 = new Date(sheet.getRange(row, startDateColumn).getValue());
      var id = sheet.getRange(row, idColumn).getValue();
      var name = sheet.getRange(row, nameColumn).getValue();
      Logger.log('Edited Status : ' + status);


      // Extract the month and year from the date.
      var month = date1.getMonth() + 1; // Months are 0-based in JavaScript.
      var year = date1.getFullYear();
      // Create the sheet name in the format "YYYY-MM".
      var sheetName = year + '-' + (month < 10 ? '0' + month : month);
      

      // Get the monthly sheet with the corresponding name.
      var monthlySheet = e.source.getSheetByName(sheetName);
      if (monthlySheet) {
        // Find the last row in the monthly sheet.
        var monthlyLastRow = monthlySheet.getLastRow();
        // Get headers from the monthly sheet.
        var monthlyHeaders = monthlySheet.getRange(1, 1, 1, monthlySheet.getLastColumn()).getValues()[0];
        // Find the index of the 'Status' column in the monthly sheet.
        var monthlyStatusColumn = monthlyHeaders.indexOf('Status') + 1;
        var monthlyIdColumn = monthlyHeaders.indexOf('Employee ID') + 1;
        var monthlyNameColumn = monthlyHeaders.indexOf('Employee Name') + 1;
        

        // Iterate through rows in the monthly sheet to find and update the corresponding row.
        var updated = false;
        for (var i = 2; i <= monthlyLastRow; i++) {
          // Get the status value from the monthly sheet.
          var existingStatus = monthlySheet.getRange(i, monthlyStatusColumn).getValue();
          var monthlyID = monthlySheet.getRange(i, monthlyIdColumn).getValue();
          var monthlyName = monthlySheet.getRange(i, monthlyNameColumn).getValue();
          
          // Log the existing status
          Logger.log('Monthly Status: ' + existingStatus);

          // Check if the status matches.
          if (id === monthlyID && name === monthlyName &&status !== existingStatus) {
            // Update the row in the monthly sheet with new values from the main sheet.
            monthlySheet.getRange(i, 1, 1, sheet.getLastColumn()).setValues([sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]]);
            Logger.log('Monthly Changed ');

            // Send an email notification if the status has changed
            if (email) {
              MailApp.sendEmail({
                to: email,
                subject: 'Leave Application Status Update',
                body: 'The status of your leave application has been updated to: ' + status + '.\n\nAny inquiries please contact HR.'
              });
            }
            Logger.log('Status updated and email sent to: ' + email);
            updated = true;
            break; // Exit the loop once the row is updated.
          }
        }
      }
    }
  } catch (error) {
    // Log any errors that occur during execution.
    Logger.log('Error in onEdit: ' + error.message);
  }
}
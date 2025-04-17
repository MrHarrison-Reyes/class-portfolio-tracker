/**
 * Before creating the following script, your Google Sheet must consist of the following tabs:
 * Class Attendance Tracker
 * Math Portfolio Progress
 * Then a sheet for each individual student with student email and parent email
 * I have created a template here: https://docs.google.com/spreadsheets/d/1esrAGQllzy7zVb60ldvzii7Ay4WwL5G21e3FoyIadhQ/edit?usp=sharing
 */


function sendMonthlySummaryEmails() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Open the active spreadsheet
    var sheets = spreadsheet.getSheets(); // Get all sheets
    var attendanceSheet = sheets[0]
  
    var today = new Date();
  
     // Start the loop from the fourth sheet (index 3)
    for (var i = 2; i < sheets.length; i++) {
      var sheet = sheets[i]; // Get the current sheet
  
      Logger.log("Processing sheet: " + sheet.getName()); // Log the sheet being processed
  
      // Extract student and parent details
      var studentName = sheet.getRange('B1').getValue(); // Get the student's name based on individual sheet and cell
      var studentEmail = sheet.getRange('B6').getValue(); // Get the student's email based on individual sheet and cell
      var parentEmail = sheet.getRange('B7').getValue(); // Get the parent's email based on individual sheet and cell
  
      // Skip the sheet if required data is missing
      if (!studentName || !studentEmail) {
        Logger.log("Skipping sheet: " + sheet.getName() + " due to missing data.");
        continue; // Continue to the next sheet
      }
  
      // Get test data and generate the summary
      var completedTests = getTestCompleted(sheet);
      var completionPercentage = getTestCompletionPercentage(sheet, completedTests);
      var studentAttendance = getStudentAttendance(studentName, attendanceSheet);
      var summary = summaryOfStudent(completedTests, completionPercentage, sheet.getName(), studentAttendance, today.getMonth());
  
       
      Logger.log(summary)
    
    // Send the email
    // emailTo(studentName, studentEmail, parentEmail, summary, getCurrentMonthText(today.getMonth()));
  
    }
  }
  
  // Get the task list information: test name and has it been completed
  function getTestCompleted(sheet){
    var range = sheet.getRange('A8').offset(0, 0, 46, 2);   // Get the range of cells
    var testsGroup = range.getValues();    // Get values as a 2D array
    var completedTest = []
  
    // Add the completed test into the array of completedTest
    testsGroup.forEach((test) => {
      if(test[1] === "High Competence" || test[1] === "Standard Competence"){
        var nameOfTest = test[0].split(" - ")
        completedTest.push(nameOfTest[1])
      } 
    })
  
    // Log confirmation
    Logger.log("Obtaining the list of test completed.");
  
    return completedTest
  }
  
  // Gets the percentage of test completed by student
  function getTestCompletionPercentage(sheet, completedTests){
    var range = sheet.getRange('A8').offset(0, 0, 47, 2);   // Get the range of cells
    var testsGroup = range.getValues();    // Get values as a 2D array
    Logger.log("Obtaining the percentage.");
    var convertedToPercent = Math.round(completedTests.length/(testsGroup.length -1) * 100) + "%"  
    return convertedToPercent
  }
  
  // Create a summary string for the completed tests
  function summaryOfStudent(completedTests, percentageOfCompletion, sheetName, studentAttendance, currentMonth){
  
    var testSummary = completedTests.length > 0 
      ? "Completed Tests: " + completedTests.join(", ") 
      : "No tests have been completed.";
  
    // Add the percentage of completion to the summary
    var percentageSummary = "Total Percentage of Completion: " + percentageOfCompletion;
  
    // Number of absences
    var absences = "Number of Absences for the month of " + getCurrentMonthText(currentMonth) + ": " + getStudentAbsenses(studentAttendance, currentMonth) + " classes."
  
    // Combine the summaries
    var summary = testSummary + "\n" + percentageSummary + "\n" + absences;
  
    // Log confirmation
    Logger.log("Summary is completed for sheet " + sheetName);
  
    return summary
  }
  
  // An email to send to student and parent the summary of work.
  function emailTo(studentName, studentEmail, parentEmail, summary, currentMonth) {
  
    // Define the subject and body of the email
    var subject = "Math MCAS Completion Summary: " + currentMonth;
    var body = "Hello!,\n\n" +
               "It is [INSERT TEACHER's NAME], your child's Portfolio teacher. I hope you are doing well. I will be sending out monthly reports on your scholar's progress. Here is summary of your scholar, " + studentName + ", for " + currentMonth + ": \n\n" +
               summary + "\n\n" +
               'If you have any questions, please contact me via email or call/text [INSERT PHONE NUMBER]. \n\n' +
               "Best regards,\n[INSERT TEACHER NAME]";
  
     // Define the email options with CC
    var emailOptions = {
      cc: studentEmail // Add CC recipient
    };
  
    // Send the email
    GmailApp.sendEmail(parentEmail, subject, body, emailOptions);
  
    // Log confirmation
    Logger.log("Email sent to " + parentEmail + " for student " + studentName + ".");
  }
  
  // Get the current month as text
  function getCurrentMonthText(today) {
     // Get the current date
    var monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    var currentMonthText = monthNames[today]; // Use getMonth() to index the array
    return currentMonthText; // Returns the current month as text
  }
  
  // Locate the column the student is in.
  function findStudentColumn(studentName, attendanceSheet) {
    var headers = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0]; // Get the header row as a 1D array
  
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === studentName) {
        Logger.log("Found '" + studentName + "' in column: " + (i + 1)); // Column numbers are 1-based
        return i + 1; // Return the column index (1-based)
      }
    }
  
    Logger.log("Student '" + studentName + "' not found.");
    return -1; // Return -1 if not found
  }
  
  // Get the attendance record of a student
  function getStudentAttendance(studentName, sheet){
  
    // Locate the student's column
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // First row
    var studentColumn = -1;
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === studentName) {
        studentColumn = i + 1; // Convert to 1-based index
        Logger.log("Student found: " + studentName);
        break;
      }
    }
    if (studentColumn === -1) {
      Logger.log("Student not found: " + studentName);
      return []; // Exit if student is not found
    }
  
    // Fetch the attendance data and dates
    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, studentColumn).getValues(); // Get data from row 2 onward
    var attendanceData = [];
    dataRange.forEach(function(row) {
      var date = row[0]; // Date is in column 1
      var date = new Date(date);
      var month = date.getMonth() // months counting starts at 0
      var status = row[studentColumn - 1]; // Attendance status for the student
      attendanceData.push({ month: month, status: status });
    });
      
    return attendanceData; // Return the data as an array of objects
  }
  
  // Calculate the number of student absences per month
  function getStudentAbsenses(studentAttendance, currentMonth){
    var monthCounter = 0
    var absenceCounter = 0 
  
    studentAbsenses = studentAttendance.forEach( function(record) {
        if ( record.month === currentMonth){
          monthCounter++;
  
         if (record.status === "Absent"){
            absenceCounter++
          } 
        }
    })
  
    return absenceCounter + " out of " + monthCounter;
  }
  
  /** 
   * To send automated emails follow these steps in the appscript editor:
   * 1. On the left select "Triggers, symbol looks like a clock
   * 2. Bottom right of screen, click "+ Add Trigger"
   * 3. "Choose Which Function to Run:" Select "sendMonthlyEmails"
   * 4. "Choose which deployment should run:" Select "Head"
   * 5. "Select Event Source" select "Time-Driven"
   * 6. "Select type of time based trigger", select "Month Timer"
   * 7. "Select Day of Month", select which day you prefer, I prefer 28th of each month to garantee all months during school year are sent.
   * 8. "Select Time of Day", select which time you prefer, I would suggest 4pm or 7AM
   * 9. "Failure Notification Settings", select "immediately"
   * 10. Hit "Save" and give it permissions.
   * You are all set!
  */
  
  
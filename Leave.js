// Link to google form: https://docs.google.com/forms/d/1BzcvINydON1VzyqPZWa4vB5GmHVRn2vyBDGYc0vzeoc/viewform?edit_requested=true#responses
// Link to google sheet: https://docs.google.com/spreadsheets/d/1fq00lgerZGMRl6jNpKNUOwWC9yS-e_OtDLnjuOQdIZM/edit?usp=sharing

function onFormSubmission(e) {

    var sheetName = "Form Responses 1";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var activeRow = e.range.getRow();
  
    // Fetch values directly from columns A to L in a single line
    var [columnA, columnB, columnC, columnD, columnE, columnF, columnG, columnH, 
        columnI, columnJ, columnK, columnL, columnM, columnN, columnO, columnP, columnQ, columnR, columnS, 
        columnT, columnU, columnV, columnW, columnX, columnY, columnZ
    ] = sheet.getRange(activeRow, 1, 1, 26).getValues()[0];
  
    // Calculate the duration between "From Date" (Column K) and "To Date" (Column L)    
    var duration = calculateDateDifference(columnK, columnL);   
    // Set the calculated duration in Column M (index 13)
    sheet.getRange(activeRow, 13).setValue(duration);
  

    //Function for Calculate Date form From Date and To Date
  function calculateDateDifference(startDate, endDate) {
    var startTimestamp = new Date(startDate).getTime();
    var endTimestamp = new Date(endDate).getTime();
    if (isNaN(startTimestamp) || isNaN(endTimestamp) || startTimestamp > endTimestamp) {
      return "Invalid Date";
    }
    var millisecondsInADay = 1000 * 60 * 60 * 24;
    var differenceInDays = Math.floor((endTimestamp - startTimestamp) / millisecondsInADay)+1;
    return differenceInDays;
  }

   
  var leaveBalance = getLeaveBalance(columnH, columnI); // Calculate leave balance based on the leave type (Column )
  // Check if leave duration exceeds leave balance
  if (duration > leaveBalance) {
    // Send one type of email notification (insufficient balance)
    sendInsufficientBalanceEmail(columnH, columnI, duration);
    sheet.getRange(activeRow, 14).setValue("Reject-Auto");
  } else {
    // Send another type of email notification (sufficient balance)
    sendSufficientBalanceEmail(columnH, columnI, duration);
  }

  function getLeaveBalance(email,leaveType) {
  var leaveBalanceSheetName = "LeaveBalance"; // Name of the "LeaveBalance" sheet
  var leaveBalanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    leaveBalanceSheetName
  );
  var leaveBalanceData = leaveBalanceSheet.getDataRange().getValues();

  // Find the row that matches the email and leave type
  for (var i = 1; i < leaveBalanceData.length; i++) {
    // Assuming row 1 contains headers
    var sheetEmail = leaveBalanceData[i][1]; // Email is in the second column
    var sheetLeaveType = leaveBalanceData[i][2]; // Leave type is in the third column
    var sheetLeaveBalance = leaveBalanceData[i][5]; // Leave balance is in the fifth column

    if (sheetEmail === email && sheetLeaveType === leaveType) {
      return sheetLeaveBalance;
      Logger.log(sheetLeaveBalance);
    }
  }

  // If no matching record is found, return a default value or handle as needed
  return 0; // Default balance if no match is found
  }

  function sendInsufficientBalanceEmail(sendTo, leaveType, duration) {
  // Implement this function to send an email for insufficient leave balance
  // You can customize the email content as needed.
  var mailSubject = "Insufficient Leave Balance - " + leaveType;
  var mailBody =
    "Dear Employee,<br>" +
    "Your request for " +
    leaveType +
    " leave has been received, but the requested duration (" +
    duration +
    " days) exceeds your available leave balance.<br>" +
    "Please review your leave balance and consider adjusting your request accordingly.<br><br>" +
    "Thank you,<br>" +
    "HR Department";

  sendMail(sendTo, mailSubject, mailBody);
  }

  function sendSufficientBalanceEmail(sendTo, leaveType, duration) {
  // Implement this function to send an email for sufficient leave balance
  // You can customize the email content as needed.

  var mailSubject = "Leave Request Submitted - " + leaveType;
  var mailBody =
    "Dear Employee,<br>" +
    "Your request for " +
    leaveType +
    " leave has been submitted. The requested duration is " +
    duration +
    " days.<br>" +
    "After approved you will get email notification"+
    "Please ensure to manage your workload accordingly during your absence.<br><br>" +
    "Thank you,<br>" +
    "HR Department";

  sendMail(sendTo, mailSubject, mailBody);
  Logger.log(sendTo, mailSubject, mailBody);
  }

  function sendMail(sendTo, mailSubject, mailBody) {
  // Implement your email sending code here (e.g., using MailApp)
  // This function should send the email to the specified recipient(s).
  MailApp.sendEmail({
    to: sendTo,
    subject: mailSubject,
    htmlBody: mailBody,
  });
    
  }




}

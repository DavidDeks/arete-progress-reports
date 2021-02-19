var gymnast = [];
var k = 0;
const CHECKBOX_END_INDEX = 31;
const EMAIL_START_INDEX = 36;
var ui = SpreadsheetApp.getUi();


function doGet() { // This name is only needed if creating a webapp
  // var emails = SpreadsheetApp.openById('18MKXCWmaP60CFKG1IWlMF3_XKefdKVO-n6OYZPSxQkQ').getDataRange().getValues(); // For use with webapp since ui isn't available in that context
  // I would need to use an array of manually entered sheet Ids if I were to make a webapp version of this project.
  var sheet = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  var level = getCheckedLevels(sheet);
  if (!verifyCheckedLevels(level)) {
    return;
  }
  var report = getReports(level);
  if (!report) {
    return;
  }
  if (!verifyQuota(getGymnastCount(report))) {
    return;
  }
  getGymnasts(sheet, report, level);
  if (!gymnast) {
    return;
  }
  sendEmails();
}

function getCheckedLevels(sheet) {
  var level = [];
  for (var i = 1; i < CHECKBOX_END_INDEX; i++){ // Starts at 1 due to header row
    if (sheet[i][0]) {
      level.push(sheet[i][1]);
    }
  }
  return level;
}

function verifyCheckedLevels(level) {
  // Create the message
  var message = 'Are you sure you want to send progress reports for the following levels?\n\n';
  message += level.join(', ');
  Logger.log(message);
  // Display the dialog box
  var response = ui.alert(message, ui.ButtonSet.YES_NO);
  // Return false if user clicks "NO" or closes the dialog box
  if (response == ui.Button.NO) {
    Logger.log('The user clicked "NO" or the dialog\'s close button on the level selection alert.');
    return false;
  }
  // Return true if user clicks "YES"
  Logger.log('The user clicked "YES" to confirm level selection.');
  return true;
}

function getReports(level) { // Filename > File Iterator > File > Spreadsheet > 1st Sheet > report/scores
  // Get the reports by name for all of the selected levels.
  var reportIterator = [];
  for (var i = 0, j = level.length; i < j; i++) {
    reportIterator.push(DriveApp.getFilesByName(level[i] + ' Progress Report (Responses)'));
  }
  // Get the files from the file iterators
  var reportFile = [];
  for (var i = 0, j = reportIterator.length; i < j; i++) {
    while (reportIterator[i].hasNext()) { // There should only be 1 file per iterator, but this will add duplicate files found to be checked later.
      reportFile.push(reportIterator[i].next()); 
    }
  }
  // Log any errors and display error message for any reports not found. End script and don't send emails if any report is not found or duplicates found.
  if (reportFile.length != level.length) {
    message = 'Emails not sent. The was a problem retrieving the selected Progress Report (Responses) files. The following files were found:\n\n' 
      + reportFile.join(', ') + '\n\nThe following levels were selected:\n\n' + level.join(', ') 
      + '\n\nPlease verify the Google Sheets exist and match the listed levels exactly and that duplicate file names do not exist.';
    Logger.log(message);
    ui.alert(message, ui.ButtonSet.OK);
    return null;
  }
  // Get spreadsheet from each file
  var spreadsheet = [];
  for (var i = 0, j = reportFile.length; i < j; i++) {
    spreadsheet.push(SpreadsheetApp.openById(reportFile[i].getId()));
  }
  // Get report from first sheet in each spreadsheet
  var report = [];
  for (var i = 0, j = spreadsheet.length; i < j; i++) {
    report.push(spreadsheet[i].getSheets()[0].getDataRange().getValues());
  }
  return report;
}

function getGymnastCount(report) {
  var count = 0;
  for (var i = 0, j = report.length; i < j; i++) {
    count += (report[i].length - 1); // -1 for header row
  }
  return count;
}

function verifyQuota(count) {
  var quota = MailApp.getRemainingDailyQuota();
  if (quota < count) {
    message = 'Emails not sent. The number of progress reports exceeds the remaining daily recipient quota.\n\nSend Quota: ' + quota + '\ngymnast: ' + count;
    Logger.log(message);
    ui.alert(message, ui.ButtonSet.OK);
    return false;
  }
  // Display message for end user to confirm they want to send emails after reviewing daily recipient quota
  message = 'Are you sure you want to send the progress reports?\n\nDaily recipient quota: ' + quota 
    + '\nProgress reports to send: ' + count + '\nRemaining quota if emails are sent: ' + (quota - count);
  Logger.log(message);
  response = ui.alert(message, ui.ButtonSet.YES_NO);
  // Process user's response
  if (response == ui.Button.NO) {
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar in response to sending emails based on the remaining quota.');
    return false;
  }
  Logger.log('The user clicked "Yes" to send the emails based on the remaining quota.');
  return true;
}

function getAllGymnasts(sheet) {
  var allGymnasts = [];
  for (var i = EMAIL_START_INDEX, j = sheet.length; i < j; i++) {
    allGymnasts.push({contactName:sheet[i][1], fullName:sheet[i][2], email:sheet[i][3]});
  }
  return allGymnasts;
}
  
function getGymnasts(sheet, report, level) {
  allGymnasts = getAllGymnasts(sheet);
  currentGymnast = [];
  for (var i = 0, x = report.length; i < x; i++) {
    for (var j = 1, y = report[i].length; j < y; j++) { // Start at 1 to skip header row
      currentGymnast = allGymnasts.filter( function(gymnast){return (gymnast.fullName.toUpperCase().trim()==report[i][j][2].toUpperCase().trim());} );
      if (!currentGymnast[0]) { // Gymnast not found in email list
        message = 'Emails not sent. ' + report[i][j][2].trim() + ' from ' + level[i] + ' was not found in the email list. Please make sure the gymnast\'s name matches exactly.';
        Logger.log(message);
        ui.alert(message, ui.ButtonSet.OK);
        gymnast = null;
        return;
      }
      currentGymnast[0].level = level[i];
      currentGymnast[0].scores = report[i][j];
      currentGymnast[0].tableContent = generateTableContent(report[i], j);
      gymnast.push(currentGymnast[0]);
    } 
  }
  return;
}

function includeExternalFile(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function generateTableContent(report, row) {
  var content = '';
  var nbrOfCols = report[0].length;
  var th = '';
  for (var i = 3; i < nbrOfCols; i++) { // First 3 columns contain Timestamp, Coach, and Gymnast. Therefore, i = 3 is where scores start.
    var currentCell = report[0][i];
    // Add the comments section if present at the end. The wording of this column is sometimes different, but always includes "comment".
    if (currentCell.trim().toUpperCase().includes('COMMENT')) {
      content += '</table> <h3>Comments:</h3> <p>' + report[row][i] + '</p>';
    }
    else {
      var currentTh = currentCell.slice(0, currentCell.indexOf(' [')); // E.g. this extracts "Floor" from "Floor [Kick lunge]"
      // Check if the current event has changed and new table header needs to be added
      if (currentTh != th) { 
        th = currentTh;
        content += '<tr> <th colspan=2>' + th + '</th> </tr>';
      }
      // Add the next skill as td
      var td = currentCell.slice(currentCell.indexOf('[') + 1, -1); // E.g. this extracts "Kick lunge" from "Floor [Kick lunge]"
      content += '<tr> <td>' + td + '</td> <td>' + report[row][i] + '</td> </tr>';
      // If it's the last column, close the table
      if (i == nbrOfCols - 1) {
        content += '</table>';
      }
    }
  }
  return content;
}

function sendEmails() {     
  for(var i = 0, j = gymnast.length; i < j; i++) {
    k = i;
    var recipient = gymnast[i].email;
    var subjectLine = "Arete Progress Report";
    var messageBody = "Please view this email on a device capable of rendering html.";
    var htmlBody = HtmlService.createTemplateFromFile('template').evaluate().getContent();
    // let htmlBody = HtmlService.createHtmlOutputFromFile('template').getContent();
    var sender = "Arete Gymnastics";
    var replyTo = "aretegymnastics01@gmail.com";
    GmailApp.sendEmail(recipient, subjectLine, messageBody, {htmlBody: htmlBody, name: sender, replyTo: replyTo});
    Logger.log('Progress report for ' + gymnast[i].fullName + ' sent to ' + recipient);
    // ui.showModalDialog(htmlBody, 'Email Preview'); // Used to preview email body, but wasn't able to get these dialog messages to actually render the html
    // return htmlBody; // Used for rendering in webapp. Could only get to return one email body.
  } 
}
var gymnast = [];
var k = 0;

function doGet() { // This name is only needed if creating a webapp
  
  var emails = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  // var emails = SpreadsheetApp.openById('18MKXCWmaP60CFKG1IWlMF3_XKefdKVO-n6OYZPSxQkQ').getDataRange().getValues(); // For use with webapp since ui isn't available in that context
  // I would need to use an array of manually entered sheet Ids if I were to make a webapp version of this project.
  
  // Verify with a message that the user wants to send emails to the selected levels. OK/Cancel prompt
  var message = 'Are you sure you want to send progress reports for the following levels?\n\n';
  var level = [];
  
  // Get selected checkboxes/levels
  for (var i = 1; i < 32; i++){
    if (emails[i][0]) {
      level.push(emails[i][1]);
    }
  }
  message += level.join(', ');
  Logger.log(message);
  
  // Display the dialog box
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(message, ui.ButtonSet.YES_NO);

  // End if user clicks cancel or closes the dialog box
  if (response == ui.Button.NO) {
    Logger.log('The user clicked "NO" or the dialog\'s close button on the level selection alert.');
    return;
  }
  Logger.log('The user clicked "YES" to confirm level selection.');
  
  // Get the score sheets by name for all of the selected levels.
  var scoreSheetsIterators = [];
  for (var i = 0; i < level.length; i++) {
    scoreSheetsIterators.push(DriveApp.getFilesByName(level[i] + ' Progress Report (Responses)')); // This returns file iterators. I will have to pull the files from the iterator which should only contain one file each.
  }
  
  // Get the files from the file iterators
  var scoreSheetsFiles = [];
  for (var i = 0; i < scoreSheetsIterators.length; i++) {
    while (scoreSheetsIterators[i].hasNext()) { // This is kind of a redundant check since there should only be 1 file on each file iterator, but will add duplicate files found.
      scoreSheetsFiles.push(scoreSheetsIterators[i].next()); 
    }
  }
  // Log any errors and display error message for any score sheets not found. End script and don't send emails if any sheet is not found or duplicates found.
  if (scoreSheetsFiles.length != level.length) {
    message = 'Emails not sent. The was a problem retrieving the selected Progress Report (Responses) files. The following files were found:\n\n' + scoreSheetsFiles.join(', ') + '\n\nThe following levels were selected:\n\n' + level.join(', ') + '\n\nPlease verify the Google Sheets exist and match the listed levels exactly and that duplicate file names do not exist.';
    Logger.log(message);
    ui.alert(message, ui.ButtonSet.OK);
    return;
  }
  // Get spreadsheets and gymnast count
  var sendQuota = MailApp.getRemainingDailyQuota();
  var scoreSs = [];
  var gymnastCount = 0;
  
  for (i = 0; i < scoreSheetsFiles.length; i++) { // scoreSheetsFiles is an array of file objects. Need to convert file to spreadsheet to sheet to row count
    var currentSs = SpreadsheetApp.openById(scoreSheetsFiles[i].getId());
    scoreSs.push(currentSs);
    var currentSheet = currentSs.getSheets()[0]; 
    gymnastCount += (currentSheet.getLastRow() - 1);
  }
  
  // Make sure send quota is >= nbr of gymnast from all levels.
  // Display message and end script execution if gymnast count exceeds quota.
  if (sendQuota < gymnastCount) {
    message = 'Emails not sent. The number of progress reports exceeds the remaining daily recipient quota.\n\nSend Quota: ' + sendQuota + '\ngymnast: ' + gymnastCount;
    Logger.log(message);
    ui.alert(message, ui.ButtonSet.OK);
    return;
  }
  // Display message for end user to confirm they want to send emails after reviewing daily recipient quota
  message = 'Are you sure you want to send the progress reports?\n\nDaily recipient quota: ' + sendQuota 
    + '\nProgress reports to send: ' + gymnastCount + '\nRemaining quota if emails are sent: ' + (sendQuota - gymnastCount);
  Logger.log(message);
  response = ui.alert(message, ui.ButtonSet.YES_NO);
  // Process user's response
  if (response == ui.Button.NO) {
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar in response to sending emails based on the remaining quota.');
    return;
  }
  Logger.log('The user clicked "Yes" to send the emails based on the remaining quota.');
  
  // Get spreadsheet scores
  var scoreSheets = [];
  for (var i = 0; i < scoreSs.length; i++) {
    let currentSheet = scoreSs[i].getSheets()[0];
    let currentScores = currentSheet.getDataRange().getValues();
    scoreSheets.push(currentScores);
  }
  
  // Create array of gymnast objects so I can look up the email and parent's name by gymnast's name
  var gymnastEmails = [];
  for (var i = 1; i < emails.length; i++) {
    let gymnast = {fullName:emails[i][2], contactName:emails[i][1], email:emails[i][3]};
    gymnastEmails.push(gymnast);
  }

  // Need to check if any gymnast is not found in email list. If not found, end script execution and don't send emails.
  // While matching emails, assign level, scores, and tableContent properties.
  // gymnast = {fullName:'', contactName:'', email:'', level: '', scores:[], tableContent:''};
  for (var i = 0; i < scoreSheets.length; i++) { // Loop over each Level/Class in the array of scoreSheets
    for (var j = 1; j < scoreSheets[i].length; j++) { // Start at 1 to skip header row
      var currentGymnast = gymnastEmails.filter( function(gymnast){return (gymnast.fullName.toUpperCase()==scoreSheets[i][j][2].toUpperCase().trim());} );
      if (!currentGymnast[0]) { // Gymnast not found in email list
        message = 'Emails not sent. ' + scoreSheets[i][j][2].trim() + ' from ' + level[i] + ' was not found in the email list. Please make sure the gymnast\'s name matches exactly.';
        Logger.log(message);
        ui.alert(message, ui.ButtonSet.OK);
        return;
      }
      currentGymnast[0].level = level[i];
      currentGymnast[0].scores = scoreSheets[i][j];
      var tableContent = generateTableContent(scoreSheets[i], j);
      currentGymnast[0].tableContent = tableContent;
      gymnast.push(currentGymnast[0]);
    }
  }

  // Generate a progress report and send an email for each gymnast.
  // Should maybe add a check for duplicate gymnasts here
  // An enhancement would be to maybe show a preview of one or all the emails being sent to verify the body looks good? This might be counterproductive though.          
  for(var i = 0; i < gymnast.length; i++) {
    k = i;
    let recipient = gymnast[i].email;
    let subjectLine = "Arete Progress Report";
    let messageBody = "Please view this email on a device capable of rendering html.";
    let htmlBody = HtmlService.createTemplateFromFile('template').evaluate().getContent();
    // let htmlBody = HtmlService.createHtmlOutputFromFile('template').getContent();
    let sender = "Arete Gymnastics";
    let replyTo = "aretegymnastics01@gmail.com";

    Logger.log('Progress report for ' + gymnast[i].fullName + ' sent to ' + recipient);
                
    // ui.showModalDialog(htmlBody, 'Email Preview'); // Used to preview email body, but wasn't able to get these dialog messages to actually render the html
    GmailApp.sendEmail(recipient, subjectLine, messageBody, {htmlBody: htmlBody, name: sender, replyTo: replyTo});
    // return htmlBody; // Used for rendering in webapp. Could only get to return one email body.
  } 
}

// Function to include the external CSS in the template
function includeExternalFile(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Function to generate the HTML table
function generateTableContent(sheet, row) { // Takes the 2D sheet array as a parameter
  var content = '';
  var nbrOfCols = sheet[0].length;
  var th = '';
  for (var i = 3; i < nbrOfCols; i++) { // First 3 columns contain Timestamp, Coach, and Gymnast. Therefore, i = 3 where scores start.
    var currentCell = sheet[0][i];
    // Add the comments section if present at the end. The wording of this column is sometimes different, but always includes "comment".
    if (currentCell.trim().toUpperCase().includes('COMMENT')) {
      content += '</table> <h3>Comments:</h3> <p>' + sheet[row][i] + '</p>';
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
      content += '<tr> <td>' + td + '</td> <td>' + sheet[row][i] + '</td> </tr>';
      // If it's the last column, close the table
      if (i == nbrOfCols - 1) {
        content += '</table>';
      }
    }
  }
  return content;
}
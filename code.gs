var gymnasts = [];
var k = 0;

function doGet() {
  
  var emails = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  
  // Verify with a message that the user wants to send emails to the selected levels. OK/Cancel prompt
  var message = 'Are you sure you want to send progress reports for the following levels?\n\n';
  var levels = [];
  
  // Get selected checkboxes/levels
  for (var i = 1; i < 22; i++){
    if (emails[i][0]) {
      levels.push(emails[i][1]);
    }
  }
  message += levels.join(', ');
  Logger.log(message);
  
  // Display the dialog box
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(message, ui.ButtonSet.YES_NO);

  // End if user clicks cancel or closes the dialog box
  if (response == ui.Button.NO) {
    Logger.log('The user clicked "NO" or the dialog\'s close button on the levels selection alert.');
  }
  else {
    Logger.log('The user clicked "YES" to confirm levels selection.');
    // Generate the templates based on the form title and questions
    // https://developers.google.com/apps-script/reference/forms/form#getitems
    // getItems()
    
    
    // Get the score sheets by name for all of the selected levels.
    var scoreSheetsIterators = [];
    for (var i = 0; i < levels.length; i++) {
      scoreSheetsIterators.push(DriveApp.getFilesByName(levels[i] + ' Progress Report (Responses)')); // This returns file iterators. I will have to pull the files from the iterator which should only contain one file each.
    }
    
    // Get the files from the file iterators
    var scoreSheetsFiles = [];
    for (var i = 0; i < scoreSheetsIterators.length; i++) {
      while (scoreSheetsIterators[i].hasNext()) { // This is kind of a redundant check since there should only be 1 file on each file iterator, but will add duplicate files found.
        scoreSheetsFiles.push(scoreSheetsIterators[i].next()); 
      }
    }
    // Log any errors and display error message for any score sheets not found. End script and don't send emails if any sheet is not found or duplicates found.
    if (scoreSheetsFiles.length != levels.length) {
      message = 'Emails not sent. The was a problem retrieving the selected Progress Report (Responses) files. The following files were found:\n\n' + scoreSheetsFiles.join(', ') + '\n\nThe following levels were selected:\n\n' + levels.join(', ') + '\n\nPlease verify the Google Sheets exist and match the listed levels exactly and that duplicate file names do not exist.';
      Logger.log(message);
      ui.alert(message, ui.ButtonSet.OK);
    }
    else { // Get spreadsheets and gymnast count
      var sendQuota = MailApp.getRemainingDailyQuota();
      var scoreSs = [];
      var gymnastCount = 0;
      
      for (i = 0; i < scoreSheetsFiles.length; i++) { // scoreSheetsFiles is an array of file objects. Need to convert file to spreadsheet to sheet to row count
        var currentSs = SpreadsheetApp.openById(scoreSheetsFiles[i].getId());
        scoreSs.push(currentSs);
        var currentSheet = currentSs.getSheets()[0]; 
        gymnastCount += (currentSheet.getLastRow() - 1);
      }
    
    // Need to add a check for duplicate gymnast names in the same level's score sheet. 
    // I guess it would be okay if they appeared in 2 separate levels because that may be possible, 
    // but may want to display a message with the names of those that appear in two separate levels.
      
      // Make sure send quota is >= nbr of gymnasts from all levels.
      // Display message and end script execution if gymnast count exceeds quota.
      if (sendQuota < gymnastCount) {
        message = 'Emails not sent. The number of progress reports exceeds the remaining daily recipient quota.\n\nSend Quota: ' + sendQuota + '\nGymnasts: ' + gymnastCount;
        Logger.log(message);
        ui.alert(message, ui.ButtonSet.OK);
      }
      else { // Display message for end user to confirm they want to send emails after reviewing daily recipient quota
        message = 'Are you sure you want to send the progress reports?\n\nDaily recipient quota: ' + sendQuota + '\nProgress reports to send: ' + gymnastCount + '\nRemaining quota if emails are sent: ' + (sendQuota - gymnastCount);
        Logger.log(message);
        response = ui.alert(message, ui.ButtonSet.YES_NO);
        
        // Process user's response
        if (response == ui.Button.NO) {
          Logger.log('The user clicked "No" or the close button in the dialog\'s title bar in response to sending emails based on the remaining quota.');
        } else {
          Logger.log('The user clicked "Yes" to send the emails based on the remaining quota.');
          
          // Get spreadsheet scores
          var scoreSheets = [];
          for (var i = 0; i < scoreSs.length; i++) {
            let currentSheet = scoreSs[i].getSheets()[0];
            let currentScores = currentSheet.getDataRange().getValues();
            scoreSheets.push(currentScores);
          }
          
          // Create array of gymnast objects so I can look up the email and parent's name by gymnast's name
          var gymnastsEmails = [];
          for (var i = 1; i < emails.length; i++) {
            let gymnast = {fullName:emails[i][7], parentName:emails[i][5], email:emails[i][11]};
            gymnastsEmails.push(gymnast);
          }
  
          // Need to check if any gymnasts are not found in email list. If not found, end script execution and don't send emails.
          // While matching emails, assign level and scores properties.
          // gymnast = {fullName:'', parentName:'', email:'', level: '', scores:[]};
          var currentGymnast = []; // Using an array because the filter function returns and array. This is the only way I know how to do the lookup.
          levelLoop:
          for (var i = 0; i < scoreSheets.length; i++) {
            scoresLoop:
            for (var j = 1; j < scoreSheets[i].length; j++) { // Start at 1 to skip header row
              currentGymnast = gymnastsEmails.filter( function(gymnast){return (gymnast.fullName.toUpperCase()==scoreSheets[i][j][2].toUpperCase().trim());} ); // Returns an array with just the one gymnast object element.
              if (!currentGymnast[0]) { 
                break levelLoop; 
              }
              currentGymnast[0].level = levels[i];
              currentGymnast[0].scores = scoreSheets[i][j];
              gymnasts.push(currentGymnast[0]);
            }
          }
          if (!currentGymnast[0]) {
            message = 'Emails not sent. ' + scoreSheets[i][j][2].trim() + ' from ' + levels[i] + ' was not found in the email list. Please make sure the gymnast\'s name matches exactly.';
            Logger.log(message);
            ui.alert(message, ui.ButtonSet.OK);
          }    
          else { // Generate a progress report and send an email for each gymnast.
            for(var i = 0; i < gymnasts.length; i++) {
              k = i;
              let recipient = gymnasts[i].email;
              let subjectLine = "Arete Progress Report";
              let messageBody = "Please view this email on a device capable of rendering html.";
              let htmlBody = HtmlService.createTemplateFromFile(gymnasts[i].level).evaluate().getContent();
              let sender = "Arete Gymnastics";
              let replyTo = "aretegymnastics01@gmail.com";
        
              Logger.log('Progress report for ' + gymnasts[i].fullName + ' sent to ' + recipient);
              
              GmailApp.sendEmail(recipient, subjectLine, messageBody, {htmlBody: htmlBody, name: sender, replyTo: replyTo});
              // return htmlBody; // Used for rendering in webapp. To work, have to remove getContent(), comment out all ui/alert references, and I think rename template to index.
          
          
            } // close send emails quota confirmation message
          } // close send quota exceeded error message
        } // close score sheets lookup error message
      } // close levels selection confirmation
    } // close sendEmail else statement
  } // close email not found else statement
} // close doGet

// Function to include the external CSS in the templates, rather than having to copy/paste inline CSS for each one.
function includeExternalFile(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
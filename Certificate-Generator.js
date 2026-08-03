// Function to generate certificates and email them safely from your Samsung phone
function generateAndEmailCertificates() {
  
  // 1. Your Google Doc Certificate Template ID (the unique design file code)
  var templateId = "1SoQil8RCvzgnXaOO54o5KJYhInD8_Xw8-VKeVDOXizM";
  
  // 2. Direct web link to your participant Google Sheet
  var sheetUrl = "https://docs.google.com/spreadsheets/d/1Jo2E_yn0vckq6Ohib3bHtvU-FVvtOoU77Q2jgoKe4oY/edit?usp=drivesdk"; 
  
  // 3. Name of the folder in Google Drive where certificates are stored
  var folderName = "Generated Certificates";
  
  // Look for the 'Generated Certificates' folder in Drive; if missing, create it automatically
  var targetFolder;
  var folderSearch = DriveApp.getFoldersByName(folderName);
  if (folderSearch.hasNext()) {
    targetFolder = folderSearch.next();
  } else {
    targetFolder = DriveApp.createFolder(folderName);
  }
  
  // Open your Google Sheet directly using its web link
  var spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  
  // Select the specific tab named "Certificates with mail"
  var sheet = spreadsheet.getSheetByName("Certificates with mail");
  
  // If that specific tab name is not found, default to the active tab
  if (!sheet) {
    sheet = spreadsheet.getActiveSheet();
  }
  
  // Count total rows starting from Row 2 (skipping header titles in Row 1)
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No participant entries found in the sheet.");
    return;
  }
  
  // Fetch data across 5 columns: Column A (Name) through Column E (Email)
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  
  // Process each participant row one by one in a loop
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    var participantName = row[0] ? row[0].toString().trim() : ""; // Column A: Name
    var topicName       = row[1] ? row[1].toString().trim() : ""; // Column B: Topic
    var dateString      = row[2];                                 // Column C: Date
    var certId          = row[3] ? row[3].toString().trim() : ""; // Column D: Certificate ID
    var recipientEmail  = row[4] ? row[4].toString().trim() : ""; // Column E: Email Address
    
    // Skip empty rows if there is no name listed
    if (participantName === "") {
      continue;
    }
    
    // Format raw date objects into clean text (for example: 15 August 2026)
    if (dateString instanceof Date) {
      dateString = Utilities.formatDate(dateString, Session.getScriptTimeZone(), "dd MMMM yyyy");
    }
    
    // Create a copy of your certificate template inside your 'Generated Certificates' folder
    var templateFile = DriveApp.getFileById(templateId);
    var newFile = templateFile.makeCopy("Certificate - " + participantName, targetFolder);
    
    // Open the new certificate document to replace text placeholders
    var doc = DocumentApp.openById(newFile.getId());
    var body = doc.getBody();
    
    // Replace text placeholders with actual participant details
    body.replaceText("{{Name}}", participantName);
    body.replaceText("{{Topic}}", topicName);
    body.replaceText("{{Date}}", dateString);
    body.replaceText("{{Certificate ID}}", certId);
    
    // Apply the calligraphic font (Caveat) to the participant's name safely
    var nameRange = body.findText(participantName);
    if (nameRange) {
      var nameElement = nameRange.getElement().asText();
      nameElement.setFontFamily("Caveat"); 
      nameElement.setFontSize(24); 
    }
    
    // Save changes and close the completed Google Doc
    doc.saveAndClose();
    
    // SEND EMAIL: Only send if a valid email address exists in Column E
    if (recipientEmail !== "") {
      
      // Convert the Google Doc certificate into a PDF attachment
      var pdfAttachment = newFile.getAs("application/pdf");
      
      // Define email Subject line and Message text
      var subject = "Your Certificate of Completion - " + topicName;
      var bodyMessage = "Dear " + participantName + ",\n\n" +
                        "Congratulations on successfully completing the program on '" + topicName + "'.\n" +
                        "Please find your certificate attached to this email.\n\n" +
                        "Best regards,\n" +
                        "Program Coordinator";
      
      // Send the email using Google Mail Service
      MailApp.sendEmail({
        to: recipientEmail,
        subject: subject,
        body: bodyMessage,
        attachments: [pdfAttachment]
      });
      
      Logger.log("SUCCESS: Email sent to " + recipientEmail);
    } else {
      Logger.log("WARNING: No email address found in Column E for " + participantName);
    }
  }
  
  Logger.log("Process Complete!");
}

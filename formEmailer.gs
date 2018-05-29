var EMAIL_SENT = "EMAIL_SENT";

function sendEmails() {

var sheet = SpreadsheetApp.getActiveSheet();

  var startRow = 2;  // First row of data to process
  var numRows = 5;   // Number of rows to process
  // Fetch the range of cells A2:C6
  var dataRange = sheet.getRange(startRow, 1, numRows, 3)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0];  // First column
    var message = row[1];       // Second column
    
    var doc = DocumentApp.create('Form Email Sender');
    var body = doc.getBody();
    body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Name:');
    body.appendParagraph('Sunny');
    
    body.appendParagraph(emailAddress);
    body.appendParagraph(message + ' ' + '☎');
    
    ////Highlighting text///
  var textToHighlight = message;
  var highlightStyle = {};
  highlightStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FF0000';
  var paras = doc.getParagraphs();
  var textLocation = {};
  var j;

  for (j=0; j<paras.length; ++j) {
    textLocation = paras[j].findText(textToHighlight);
    if (textLocation != null && textLocation.getStartOffset() != -1) {
      textLocation.getElement().setAttributes(textLocation.getStartOffset(),textLocation.getEndOffsetInclusive(), highlightStyle);
    }
  }
  /////////////////////////////////
  
    var emailSent = row[2];     // Third column
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var subject = "Sending pdf email";

     doc.saveAndClose();
     var pdf = DriveApp.getFileById(doc.getId()).getBlob().getAs('application/pdf').setName('Form Sender');
     MailApp.sendEmail(emailAddress, subject, message, {attachments:[pdf]});
     sheet.getRange(startRow + i, 3).setValue(EMAIL_SENT);
    
     DriveApp.getFilesByName('Form Email Sender').next().setTrashed(true);
    
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}


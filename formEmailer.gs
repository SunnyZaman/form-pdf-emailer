var EMAIL_SENT = "EMAIL_SENT";

function onOpen() {
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('Send Email')
  .addItem('send emails','showPrompt')
  .addToUi();
}

function sendEmails2(email) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lrow = sheet.getLastRow();
  var lcol = sheet.getLastColumn();
  var lastcolumn=lcol;
  if( SpreadsheetApp.getActiveSheet().getRange(1,lcol).getValue() != 'Email Confirmation'){
    lastcolumn=lcol+1;
  }
  SpreadsheetApp.getActiveSheet().getRange(1,lastcolumn).setValue('Email Confirmation');
  var startRow = 2;  // First row of data to process
  var dataRange = sheet.getRange(startRow, 1, lrow -1, lastcolumn);
  var headingsRange = sheet.getRange(1, 1, 1, lastcolumn); //added
  var dataH = headingsRange.getValues();//added
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailSent = row[lastcolumn-1];    
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var file = DriveApp.getFileById("1L9oIyqnqLQhVbPlid8zkromD-MS1hjh_8t1prD9CqUU");
      var newId = file.getId();
      var doc = DocumentApp.openById(newId);
      doc.setName('Title of File');
      var body = doc.getBody();
      body.insertParagraph(0, doc.getName())
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);
      for (var k = 0; k < dataH.length; ++k) {
        var rowH = dataH[k];
        var n;
        for(n=0;n<lastcolumn-1;n++){
          var contentH =rowH[n];
          // Define a custom paragraph style.
          var style = {};
          style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
          DocumentApp.HorizontalAlignment.LEFT;
          style[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
          style[DocumentApp.Attribute.FONT_SIZE] = 18;
          style[DocumentApp.Attribute.BOLD] = true;
          style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#104599';
          var style2 = {};
          style2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
          DocumentApp.HorizontalAlignment.LEFT;
          style2[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
          style2[DocumentApp.Attribute.FONT_SIZE] = 12;
          style2[DocumentApp.Attribute.BOLD] = false;
          style2[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
          var headcont=  body.appendParagraph(contentH);
          headcont.setAttributes(style);
          var content =row[n];
          var cont = body.appendParagraph(content);
          cont.setAttributes(style2);
          body.appendParagraph("\n");
        }
      }
      var subject = "Sending pdf email";
      doc.saveAndClose();
      var pdf = DriveApp.getFileById(doc.getId()).getBlob().getAs('application/pdf').setName('Form Sender');
      MailApp.sendEmail(email, subject, 'This is a message filler', {attachments:[pdf]});
      sheet.getRange(startRow + i, lastcolumn).setValue(EMAIL_SENT);
      doc = DocumentApp.openById(newId);
      doc.setText('');
      doc.setName('CCS Blank Google Doc Template');
      doc.saveAndClose();
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}

function showPrompt() {
  var ui=SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('emailBox')
  .setHeight(100)
  .setWidth(300);
  ui.showModalDialog(html, 'Email'); 
}

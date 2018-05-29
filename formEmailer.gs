
function onOpen() {
  var ui=SpreadsheetApp.getUi();

  ui.createMenu('Send Email')
  .addItem('send emails','showPrompt')
  .addToUi();
}

function showPrompt() {
  var ui = SpreadsheetApp.getUi(); 
  var result = ui.prompt(
      'Email Sender',
      'Send email to:',
      ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    sendEmails2(text);
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
  }
}

function sendEmails() {
var EMAIL_SENT = "EMAIL_SENT";

var sheet = SpreadsheetApp.getActiveSheet();

  var startRow = 2;  // First row of data to process
  var lrow = sheet.getLastRow();
  var lcol = sheet.getLastColumn();

  var lastcolumn=lcol;
  if( SpreadsheetApp.getActiveSheet().getRange(1,lcol).getValue() != 'Email Confirmation'){
    lastcolumn=lcol+1;
  }
  
  SpreadsheetApp.getActiveSheet().getRange(1,lastcolumn).setValue('Email Confirmation');
  var startRow = 2;

  var dataRange = sheet.getRange(startRow, 1, lrow -1, lastcolumn)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];

    var emailSent = row[lastcolumn-1];     
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
    
    
        var doc = DocumentApp.create('Form Email Sender');
  var body = doc.getBody();
  body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Name:');
  body.appendParagraph('Sunny');
       var n;
    for(n=0;n<lastcolumn-1;n++){
var content =row[n];
     body.appendParagraph(content);
     }
      var subject = "Sending pdf email";


    ////Highlighting text///
    /*var textToHighlight = message;
    var highlightStyle = {};
    highlightStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FF0000';
    var paragraph = doc.getParagraphs();
    var textLocation = {};
    var j;

    for (j=0; j<paragraph.length; ++j) {
      textLocation = paragraph[j].findText(textToHighlight);
      if (textLocation != null && textLocation.getStartOffset() != -1) {
        textLocation.getElement().setAttributes(textLocation.getStartOffset(),textLocation.getEndOffsetInclusive(), highlightStyle);
      }
    }*/
    /////////////////////////////////

     doc.saveAndClose();
     var pdf = DriveApp.getFileById(doc.getId()).getBlob().getAs('application/pdf').setName('Form Sender');
     MailApp.sendEmail(emailAddress, subject, message, {attachments:[pdf]});
   sheet.getRange(startRow + i, lastcolumn).setValue(EMAIL_SENT);
    
     DriveApp.getFilesByName('Form Email Sender').next().setTrashed(true);
    
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}


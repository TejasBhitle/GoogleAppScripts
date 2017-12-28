/* Called when spreadsheet opens */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Bulk Email Script') // create a menu option
      .addItem('Open', 'openDialog') //adding a menu option
      .addToUi();
}

/* Creating a dialog */
function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index');
  SpreadsheetApp.getUi().showModalDialog(html, 'Bulk Email GoogleAppScript');
}

/* Utility function */
function sendBulkEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var table = sheet.getDataRange().getValues();
  try{
    for (var i = 1; i < table.length; i++) {
      recipent = table[i][0];
      subject = "Test";
      body = table[i][1];
      MailApp.sendEmail(recipent,subject,body);
    }
    SpreadsheetApp.getUi().alert("Done !!");
    google.script.host.close()//close dialog
  }
  catch(e){
    var ui = SpreadsheetApp.getUi();
    ui.alert("Unsuccessful","Error report" 
      +"\r\nMessage: " + e.message+ "\r\nFile: " + e.fileName+ "\r\nLine: " + e.lineNumber,ui.Button.OK)
    google.script.host.close()//close dialog
  }
  
}

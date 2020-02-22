

function sendEmailsKentSt(attachment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email automation');

    var emailAddress = sheet.getRange("F2").getValue();
    var subject = sheet.getRange("F7").getValue();
    var message = sheet.getRange("F8").getValue();
    var carbonCopyEmailAddresses = sheet.getRange("F3").getValue();
    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic dashboard mail sender',
        cc: carbonCopyEmailAddresses};
        
    MailApp.sendEmail(emailAddress, subject, message, emailOptions);
  }

//}

function exportTopdfKentSt() {
    var exportFolderID = '1aZl2g9n7dCYoZ9lq_sNAhrmlSrA7ffap';
    var exportFileName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD Kent St").getRange(5,1,1,1).getValue();
 //   var exportRange =  DASHBOARD_SPREADSHEET.SixtyCastlereaghDashboardSheet.pdfExportRange;
    var exportOptions = {
        sheetId: 385035257,
        exportFolderId: exportFolderID,
        exportFileName: exportFileName,
        range:{r1: 0,
            r2: 60,
            c1: ColumnNames.letterToColumnStart0('A'),
            c2: ColumnNames.letterToColumnStart0('CU')},
        repeatHeader: true,
        portrait: false,
        fileFormat: 'pdf'
    };  
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardSheet = ss.getSheetByName('DASHBOARD Kent St');
    dashboardSheet.showSheet();
   
    var pdfFile = export(exportOptions);
    sendEmailsKentSt(pdfFile);
    dashboardSheet.hideSheet();
}
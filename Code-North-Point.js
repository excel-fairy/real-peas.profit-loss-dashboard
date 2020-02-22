

function sendEmailsNorthPoint(attachment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email automation');

    var emailAddressNorthPoint = sheet.getRange("C2").getValue();
    var subjectNorthPoint = sheet.getRange("C7").getValue();
    var messageNorthPoint = sheet.getRange("C8").getValue();
    var carbonCopyEmailAddressesNorthPoint = sheet.getRange("C3").getValue();
    var emailOptionsNorthPoint = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic North Point dashboard mail sender',
        cc: carbonCopyEmailAddressesNorthPoint};
        
    MailApp.sendEmail(emailAddressNorthPoint, subjectNorthPoint, messageNorthPoint, emailOptionsNorthPoint);
  }

//}

function exportToPdfNorthPoint() {
    var exportFolderID = '1aZl2g9n7dCYoZ9lq_sNAhrmlSrA7ffap';
    var exportFileName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD North Point").getRange(5,1,1,1).getValue();
//    var exportRange =  DASHBOARD_SPREADSHEET.NorthPointDashboardSheet.pdfExportRange;
    var exportOptions = {
        sheetId: 1513709729,
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
    var dashboardSheet = ss.getSheetByName('DASHBOARD North Point');
    dashboardSheet.showSheet();
   
    var pdfFile = export(exportOptions);
    sendEmailsNorthPoint(pdfFile);
    dashboardSheet.hideSheet();
}
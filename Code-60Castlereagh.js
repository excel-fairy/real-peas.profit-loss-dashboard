

function sendEmailsSixtyCastlereagh(attachment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email automation');

    var emailAddress = sheet.getRange("D2").getValue();
    var subject = sheet.getRange("D7").getValue();
    var message = sheet.getRange("D8").getValue();
    var carbonCopyEmailAddresses = sheet.getRange("D3").getValue();
    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic dashboard mail sender',
        cc: carbonCopyEmailAddresses};
        
    MailApp.sendEmail(emailAddress, subject, message, emailOptions);
  }

//}

function exportToPdfSixtyCastlereagh() {
    var exportFolderID = '1aZl2g9n7dCYoZ9lq_sNAhrmlSrA7ffap';
    var exportFileName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD 60Castlereagh").getRange(5,1,1,1).getValue();
 //   var exportRange =  DASHBOARD_SPREADSHEET.SixtyCastlereaghDashboardSheet.pdfExportRange;
    var exportOptions = {
        sheetId: 578643471,
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
    var dashboardSheet = ss.getSheetByName('DASHBOARD 60Castlereagh');
    dashboardSheet.showSheet();
   
    var pdfFile = export(exportOptions);
    sendEmailsSixtyCastlereagh(pdfFile);
    dashboardSheet.hideSheet();
}
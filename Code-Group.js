
function sendEmailsGroup(attachment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email automation');

    var emailAddress = sheet.getRange("E2").getValue();
    var subject = sheet.getRange("E7").getValue();
    var message = sheet.getRange("E8").getValue();
    var carbonCopyEmailAddresses = sheet.getRange("E3").getValue();
    
    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic Group dashboard mail sender',
        cc: carbonCopyEmailAddresses};
        
    MailApp.sendEmail(emailAddress, subject, message, emailOptions);
  }

//}

function exportToPdfBranch() {
    var exportFolderID = '1aZl2g9n7dCYoZ9lq_sNAhrmlSrA7ffap';
    var exportFileName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD Group").getRange(5,1,1,1).getValue();
//    var exportRange =  DASHBOARD_SPREADSHEET.pittStreetStoreDashboardSheet.pdfExportRange;
    var exportOptions = {
        sheetId: 826502350,
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
    var dashboardSheet = ss.getSheetByName('DASHBOARD Group');
    dashboardSheet.showSheet();
   
    var pdfFile = export(exportOptions);
    sendEmailsGroup(pdfFile);
    dashboardSheet.hideSheet();
}
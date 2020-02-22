/*var DASHBOARD_SPREADSHEET = {
    pittStreetStoreDashboardSheet: { 
        name: 'Pitt Street Store',
        dashSheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD Pitt Street'),
        pdfExportRange: {
            r1: 0,
            r2: 60,
            c1: ColumnNames.letterToColumnStart0('A'),
            c2: ColumnNames.letterToColumnStart0('CU')},
        fileName: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD Pitt Street").getRange(3,3,1,1).getValue()
        },
        
            NorthPointDashboardSheet: { 
        name: 'North Point Store',
        dashSheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD North Point'),
        pdfExportRange: {
            r1: 0,
            r2: 60,
            c1: ColumnNames.letterToColumnStart0('A'),
            c2: ColumnNames.letterToColumnStart0('CU')},
        fileName: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD North Point").getRange(1,3,1,1).getValue()
        },
    

SixtyCastlereaghDashboardSheet: { 
        name: '60Castlereagh Store',
        dashSheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 60Castlereagh'),
        pdfExportRange: {
            r1: 0,
            r2: 60,
            c1: ColumnNames.letterToColumnStart0('A'),
            c2: ColumnNames.letterToColumnStart0('CU')},
        fileName: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD 60Castlereagh").getRange(1,3,1,1).getValue()
        }
        };
*/
function sendEmailsPittStreet(attachment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email automation');

    var emailAddressPittStreet = sheet.getRange("B2").getValue();
    var subjectPittStreet = sheet.getRange("B7").getValue();
    var messagePittStreet = sheet.getRange("B8").getValue();
    var carbonCopyEmailAddressesPittStreet = sheet.getRange("B3").getValue();
    
    var emailOptionsPittStreet = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic Pitt Street dashboard mail sender',
        cc: carbonCopyEmailAddressesPittStreet};
        
    MailApp.sendEmail(emailAddressPittStreet, subjectPittStreet, messagePittStreet, emailOptionsPittStreet);
  }

//}

function exportToPdfPittStreet() {
    var exportFolderID = '1aZl2g9n7dCYoZ9lq_sNAhrmlSrA7ffap';
    var exportFileName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD Pitt Street").getRange(5,1,1,1).getValue();
//    var exportRange =  DASHBOARD_SPREADSHEET.pittStreetStoreDashboardSheet.pdfExportRange;
    var exportOptions = {
        sheetId: 2043617409,
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
    var dashboardSheet = ss.getSheetByName('DASHBOARD Pitt Street');
    dashboardSheet.showSheet();
   
    var pdfFile = export(exportOptions);
    sendEmailsPittStreet(pdfFile);
    dashboardSheet.hideSheet();
}
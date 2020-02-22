var SPREADSHEET = {
    spreadSheet: SpreadsheetApp.getActiveSpreadsheet(),
    sheets: {
        dashboard:{
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD"),
            exportFileNameCell: 'A5',
            branchCell: 'J10',
            dateCell: 'Y10',
            syncStatusCell: 'AM9',
            exportRange: {
                r1: 0,
                r2: 60,
                c1: ColumnNames.letterToColumnStart0('A'),
                c2: ColumnNames.letterToColumnStart0('CU')
            }
        },
        dataValid: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data valid"),
            exportFolderIdCell: 'L2'
        },
        emailAutomation: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email automation"),
            branchRowStart0: 0,
            recipientRowStart0: 1,
            carboCopyRowStart0: 2,
            subjectRowStart0: 3,
            messageRowStart0: 4,
            firstDataCol: ColumnNames.letterToColumn('B'),
            lastDataCol: ColumnNames.letterToColumn('F'),
            firstDataRow: 1,
            lastDataRow: 5
        }
    }
};

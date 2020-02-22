function exportDashboardForAllBranches() {
    var protection = protectDashboardSheet();
    setDate();
    var branchesEmailData = getEmailData();
    branchesEmailData.forEach(function (branchEmailData) {
        setBranch(branchEmailData.branchName);
        waitSyncDone();
        exportDashboardForOneBranch(branchEmailData);
    });

    // Unprotect the sheet now the export is done
    protection.remove();
}

function exportDashboardForOneBranch(branchEmailData) {
    var exportOptions = {
        sheetId: SPREADSHEET.sheets.dashboard.sheet.getSheetId(),
        exportFolderId: getExportFolderId(),
        exportFileName: getExportFileName(),
        range: SPREADSHEET.sheets.dashboard.exportRange,
        repeatHeader: true,
        portrait: false,
        fileFormat: 'pdf',
        margin: {
            top: 0,
            left: 0,
            right: 0,
            bottom: 0
        }
    };

    var pdfFile = ExportSpreadsheet.apply(exportOptions);
    sendEmail(pdfFile, branchEmailData);
}

function sendEmail(attachment, branchEmailData) {
    var emailAddress = branchEmailData.recipient;
    var subject = branchEmailData.subject;
    var message = branchEmailData.message;
    var carbonCopyEmailAddresses = branchEmailData.carbonCopy;

    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic dashboard mail sender',
        cc: carbonCopyEmailAddresses};

    MailApp.sendEmail(emailAddress, subject, message, emailOptions);
}

function getEmailData() {
    var values = SPREADSHEET.sheets.emailAutomation.sheet.getRange(
        SPREADSHEET.sheets.emailAutomation.firstDataRow,
        SPREADSHEET.sheets.emailAutomation.firstDataCol,
        SPREADSHEET.sheets.emailAutomation.lastDataRow - SPREADSHEET.sheets.emailAutomation.firstDataRow + 1,
        SPREADSHEET.sheets.emailAutomation.lastDataCol - SPREADSHEET.sheets.emailAutomation.firstDataCol + 1
    ).getValues();

    var nbBranches = values[SPREADSHEET.sheets.emailAutomation.branchRowStart0].length;
    var retVal = [];
    for (var i = 0; i < nbBranches + 1; i++) {
        var branchName = values[SPREADSHEET.sheets.emailAutomation.branchRowStart0][i];
        if(branchName && branchName !== '') {
            retVal.push({
                branchName: branchName,
                recipient: values[SPREADSHEET.sheets.emailAutomation.recipientRowStart0][i],
                carbonCopy: values[SPREADSHEET.sheets.emailAutomation.carboCopyRowStart0][i],
                subject: values[SPREADSHEET.sheets.emailAutomation.subjectRowStart0][i],
                message: values[SPREADSHEET.sheets.emailAutomation.messageRowStart0][i],
            });
        }
    }
    return retVal;
}

function getExportFolderId() {
    return SPREADSHEET.sheets.dataValid.sheet.getRange(SPREADSHEET.sheets.dataValid.exportFolderIdCell).getValue();
}

function getExportFileName() {
    return SPREADSHEET.sheets.dashboard.sheet.getRange(SPREADSHEET.sheets.dashboard.exportFileNameCell).getValue();
}

function setBranch(branchName) {
    SPREADSHEET.sheets.dashboard.sheet.getRange(SPREADSHEET.sheets.dashboard.branchCell).setValue(branchName);
}

function setDate() {
    function zeroPad(n, width) {
        n = n + '';
        return n.length >= width ? n : new Array(width - n.length + 1).join('0') + n;
    }

    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    var formattedDate = zeroPad(yesterday.getDate(), 2)
        + "/" + zeroPad(yesterday.getMonth() + 1, 2)
        + "/" + zeroPad(yesterday.getFullYear(), 4);
    SPREADSHEET.sheets.dashboard.sheet.getRange(SPREADSHEET.sheets.dashboard.dateCell).setValue(formattedDate);
}

function waitSyncDone() {
    // Wait and make sure the dashboard is fully loaded
    Utilities.sleep(10000);

    // Loop until the sync status indicator updates to "Done"
    // Doesn't work since the getValue() does not get the actual value (seems cached ...)
    // var syncStatus;
    // do {
        // syncStatus = SPREADSHEET.sheets.dashboard.sheet.getRange(SPREADSHEET.sheets.dashboard.syncStatusCell).getValue();
        // Utilities.sleep(400);
    // } while (syncStatus !== "DONE");
}

/**
 * Protect the dashboard sheet (and removes everyone from the edit whitelist) while it is being exported
 * (prevent any user to update it)
 */
function protectDashboardSheet() {
    var protection = SPREADSHEET.sheets.dashboard.sheet.protect().setDescription('Export sheet protection');
    protection.removeEditors(protection.getEditors());
    return protection;
}
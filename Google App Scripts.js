const SHEET_ID = 'Sheed_ID'; // Replace with your actual Google Sheet ID
const SHEET_NAME = 'Sheet_Name'; // Adjusted to your actual sheet name

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Sender')
    .addItem('Send Emails', 'showDialog')
    .addToUi();
}

function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('form')
    .setWidth(300)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Emails');
}

function sendEmails(startRow, numOfEmails) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const dataRange = sheet.getRange(startRow, 1, numOfEmails, 3);
  const data = dataRange.getValues();

  data.forEach((row, index) => {
    const rowIndex = startRow + index;
    const emailAddress = row[0];
    const subject = row[1];
    const body = row[2];
    // const trackingImageUrl = `${ScriptApp.getService().getUrl()}?event=open&row=${rowIndex}`;
    // const trackingImageUrl = `?event=open&row=${rowIndex}`;
    const trackingImageUrl = `Deploy_URL?event=open&row=${rowIndex}`; // Replace with your deploye url

    try {
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: body + `<a href="${trackingImageUrl}" style="height:1px;width:1px;position:absolute;visibility:hidden;">Track Open</a>`
      });

      const now = new Date();
      sheet.getRange(rowIndex, 4).setValue(now);  // Last Sent Email time
    } catch (e) {
      console.error("Failed to send email due to an error:", e.toString());
    }
  });
}

function processForm(formObject) {
  const startRow = parseInt(formObject.startRow, 10);
  const numOfEmails = parseInt(formObject.numOfEmails, 10);
  sendEmails(startRow, numOfEmails);
}

function doGet(e) {
  if (e.parameter.event === 'open' && e.parameter.row) {
    const rowIndex = parseInt(e.parameter.row, 10);
    return recordEmailOpen(rowIndex);
  }
  // Return a default message or pixel for non-matching requests
  return HtmlService.createHtmlOutput("Request not recognized");
}

function recordEmailOpen(rowIndex) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const now = new Date();

  const emailOpenInfo = sheet.getRange(rowIndex, 5, 1, 3).getValues()[0]; // fetch the range of first open, last open, total opens
  const firstOpenTime = emailOpenInfo[0];
  const lastOpenTime = emailOpenInfo[1];
  const totalOpens = emailOpenInfo[2];

  sheet.getRange(rowIndex, 6).setValue(now);  // Update the last open time
  sheet.getRange(rowIndex, 7).setValue((totalOpens || 0) + 1);  // Increment the total opens

  if (!firstOpenTime) {
    sheet.getRange(rowIndex, 5).setValue(now);  // Set the first open time if not set
  }

  return createTrackingPixelResponse(); // ensure to send a tracking pixel response
}

function createTrackingPixelResponse() {
  const base64Pixel = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
  const blob = Utilities.base64Decode(base64Pixel);
  const response = HtmlService.createHtmlOutputFromBlob(blob).setMimeType(ContentService.MimeType.GIF);
  return response;
}

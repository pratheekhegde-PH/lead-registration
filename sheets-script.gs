// Century Advantage — Lead Registration → Google Sheets
// ─────────────────────────────────────────────────────
// SETUP (one time, takes ~3 minutes):
//
// 1. Open your Google Sheet
// 2. Extensions → Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Click Save (💾)
// 5. Click Deploy → New deployment
//    • Type: Web app
//    • Execute as: Me
//    • Who has access: Anyone
// 6. Click Deploy → copy the Web App URL
// 7. Paste that URL into index.html where it says YOUR_APPS_SCRIPT_URL
// 8. Push to GitHub — Vercel will redeploy automatically
// ─────────────────────────────────────────────────────

const SHEET_NAME = 'Leads'; // Change if your sheet tab has a different name

const HEADERS = [
  'Lead ID',
  'Timestamp',
  'CP Name',
  'CP Mobile',
  'CP Email',
  'Primary Intent',
  'Lead Name',
  'Lead Mobile',
  'Project',
  'Property Type',
  'Budget',
  'BHK',
  'Additional Requirements',
  'Status',
];

function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET_NAME);

    // Create the sheet + headers if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setBackground('#6440A2')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.leadId        || '',
      data.timestamp     || new Date().toLocaleString('en-IN'),
      data.cpName        || '',
      data.cpMobile      || '',
      data.cpEmail       || '',
      data.intent        || '',
      data.leadName      || '',
      data.leadMobile    || '',
      data.project       || '',
      data.propertyType  || '',
      data.budget        || '',
      data.bhk           || '',
      data.requirements  || '',
      'New',             // default status
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', leadId: data.leadId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET handler — lets you test the endpoint is live by visiting the URL in a browser
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Century Advantage lead endpoint is live.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

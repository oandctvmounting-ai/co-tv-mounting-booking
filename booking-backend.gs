/**
 * C&O TV Mounting — Booking Intake Backend
 * =========================================
 * Deploy as a Google Apps Script Web App:
 *   1. Open https://script.google.com
 *   2. New Project → paste this entire file
 *   3. Create the target Sheet ONCE (or set BOOKINGS_SHEET_ID below to an existing sheet)
 *   4. Deploy → New deployment → type "Web app"
 *        - Execute as: Me
 *        - Who has access: Anyone (so the public booking page can POST)
 *   5. Copy the Web App URL and paste it into booking-page.html as APP_SCRIPT_URL
 *
 * The Web App expects a POST with JSON body:
 *   { name, phone, address, datetime, services, total, deposit, source }
 * and appends one row to the "Bookings" tab.
 */

const BOOKINGS_SHEET_ID = ''; // leave empty to auto-create a sheet named "C&O TV Mounting — Bookings"

function getSheet() {
  let ss;
  if (BOOKINGS_SHEET_ID) {
    ss = SpreadsheetApp.openById(BOOKINGS_SHEET_ID);
  } else {
    const existing = SpreadsheetApp.getActiveSpreadsheet();
    ss = existing;
  }
  let sheet = ss.getSheetByName('Bookings');
  if (!sheet) {
    sheet = ss.insertSheet('Bookings');
    sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Address', 'Date/Time',
                     'Primary Service', 'Add-on Services', 'TV Details',
                     'Estimated Total', 'Deposit', 'Promo', 'Source', 'Status']);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();
    const tvDetails = (data.tvDetails && data.tvDetails.length)
      ? data.tvDetails.map((t, i) =>
          `TV${i + 1}: ${t.mount || '?'} ${t.sizeLabel || ''} ${t.surfaceLabel || ''}` +
          `${t.overFireplace ? ' over-fireplace' : ''} ${t.wireLabel ? '+' + t.wireLabel : ''}`
        ).join(' | ')
      : '';

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.phone || '',
      data.address || '',
      data.datetime || '',
      data.primary || '',
      (data.addons && data.addons.length) ? data.addons.join(', ') : '',
      tvDetails,
      data.total != null ? data.total : '',
      data.deposit || '',
      data.promo || '',
      data.source || 'booking-page',
      'New'
    ]);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ status: 'C&O booking endpoint live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
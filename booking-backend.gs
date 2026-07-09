/**
 * C&O TV Mounting — Booking Intake Backend
 * =========================================
 * DEPLOYED AS A GOOGLE APPS SCRIPT WEB APP.
 *
 * Setup (one-time, ~3 min — done in your Google account):
 *   1. Open https://script.google.com  →  New Project
 *   2. Paste this ENTIRE file in (replacing any placeholder)
 *   3. Deploy  →  New deployment  →  type "Web app"
 *        - Execute as:  Me
 *        - Who has access:  Anyone   (so the public booking page can POST)
 *   4. Copy the Web App URL it gives you and send it to Jarvis.
 *      Jarvis drops it into booking-page.html as APP_SCRIPT_URL and redeploys.
 *
 * The page POSTs JSON:
 *   { name, phone, address, datetime, primary, addons[], tvDetails[],
 *     total, deposit, promo, source }
 * and this appends one row to the "Bookings" tab of the sheet below.
 *
 * SHEET: C&O TV Mounting  (ID below)
 */

const BOOKINGS_SHEET_ID = '1sOEzOQF0vFpx4l1tAxGSDS6vnLZ8j2AVUeifdcKxE5w';

function getSheet() {
  const ss = SpreadsheetApp.openById(BOOKINGS_SHEET_ID);
  let sheet = ss.getSheetByName('Bookings');
  if (!sheet) {
    sheet = ss.insertSheet('Bookings');
    sheet.appendRow([
      'Timestamp', 'Name', 'Phone', 'Address', 'Date/Time',
      'Primary Service', 'Add-on Services', 'TV Details',
      'Estimated Total', 'Deposit', 'Promo', 'Source', 'Status'
    ]);
    // freeze header + format total column as currency-ish
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    const tvDetails = (data.tvDetails && data.tvDetails.length)
      ? data.tvDetails.map((t, i) =>
          'TV' + (i + 1) + ': ' + (t.mount || '?') + ' ' + (t.sizeLabel || '') + ' ' +
          (t.surfaceLabel || '') +
          (t.overFireplace ? ' over-fireplace' : '') +
          (t.wireLabel ? ' +' + t.wireLabel : '')
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

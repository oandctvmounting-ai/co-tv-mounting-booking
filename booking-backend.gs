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
 * and this appends one row to the "Bookings" tab of the sheet below,
 * then sends a Telegram alert to the owner DM + the team group.
 *
 * SHEET: C&O TV Mounting  (ID below)
 */

const BOOKINGS_SHEET_ID = '1sOEzOQF0vFpx4l1tAxGSDS6vnLZ8j2AVUeifdcKxE5w';

// ===== Telegram alert config =====
// Get a bot token from @BotFather on Telegram, then paste it below.
// Chat IDs: owner DM = 6217602404, team group = -5510113560
const TELEGRAM_BOT_TOKEN = '8944456012:AAHYhDboqhLCzYysFGRZQ-xDT61X7aoQEKU';
const TELEGRAM_OWNER_CHAT = '6217602404';
const TELEGRAM_GROUP_CHAT = '-5510113560';

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

    // Fire Telegram alerts (best-effort — never block the booking on failure)
    try {
      sendTelegramAlert(data, tvDetails);
    } catch (te) {
      Logger.log('Telegram alert failed: ' + te);
    }

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

function sendTelegramAlert(data, tvDetails) {
  if (!TELEGRAM_BOT_TOKEN) return; // not configured yet
  const total = (data.total != null ? '$' + data.total : '');
  const deposit = data.deposit ? ('\n💳 Deposit: $' + data.deposit) : '';
  const promo = data.promo ? ('\n🏷 Promo: ' + data.promo) : '';
  const msg =
    '🔔 NEW BOOKING — C&O TV Mounting\n' +
    '👤 ' + (data.name || '?') + '  📞 ' + (data.phone || '?') + '\n' +
    '📍 ' + (data.address || '?') + '\n' +
    '🗓 ' + (data.datetime || '?') + '\n' +
    '🔧 ' + (data.primary || '?') +
    ((data.addons && data.addons.length) ? ' + ' + data.addons.join(', ') : '') + '\n' +
    (tvDetails ? '📺 ' + tvDetails + '\n' : '') +
    '💰 Est. Total: ' + total + deposit + promo + '\n' +
    '— lands in C&O Sheet (Bookings tab)';

  const url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
  const ownerPayload = { chat_id: TELEGRAM_OWNER_CHAT, text: msg };
  const groupPayload = { chat_id: TELEGRAM_GROUP_CHAT, text: msg };
  UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json',
    payload: JSON.stringify(ownerPayload) });
  UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json',
    payload: JSON.stringify(groupPayload) });
}

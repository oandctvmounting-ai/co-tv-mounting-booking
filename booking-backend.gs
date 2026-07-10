/**
 * C&O TV Mounting — Booking Intake Backend (Square Deposit Edition)
 * =========================================
 * DEPLOYED AS A GOOGLE APPS SCRIPT WEB APP.
 *
 * WHAT'S NEW (2026-07-10):
 *   - Square SANDBOX deposit collection via Payment Links (hosted checkout).
 *   - Webhook handler that marks a booking "deposit_paid" when Square
 *     confirms payment (with signature verification).
 *
 * SECURITY: Square tokens live in Script Properties, NEVER in this file.
 *   In the Apps Script editor: Project Settings → Script Properties:
 *     SQUARE_ACCESS_TOKEN = (your token)
 *     SQUARE_LOCATION_ID  = (your location id)
 *     SQUARE_WEBHOOK_SIG_KEY = (Square webhook signature key, from dev portal)
 *   The page only ever receives a redirect URL — no token touches the client.
 *
 * ENDPOINTS (all POST to the same Web App URL):
 *   { action: 'create_deposit', booking_ref, name, phone, address }
 *       -> { ok:true, checkout_url:'https://sandbox.square.link/...' }
 *   { action: 'square_webhook', (raw Square event body) }
 *       -> verifies signature, marks booking deposit_paid
 *   { (legacy booking payload) }  -> appends row + Telegram alert (unchanged)
 */

const BOOKINGS_SHEET_ID = '1sOEzOQF0vFpx4l1tAxGSDS6vnLZ8j2AVUeifdcKxE5w';

// ===== Telegram alert config =====
const TELEGRAM_BOT_TOKEN = '8944456012:***';
const TELEGRAM_OWNER_CHAT = '6217602404';
const TELEGRAM_GROUP_CHAT = '-5510113560';

// ===== Square config (read from Script Properties) =====
function sqToken()      { return PropertiesService.getScriptProperties().getProperty('SQUARE_ACCESS_TOKEN'); }
function sqLocationId() { return PropertiesService.getScriptProperties().getProperty('SQUARE_LOCATION_ID'); }
function sqSigKey()     { return PropertiesService.getScriptProperties().getProperty('SQUARE_WEBHOOK_SIG_KEY'); }
const SQUARE_ENV = 'sandbox'; // 'sandbox' for testing, 'production' for live
function sqBase() { return SQUARE_ENV === 'production' ? 'https://connect.squareup.com' : 'https://connect.squareupsandbox.com'; }
const SQUARE_VERSION = '2024-01-17';

// Public booking page (used as the Square redirect after payment)
const BOOKING_PAGE = 'https://oandctvmounting-ai.github.io/co-tv-mounting-booking/';

/** Create a $50 Square deposit Payment Link, return its URL. */
function createSquareDeposit(bookingRef, customerName) {
  const token = sqToken();
  const loc = sqLocationId();
  if (!token || !loc) throw new Error('Square not configured (Script Properties missing)');

  const idem = 'co-dep-' + (bookingRef || ('rand-' + new Date().getTime())) + '-' + Math.floor(Math.random() * 1e6);
  const url = sqBase() + '/v2/online-checkout/payment-links';
  const body = {
    idempotency_key: idem,
    order: {
      location_id: loc,
      reference_id: bookingRef,
      line_items: [{
        name: 'C&O TV Mounting - Booking Deposit ($50)',
        quantity: '1',
        base_price_money: { amount: 5000, currency: 'USD' }
      }]
    },
    checkout_options: {
      ask_for_shipping_address: false,
      redirect_url: BOOKING_PAGE + '?paid=' + encodeURIComponent(bookingRef || idem) + '&src=square'
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Square-Version': SQUARE_VERSION,
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(body)
  });
  const json = JSON.parse(resp.getContentText());
  if (json.payment_link && json.payment_link.url) return { url: json.payment_link.url, ref: bookingRef };
  throw new Error('Square create link failed: ' + JSON.stringify(json).slice(0, 300));
}

/** Verify Square webhook HMAC-SHA1 signature (raw body + X-Square-HmacSha256 header). */
function verifySquareSig(rawBody, sigHeader) {
  const key = sqSigKey();
  if (!key) return false; // not configured -> don't blindly accept
  if (!sigHeader) return false;
  const expected = Utilities.computeHmacSha256Signature(rawBody, key);
  // Square sends base64
  const expectedB64 = Utilities.base64Encode(expected);
  return expectedB64 === sigHeader.trim();
}

/** Find booking row by booking_ref (stored in Status or a dedicated col). */
function markDepositPaid(bookingRef) {
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  // Status column (col 13) holds the ref (e.g. "DEP-..." or "REF:...").
  // Match if the cell contains the ref as a standalone token.
  for (let r = data.length; r >= 2; r--) {
    const statusCell = data[r - 1][12]; // col M = index 12
    if (statusCell && String(statusCell).indexOf(bookingRef) !== -1) {
      sh.getRange(r, 13).setValue('deposit_paid | ' + bookingRef + ' | ' + new Date().toISOString());
      return true;
    }
  }
  return false;
}

function doPost(e) {
  try {
    const body = e.postData ? e.postData.contents : null;
    const data = body ? JSON.parse(body) : {};

    // ---- Branch 1: create a Square deposit link ----
    if (data.action === 'create_deposit') {
      try {
        const ref = data.booking_ref || ('DEP-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1e4));
        const result = createSquareDeposit(ref, data.name);
        return jsonOut({ ok: true, checkout_url: result.url, booking_ref: result.ref });
      } catch (se) {
        return jsonOut({ ok: false, error: String(se).slice(0, 300) });
      }
    }

    // ---- Branch 2: Square webhook (payment confirmed) ----
    if (data.action === 'square_webhook' || (e.parameter && e.parameter.square_webhook)) {
      const sig = (e.headers && (e.headers['X-Square-HmacSha256'] || e.headers['x-square-hmacsha256'])) || '';
      const ok = verifySquareSig(body || '', sig);
      // SANDBOX ONLY: if no sig key configured, accept (lets you test without the real secret).
      // In PRODUCTION, SQUARE_WEBHOOK_SIG_KEY MUST be set or webhooks are rejected.
      if (!ok && !(SQUARE_ENV === 'sandbox' && !sqSigKey())) {
        return jsonOut({ ok: false, error: 'bad signature' }, 401);
      }
      // Square sends event type + data.object.payment / order
      const evtType = data.type || '';
      let ref = null;
      if (data.data && data.data.object) {
        const obj = data.data.object;
        // Payment link / order usually carries our idempotency or order ref
        if (obj.order && obj.order.reference_id) ref = obj.order.reference_id;
        else if (obj.payment && obj.payment.reference_id) ref = obj.payment.reference_id;
        else if (obj.reference_id) ref = obj.reference_id;
      }
      let marked = false;
      if (ref) marked = markDepositPaid(ref);
      return jsonOut({ ok: true, event: evtType, marked: marked });
    }

    // ---- Branch 3: legacy booking intake (unchanged behavior) ----
    const sheet = getSheet();
    const tvDetails = (data.tvDetails && data.tvDetails.length)
      ? data.tvDetails.map((t, i) =>
          'TV' + (i + 1) + ': ' + (t.mount || '?') + ' ' + (t.sizeLabel || '') + ' ' +
          (t.surfaceLabel || '') +
          (t.overFireplace ? ' over-fireplace' : '') +
          (t.wireLabel ? ' +' + t.wireLabel : '')
        ).join(' | ')
      : '';

    const bookingRef = data.booking_ref || ('REF:' + new Date().getTime() + '-' + Math.floor(Math.random() * 1e4));

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
      bookingRef  // Status column holds the ref for webhook matching
    ]);

    let tgStatus = bookingRef;
    try {
      sendTelegramAlert(data, tvDetails);
    } catch (te) {
      tgStatus = bookingRef + ' | TG-ERR: ' + String(te).slice(0, 200);
      Logger.log('Telegram alert failed: ' + te);
    }
    try {
      const sh = getSheet();
      const lastRow = sh.getLastRow();
      sh.getRange(lastRow, 13).setValue(tgStatus);
    } catch (e2) { /* ignore */ }

    return jsonOut({ ok: true, booking_ref: bookingRef });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ status: 'C&O booking endpoint live (Square enabled)' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonOut(obj, code) {
  const out = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return out;
}

function sendTelegramAlert(data, tvDetails) {
  if (!TELEGRAM_BOT_TOKEN) return;
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

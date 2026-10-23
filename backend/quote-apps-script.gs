/**
 * Quote Desk — the order desk behind ai-dema-solutions.com/quote
 *
 * The quote page is a static file on GitHub Pages, so it has no server. This web
 * app is its back office: it takes the order, checks the money against its own
 * copy of the rate card, writes a row in the sheet, and sends two emails — the
 * confirmation with its PDF to the customer, and the whole record to the desk.
 *
 * ── Deploying it ────────────────────────────────────────────────────────────
 *  1. Create a Google Sheet, "Quote Desk — Orders".
 *  2. Extensions › Apps Script, paste this file, save.
 *  3. Deploy › New deployment › Web app
 *       Execute as:     Me
 *       Who has access: Anyone
 *     Authorise it (it sends mail as the account that deploys it).
 *  4. Copy the /exec URL into ORDERS_WEBHOOK in quote/index.html, and push.
 *
 * ── Changing it afterwards ──────────────────────────────────────────────────
 *  Deploy › Manage deployments › pencil › New version. **Never** "New
 *  deployment": that mints a different /exec URL and the page stops reaching it,
 *  exactly as it would for the waitlist form on the home page.
 */

var SHEET = 'orders';
var DESK = 'billing@dema-solutions.com';   // where the desk's copy goes
var SENDER_NAME = 'DEMA Solutions';

/* Who is issuing the document — from Dema Solutions' own invoices. */
var COMPANY = {
  name: 'Dema Solutions Srl',
  address: 'Via Antonio Gramsci, 10/12 - 20900 - Monza (MB)',
  vat: '10252060966',
  taxCode: '10252060966',
  email: 'billing@dema-solutions.com',
  phone: '0399467678',
  site: 'www.dema-solutions.com',
  iban: 'IT16K0306912999100000003492',
  swift: 'BCITITMM',
  paymentMethod: 'Bonifico bancario',
  terms: 'Payment on receipt. Work starts when payment is received.'
};

/* The rate card, Listino tecnico 2026. Kept here so the price can be checked
   rather than believed: a browser can send any number it likes. */
var AI_RATES = { basic: 0.02, 'good-enough': 0.04 };
var FLOORS = { basic: 5, 'good-enough': 8, professional: 12 };
var BANDS = {
  0.06: ['English (United Kingdom)', 'English (United States)', 'French (France)', 'Italian (Italy)',
    'Spanish (Spain)', 'Spanish (Latin America)', 'German (Germany)'],
  0.07: ['Portuguese (Brazil)', 'Portuguese (Portugal)', 'Dutch (Netherlands)', 'Flemish (Belgium)',
    'Romanian (Romania)', 'Russian (Russia)', 'Polish (Poland)', 'Bulgarian (Bulgaria)'],
  0.09: ['Czech (Czechia)', 'Slovak (Slovakia)', 'Croatian (Croatia)', 'Serbian (Serbia)',
    'Hungarian (Hungary)', 'Slovenian (Slovenia)'],
  0.10: ['Ukrainian (Ukraine)', 'Turkish (Turkey)', 'Bosnian (Bosnia and Herzegovina)',
    'Albanian (Albania)', 'Greek (Greece)', 'Danish (Denmark)'],
  0.11: ['Swedish (Sweden)', 'Norwegian (Norway)', 'Finnish (Finland)'],
  0.13: ['Estonian (Estonia)', 'Latvian (Latvia)', 'Lithuanian (Lithuania)', 'Georgian (Georgia)',
    'Maltese (Malta)', 'Persian (Iran)', 'Hebrew (Israel)', 'Arabic (Saudi Arabia)'],
  0.14: ['Basque (Spain)', 'Icelandic (Iceland)', 'Hindi (India)', 'Vietnamese (Vietnam)',
    'Thai (Thailand)', 'Swahili (Kenya)', 'Korean (South Korea)', 'Japanese (Japan)',
    'Chinese (Simplified)', 'Chinese (Traditional)']
};
var LAYOUT = { minimumCharge: 15, pagesIncluded: 15, perPage: 1 };
var VOLUME_BREAKS = [[50000, 0.15], [20000, 0.1], [5000, 0.05], [0, 0]];

function rateFor(tier, languageName) {
  if (AI_RATES[tier]) return AI_RATES[tier];
  for (var rate in BANDS) {
    if (BANDS[rate].indexOf(languageName) !== -1) return Number(rate);
  }
  return 0.14; // not on the card: the top band, never the cheapest
}

function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

function layoutTotal(order) {
  if (!order.layout || !order.layout.pages) return 0;
  var per = order.layout.pages <= LAYOUT.pagesIncluded
    ? LAYOUT.minimumCharge
    : order.layout.pages * LAYOUT.perPage;
  return round2(per * (order.targets || []).length);
}

/** What this job costs at list price, so a submitted total can be checked. */
function listPrice(order) {
  var words = Number(order.words) || 0;
  var discount = 0;
  for (var i = 0; i < VOLUME_BREAKS.length; i++) {
    if (words >= VOLUME_BREAKS[i][0]) { discount = VOLUME_BREAKS[i][1]; break; }
  }
  var floor = FLOORS[order.tier] || 0;
  var total = 0;
  for (var t = 0; t < (order.targets || []).length; t++) {
    total += Math.max(floor, words * rateFor(order.tier, order.targets[t]) * (1 - discount));
  }
  return round2(total + layoutTotal(order));
}

/* ──────────────────────────────────────────────────────────────── the web app */

function doGet() {
  return json({ ok: true, service: 'quote-desk-orders' });
}

function doPost(e) {
  try {
    var order = JSON.parse(e.postData.contents);
    if (!order.id || !order.customer || !order.customer.email) {
      return json({ ok: false, error: 'incomplete order' });
    }

    // Never take the price on trust: below list, the list price wins.
    var list = listPrice(order);
    var submitted = Number(order.net) || 0;
    var net = submitted >= list ? submitted : list;
    var repriced = net !== submitted;
    var vat = Number(order.vat) || 0;
    if (repriced) vat = round2(vat ? net * 0.22 : 0);
    var gross = round2(net + vat);

    record(order, net, vat, gross, repriced);

    var pdf = confirmationPdf(order, net, vat, gross);
    MailApp.sendEmail({
      to: order.customer.email,
      name: SENDER_NAME,
      subject: 'Your order ' + order.id + ' - ' + order.solution,
      htmlBody: customerBody(order, net, vat, gross),
      attachments: [pdf]
    });
    MailApp.sendEmail({
      to: DESK,
      name: SENDER_NAME,
      replyTo: order.customer.email,
      subject: 'New order ' + order.id + ' - ' + order.solution + ' - ' +
        gross.toFixed(2) + ' ' + order.currency,
      htmlBody: deskBody(order, net, vat, gross, repriced),
      attachments: [pdf]
    });

    return json({ ok: true, id: order.id });
  } catch (err) {
    // The order must not vanish because the mail or the sheet misbehaved.
    try {
      MailApp.sendEmail(DESK, 'Quote Desk: an order could not be processed',
        String(err) + '\n\n' + (e && e.postData ? e.postData.contents : ''));
    } catch (ignored) {}
    return json({ ok: false, error: String(err) });
  }
}

function json(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────────────────────────────────────────────── the sheet */

var HEADERS = ['Received', 'Reference', 'Customer', 'Email', 'Company', 'Country', 'VAT number',
  'Project', 'Solution', 'From', 'Into', 'Words', 'Subject', 'Desktop publishing', 'Delivery',
  'Net', 'VAT', 'Total', 'Currency', 'VAT treatment', 'Re-priced', 'Reference (customer)', 'Notes'];

function record(order, net, vat, gross, repriced) {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(SHEET) || book.insertSheet(SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  var b = order.billing || {};
  sheet.appendRow([
    new Date(), order.id, order.customer.name, order.customer.email,
    b.company ? b.companyName : 'individual', b.countryName || b.country || '', b.vatNumber || '',
    order.project || '', order.solution, order.source, (order.targets || []).join(', '),
    (Number(order.words) || 0) * (order.targets || []).length, order.subject,
    order.layout ? order.layout.pages + ' pages' : 'no',
    order.eta ? new Date(order.eta) : '', net, vat, gross, order.currency,
    order.vatTreatment || '', repriced ? 'yes' : '', order.customer.reference || '',
    order.customer.note || ''
  ]);
}

/* ──────────────────────────────────────────────────────────── the paperwork */

function money(amount, currency) {
  var symbol = currency === 'USD' ? '$' : '€';
  return symbol + Number(amount).toFixed(2);
}

function when(iso) {
  if (!iso) return '-';
  return Utilities.formatDate(new Date(iso), 'Europe/Rome', 'EEE d MMM, HH:mm') + ' CET';
}

function escape(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function recipientBlock(order) {
  var b = order.billing || {};
  var lines = [];
  lines.push(b.company && b.companyName ? b.companyName : (b.name || order.customer.name));
  if (b.company && b.name) lines.push(b.name);
  if (b.address1) lines.push(b.address1);
  if (b.address2) lines.push(b.address2);
  var city = [b.postalCode, b.city].filter(String).join(' ');
  if (city) lines.push(city + (b.country ? ' (' + b.country + ')' : ''));
  if (b.countryName) lines.push(b.countryName);
  if (b.telephone) lines.push(b.telephone);
  return lines;
}

/**
 * The confirmation, on the shape of a Dema Solutions invoice. It is an order
 * confirmation and says so: the fiscal invoice goes out through the SdI.
 */
function confirmationPdf(order, net, vat, gross) {
  var navy = '#1b2a6b';
  var count = (order.targets || []).length;
  var block = recipientBlock(order).map(function (line, i) {
    return '<div style="' + (i === 0 ? 'font-weight:700;color:#111' : 'color:#6e768c') + '">' +
      escape(line) + '</div>';
  }).join('');

  var rows = '<tr style="background:#f0f2f7">' +
    '<td style="padding:10px 8px"><strong>Translation services ' + escape(order.id) + '</strong>' +
    '<div style="color:#5c7a6a">' + escape(order.source) + ' &gt; ' + escape((order.targets || []).join(', ')) +
    ' - ' + ((Number(order.words) || 0) * count) + ' words - ' + escape(order.solution) + '</div></td>' +
    '<td style="padding:10px 8px;text-align:right;color:#6e768c">' + count +
    (count === 1 ? ' language' : ' languages') + '</td>' +
    '<td style="padding:10px 8px;text-align:right">' + money(net - layoutTotal(order), order.currency) +
    '</td></tr>';
  if (layoutTotal(order)) {
    rows += '<tr><td style="padding:10px 8px"><strong>Desktop publishing</strong>' +
      '<div style="color:#5c7a6a">' + order.layout.pages + ' pages per target language</div></td>' +
      '<td style="padding:10px 8px;text-align:right;color:#6e768c">' + count + ' Qt</td>' +
      '<td style="padding:10px 8px;text-align:right">' + money(layoutTotal(order), order.currency) +
      '</td></tr>';
  }

  var html =
    '<html><body style="font:11pt Helvetica,Arial,sans-serif;color:#111;margin:40px">' +
    '<table style="width:100%"><tr>' +
    '<td><span style="font-size:22pt;font-weight:700;color:' + navy + '">DEMA</span>' +
    '<span style="font-size:22pt;color:' + navy + '"> Solutions</span></td>' +
    '<td style="text-align:right;font-size:8pt;color:' + navy + '">' +
    '<strong>' + escape(COMPANY.name) + '</strong><br>' + escape(COMPANY.address) + '<br>' +
    'Vat Id ' + COMPANY.vat + ' - C.F. ' + COMPANY.taxCode + '</td></tr></table>' +
    '<div style="text-align:right;margin-top:18px;color:' + navy + '">' +
    '<strong style="font-size:12pt">ORDER CONFIRMATION</strong><br>no. ' + escape(order.id) +
    ' of ' + Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd') + '</div>' +
    '<hr style="border:none;border-top:1px solid ' + navy + ';margin:16px 0">' +
    '<table style="width:100%;font-size:9pt"><tr style="vertical-align:top">' +
    '<td style="width:48%"><strong style="color:' + navy + '">VAT ID</strong> ' +
    escape((order.billing || {}).vatNumber || (order.billing || {}).countryName || '-') +
    '<br><strong style="color:' + navy + '">EMAIL</strong> ' + escape(order.customer.email) + '</td>' +
    '<td><div style="color:' + navy + ';font-weight:700;font-size:8pt">RECIPIENT</div>' + block + '</td>' +
    '</tr></table>' +
    '<hr style="border:none;border-top:1px solid ' + navy + ';margin:16px 0">' +
    '<div style="color:' + navy + ';font-weight:700;font-size:8pt">SUBJECT</div>' +
    '<div style="font-size:12pt;font-weight:700;margin-bottom:18px">Order ' + escape(order.id) +
    (order.project ? ' - ' + escape(order.project) : '') + '</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:9.5pt">' +
    '<tr style="color:' + navy + ';font-size:7.5pt;font-weight:700;border-bottom:1px solid ' + navy + '">' +
    '<td style="padding:6px 8px">DESCRIPTION</td><td style="padding:6px 8px;text-align:right">QUANTITY</td>' +
    '<td style="padding:6px 8px;text-align:right">AMOUNT</td></tr>' + rows + '</table>' +
    '<table style="width:100%;margin-top:26px;font-size:9.5pt"><tr style="vertical-align:top">' +
    '<td style="width:48%"><div style="color:' + navy + ';font-weight:700;font-size:8pt">PAYMENT METHOD</div>' +
    escape(COMPANY.paymentMethod) + '<br><span style="color:#6e768c">IBAN: ' + COMPANY.iban +
    '<br>SWIFT: ' + COMPANY.swift + '</span>' +
    '<div style="color:' + navy + ';font-weight:700;font-size:8pt;margin-top:16px">DELIVERY</div>' +
    escape(when(order.eta)) + '</td>' +
    '<td><div style="color:' + navy + ';font-weight:700;font-size:8pt;border-bottom:1px solid ' + navy +
    ';padding-bottom:4px">VAT SUMMARY</div>' +
    '<table style="width:100%;font-size:8.5pt;margin-top:8px"><tr style="color:' + navy + ';font-weight:700">' +
    '<td>TOTAL AMOUNT</td><td style="text-align:right">TAXES</td></tr>' +
    '<tr><td colspan="2" style="color:#6e768c">' + escape(order.vatTreatment || '') + '</td></tr>' +
    '<tr><td>' + money(net, order.currency) + '</td><td style="text-align:right">' +
    money(vat, order.currency) + '</td></tr></table>' +
    '<div style="border-top:1px solid #ccd2e1;margin-top:10px;padding-top:10px;color:' + navy +
    ';font-weight:700;font-size:11pt">Total amount ' +
    '<span style="float:right;color:#111;font-size:12pt">' + money(gross, order.currency) + '</span></div>' +
    '</td></tr></table>' +
    '<hr style="border:none;border-top:1px solid ' + navy + ';margin:40px 0 8px">' +
    '<table style="width:100%;font-size:8pt;color:#6e768c"><tr style="vertical-align:top">' +
    '<td><strong style="color:' + navy + '">' + escape(COMPANY.name) + '</strong><br>' +
    COMPANY.site + ' - ' + COMPANY.email + '<br>tel: ' + COMPANY.phone + '</td>' +
    '<td style="text-align:right">Order confirmation ' + escape(order.id) + ' - 1 / 1<br>' +
    'This is an order confirmation, not a fiscal invoice.<br>' + escape(COMPANY.terms) + '</td>' +
    '</tr></table></body></html>';

  return Utilities.newBlob(html, MimeType.HTML, 'order-confirmation-' + order.id + '.html')
    .getAs(MimeType.PDF)
    .setName('order-confirmation-' + order.id + '.pdf');
}

function table(rows) {
  return '<table style="border-collapse:collapse;font:14px Helvetica,Arial,sans-serif">' +
    rows.map(function (r) {
      return '<tr><td style="padding:6px 18px 6px 0;color:#667;vertical-align:top">' + escape(r[0]) +
        '</td><td style="padding:6px 0"><strong>' + escape(r[1]) + '</strong></td></tr>';
    }).join('') + '</table>';
}

function shell(title, body) {
  return '<div style="max-width:620px;font:14px/1.6 Helvetica,Arial,sans-serif;color:#131a38">' +
    '<h1 style="font-size:20px;margin:0 0 16px">' + escape(title) + '</h1>' + body +
    '<p style="margin-top:24px;color:#667;font-size:12px">' + escape(COMPANY.name) + ' · ' +
    COMPANY.site + '</p></div>';
}

function customerBody(order, net, vat, gross) {
  return shell('Thank you - your order is in',
    '<p>We have your order and the confirmation is attached. The short version:</p>' +
    table([
      ['Reference', order.id],
      ['Solution', order.solution],
      ['Into', (order.targets || []).join(', ')],
      ['Delivery', when(order.eta)],
      ['Total due', money(gross, order.currency)]
    ]) +
    '<p>' + escape(order.vatTreatment || '') + '</p>' +
    '<p>' + escape(COMPANY.terms) + '</p>' +
    '<p>Any question, answer this email and it reaches the desk handling the job.</p>');
}

function deskBody(order, net, vat, gross, repriced) {
  var b = order.billing || {};
  var rows = [
    ['Reference', order.id],
    ['Project', order.project || '-'],
    ['Solution', order.solution + ' (' + (order.stages || []).join(' > ') + ')'],
    ['From', order.source],
    ['Into', (order.targets || []).join(', ')],
    ['Subject', order.subject],
    ['Words', (Number(order.words) || 0) * (order.targets || []).length],
    ['Desktop publishing', order.layout ? order.layout.pages + ' pages per language' : 'no'],
    ['Delivery promised', when(order.eta)],
    ['Net', money(net, order.currency)],
    ['VAT', vat ? money(vat, order.currency) : 'none'],
    ['Total due', money(gross, order.currency)],
    ['VAT treatment', order.vatTreatment || '-'],
    ['Customer', order.customer.name],
    ['Email', order.customer.email],
    ['Billing name', b.company ? b.companyName + ' (company)' : (b.name || '-') + ' (individual)'],
    ['Address', [b.address1, b.address2, b.postalCode, b.city, b.countryName].filter(String).join(', ') || '-'],
    ['VAT number', b.vatNumber || '-'],
    ['Telephone', b.telephone || '-'],
    ['Customer reference', order.customer.reference || '-'],
    ['Notes for the translator', order.customer.note || '-']
  ];
  if (repriced) {
    rows.unshift(['Check this', 'the page sent a total below the list price; the list price was used']);
  }
  return shell('New order · ' + order.id, table(rows));
}

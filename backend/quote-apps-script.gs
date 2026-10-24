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

/*
 * Where the customer's files and the order's paperwork are kept.
 *
 * The quote page counts words in the browser and used to send nothing but the
 * number: an order arrived with no documents to translate, which is most of a
 * translation job missing. The files now follow the order into a Drive folder
 * of their own, one per order, and the confirmation is filed beside them so
 * there is a record of what was sent and to whom.
 *
 * The parent folder is created on first use and its id kept in Script
 * properties, so it survives and is not made again.
 */
var LIBRARY_NAME = 'Quote Desk — Orders';
var MAX_UPLOAD_BYTES = 20 * 1024 * 1024;   // per file, before base64
var MAX_FILES = 30;

function libraryFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('ORDERS_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (err) {}
  }
  var found = DriveApp.getFoldersByName(LIBRARY_NAME);
  var folder = found.hasNext() ? found.next() : DriveApp.createFolder(LIBRARY_NAME);
  props.setProperty('ORDERS_FOLDER_ID', folder.getId());
  return folder;
}

/** The folder of one order, made the first time it is asked for. */
function orderFolder(id, customerName, code, project) {
  var library = libraryFolder();

  /* The uploads that follow an order know only its reference, and the folder
     is not named after it - so the folder is found through the row, which
     carries its link. A name can be edited; a link in the sheet cannot drift. */
  var row = findOrderRow(id);
  if (row) {
    var link = String(row[HEADERS.indexOf('Files')] || '');
    var match = /folders\/([A-Za-z0-9_-]+)/.exec(link);
    if (match) {
      try { return DriveApp.getFolderById(match[1]); } catch (err) {}
    }
  }

  /* "mara_bianchi - manuale_cnc_2026-10-24": the customer first, so one
     customer's work sits together in the listing, then what the job is and the
     day it was asked for. The reference is not in the name - the sheet holds
     the folder's link, which is how an upload finds it again. */
  var day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var what = slug(project) || slug(id);
  var name = [code || slug(customerName), what + '_' + day]
    .filter(function (part) { return part; }).join(' - ');

  var folder = library.createFolder(name);
  folder.setDescription('Quote Desk order ' + id + ', received ' + new Date().toISOString());
  return folder;
}
var DESK = 'g.gambuto@dema-solutions.com';   // where the desk's copy goes
var SENDER_NAME = 'DEMA Solutions';

/*
 * Who the confirmation comes from.
 *
 * Apps Script sends as the Google account it was deployed from, which is a
 * Gmail address. The customer should see the desk's own address instead, and
 * Microsoft Graph is how that happens - see graphCredentials() below for the
 * one-off setup. Without it the Gmail account is used and nothing breaks.
 */
var SENDER_FROM = 'g.gambuto@dema-solutions.com';

/*
 * Sending as g.gambuto@dema-solutions.com, through Microsoft 365 itself.
 *
 * Apps Script sends as the Google account it was deployed from, which is a
 * Gmail address - poor on a document a client reads. The obvious roads are
 * shut: Gmail will send as another address only through that address's own
 * SMTP, and this tenant answers 535 5.7.3 to basic SMTP authentication, which
 * is policy and not a per-mailbox switch.
 *
 * Microsoft Graph is the way that costs nothing and brings in no third party:
 * the mail leaves the real mailbox, on Microsoft's own servers, with the
 * tenant's own SPF and DKIM behind it.
 *
 * It needs one thing done once, by someone who administers the tenant:
 *
 *  1. Entra admin centre > App registrations > New registration. Any name.
 *  2. Certificates & secrets > New client secret. Copy the VALUE, not the id.
 *  3. API permissions > Microsoft Graph > Application permissions >
 *     Mail.Send > Add, then "Grant admin consent". Application, not delegated.
 *  4. Overview gives the Application (client) ID and the Directory (tenant) ID.
 *  5. Put all three in Project Settings > Script properties, never in this
 *     file: GRAPH_TENANT, GRAPH_CLIENT, GRAPH_SECRET.
 *
 * Narrow it afterwards if you like: an application access policy can limit the
 * app to this one mailbox rather than every mailbox in the tenant.
 *
 * With any of the three missing nothing changes: the mail goes out through the
 * Gmail account as before, and the desk's copy says so.
 */
function graphCredentials() {
  try {
    var props = PropertiesService.getScriptProperties();
    var tenant = props.getProperty('GRAPH_TENANT');
    var client = props.getProperty('GRAPH_CLIENT');
    var secret = props.getProperty('GRAPH_SECRET');
    if (tenant && client && secret) {
      return { tenant: tenant, client: client, secret: secret };
    }
  } catch (err) {}
  return null;
}

/** A token for the app itself, cached for the hour Microsoft gives it. */
function graphToken(creds) {
  var cache = CacheService.getScriptCache();
  var kept = cache.get('graph_token');
  if (kept) return kept;

  var res = UrlFetchApp.fetch(
    'https://login.microsoftonline.com/' + encodeURIComponent(creds.tenant) + '/oauth2/v2.0/token', {
      method: 'post',
      payload: {
        client_id: creds.client,
        client_secret: creds.secret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      },
      muteHttpExceptions: true
    });
  if (res.getResponseCode() !== 200) {
    throw new Error('token HTTP ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  }
  var body = JSON.parse(res.getContentText());
  /* A minute short of what Microsoft says, so a token never expires in flight. */
  cache.put('graph_token', body.access_token, Math.max(60, (body.expires_in || 3600) - 60));
  return body.access_token;
}

function graphSend(creds, options) {
  var attachments = (options.attachments || []).map(function (blob) {
    return {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: blob.getName(),
      contentType: blob.getContentType ? blob.getContentType() : 'application/octet-stream',
      contentBytes: Utilities.base64Encode(blob.getBytes())
    };
  });

  var message = {
    subject: options.subject,
    body: { contentType: 'HTML', content: options.htmlBody },
    toRecipients: [{ emailAddress: { address: options.to } }]
  };
  if (options.replyTo) message.replyTo = [{ emailAddress: { address: options.replyTo } }];
  if (attachments.length) message.attachments = attachments;

  var res = UrlFetchApp.fetch(
    'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(SENDER_FROM) + '/sendMail', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + graphToken(creds) },
      payload: JSON.stringify({ message: message, saveToSentItems: true }),
      muteHttpExceptions: true
    });
  var code = res.getResponseCode();
  /* Graph answers 202 and an empty body when it has taken the message. */
  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ' ' + res.getContentText().slice(0, 300));
  }
}

/*
 * The token the back office reads orders with.
 *
 * It lives in Project Settings > Script properties under ADMIN_TOKEN, never in
 * this repository. With no token set the reading endpoint stays shut, so a
 * half-finished deployment cannot hand the order book to the open internet.
 */
/*
 * Who may read the order book.
 *
 * A name and a password rather than one shared string, because the desk is a
 * few people and a login is what they expect to see. Be clear about what this
 * is worth: the page is a static file with no server of its own, so the only
 * thing that can really decide is this script. The pair is checked here, over
 * HTTPS, and nothing but the answer ever reaches the page. That is the same
 * strength the single token had - two boxes instead of one, and a habit people
 * already have.
 *
 * With ADMIN_USER or ADMIN_PASS unset the book stays shut.
 */
function adminLogin(user, pass) {
  var props;
  try { props = PropertiesService.getScriptProperties(); } catch (err) { return false; }
  var wantUser = props.getProperty('ADMIN_USER') || '';
  var wantPass = props.getProperty('ADMIN_PASS') || '';
  if (!wantUser || !wantPass) return false;
  return sameSecret(String(user || '').trim().toLowerCase(), wantUser.trim().toLowerCase()) &&
    sameSecret(String(pass || ''), wantPass);
}

function adminToken() {
  try {
    return PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN') || '';
  } catch (err) {
    return '';
  }
}

/** Compare without leaking the answer through how long it took. */
function sameSecret(given, want) {
  if (!want || !given || given.length !== want.length) return false;
  var diff = 0;
  for (var i = 0; i < want.length; i++) {
    diff |= want.charCodeAt(i) ^ given.charCodeAt(i);
  }
  return diff === 0;
}

function sameToken(given) {
  var want = adminToken();
  return sameSecret(given, want);
}

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

    /*
     * The same /exec serves two callers. The quote page posts an order and has
     * no action; the back office posts {action: 'orders', token} and gets the
     * book back. The token travels in the body on purpose: a query string ends
     * up in browser history and in referrers, and Apps Script cannot read a
     * custom header (a custom header would force a CORS preflight, and the web
     * app answers no OPTIONS).
     */
    if (order.action === 'orders') {
      /* Either a name and a password, or the older shared token: both are
         checked here and nowhere else. */
      var allowed = order.user
        ? adminLogin(order.user, order.pass)
        : sameToken(order.token);
      if (!allowed) return json({ ok: false, error: 'refused' });
      return json({ ok: true, headers: HEADERS, orders: orders() });
    }

    if (order.action === 'upload') return json(receiveFile(order));
    if (order.action === 'status') return json(orderStatus(order));

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

    /* The folder first, so the row can carry its link and the uploads that
       follow have somewhere to land. */
    var folder = orderFolder(order.id, order.customer.name, customerCode(order), order.project);
    record(order, net, vat, gross, repriced, folder.getUrl());

    var pdf = confirmationPdf(order, net, vat, gross);
    /* Filed, not only attached: an order confirmation nobody kept a copy of is
       a document that cannot be produced when it is asked for. */
    try { folder.createFile(pdf); } catch (err) {}
    var senderTrouble = send({
      to: order.customer.email,
      name: SENDER_NAME,
      // The account that sends is not the account that handles the job, so an
      // answer has to be pointed at the desk rather than at the sending mailbox.
      replyTo: DESK,
      subject: 'Your order ' + order.id + ' - ' + order.solution,
      htmlBody: customerBody(order, net, vat, gross),
      attachments: [pdf]
    });
    send({
      to: DESK,
      name: SENDER_NAME,
      replyTo: order.customer.email,
      subject: 'New order ' + order.id + ' - ' + order.solution + ' - ' +
        gross.toFixed(2) + ' ' + order.currency,
      htmlBody: deskBody(order, net, vat, gross, repriced, senderTrouble),
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

/**
 * Send one message: through Microsoft 365 when the tenant is configured,
 * otherwise through the Gmail account the script was deployed from.
 *
 * An order must never fail to reach anyone over a sender address, so a refusal
 * falls back to Gmail and says why. The desk's copy carries that sentence, so
 * a broken sender turns up in the inbox rather than in a customer's complaint.
 *
 * @returns {string|null} why it fell back, or null when all was well
 */
function send(options) {
  var creds = graphCredentials();
  if (!creds) {
    MailApp.sendEmail(options);
    return null;
  }
  try {
    graphSend(creds, options);
    return null;
  } catch (err) {
    MailApp.sendEmail(options);
    return 'Microsoft 365 refused the message, so it went out from the Gmail account instead (' +
      err + ')';
  }
}

function json(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* --------------------------------------------------------------- the status */

/**
 * What one customer may know about one order.
 *
 * The reference alone is not enough: it is printed on a confirmation that may
 * be forwarded, and sequential references are easy to guess at. The email has
 * to match the row as well, and only the handful of fields the customer
 * already knows come back - never the billing block, never another order.
 *
 * A wrong pair is answered exactly as an unknown one, so this cannot be used
 * to find out which references exist.
 */
function orderStatus(req) {
  var id = String(req.id || '').trim();
  var email = String(req.email || '').trim().toLowerCase();
  if (!id || !email) return { ok: false, error: 'not found' };

  var row = findOrderRow(id);
  if (!row) return { ok: false, error: 'not found' };

  var at = function (name) { return row[HEADERS.indexOf(name)]; };
  if (String(at('Email')).trim().toLowerCase() !== email) return { ok: false, error: 'not found' };

  var iso = function (value) {
    return Object.prototype.toString.call(value) === '[object Date]' ? value.toISOString() : '';
  };

  return {
    ok: true,
    order: {
      reference: id,
      status: String(at('Status') || ''),
      received: iso(at('Received')),
      delivery: iso(at('Delivery')),
      project: String(at('Project') || ''),
      solution: String(at('Solution') || ''),
      from: String(at('From') || ''),
      into: String(at('Into') || ''),
      words: Number(at('Words')) || 0,
      total: Number(at('Total')) || 0,
      currency: String(at('Currency') || 'EUR')
    }
  };
}

/* ------------------------------------------------------------- the delivery */

var DELIVERY_FOLDER = 'delivery';
var MAIL_ATTACHMENT_LIMIT = 24 * 1024 * 1024;   // MailApp refuses much above 25MB

/**
 * A menu in the sheet, because the desk lives in the sheet.
 *
 * Runs on open; the first time, open the script editor and run onOpen by hand
 * once to authorise it.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quote Desk')
    .addItem('Deliver the selected order', 'deliverSelected')
    .addToUi();
}

/**
 * Send the finished translation to the customer and mark the order delivered.
 *
 * The desk puts the finished files in a "delivery" folder inside the order's
 * own folder, picks any cell on that order's row, and runs this. The files are
 * attached rather than shared by link: the order folder also holds the
 * customer's own documents and the confirmation, and nothing here should
 * quietly make those readable by anyone holding a URL.
 */
function deliverSelected() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET) {
    ui.alert('Open the "' + SHEET + '" sheet and select a row on the order first.');
    return;
  }

  var rowIndex = sheet.getActiveRange().getRow();
  if (rowIndex < 2) { ui.alert('Select a row on an order, not the heading.'); return; }

  var row = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];
  var at = function (name) { return row[HEADERS.indexOf(name)]; };
  var id = String(at('Reference'));
  var email = String(at('Email'));
  if (!id || !email) { ui.alert('That row carries no reference or no email.'); return; }

  var folder = orderFolder(id, '');
  var out = folder.getFoldersByName(DELIVERY_FOLDER);
  if (!out.hasNext()) {
    ui.alert('Put the finished files in a "' + DELIVERY_FOLDER + '" folder inside ' +
      id + ' first.' + String.fromCharCode(10, 10) + folder.getUrl());
    return;
  }

  var files = [];
  var bytes = 0;
  var walk = out.next().getFiles();
  while (walk.hasNext()) {
    var file = walk.next();
    bytes += file.getSize();
    files.push(file.getBlob());
  }
  if (!files.length) { ui.alert('That delivery folder is empty.'); return; }
  if (bytes > MAIL_ATTACHMENT_LIMIT) {
    ui.alert('Those files come to ' + Math.round(bytes / 1048576) + ' MB, too much for one ' +
      'email. Send them from Drive and set the status by hand.');
    return;
  }

  var answer = ui.alert('Send ' + files.length + ' file(s) to ' + email + ' for ' + id + '?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  send({
    to: email,
    name: SENDER_NAME,
    replyTo: DESK,
    subject: 'Your translation is ready - ' + id,
    htmlBody: shell('Your translation is ready',
      '<p>' + escape(String(at('First name'))) + ',</p>' +
      '<p>the translation of <strong>' + escape(String(at('Project') || id)) + '</strong> ' +
      'from ' + escape(String(at('From'))) + ' into ' + escape(String(at('Into'))) +
      ' is attached.</p>' +
      '<p>Thank you for working with us.</p>'),
    attachments: files
  });

  sheet.getRange(rowIndex, HEADERS.indexOf('Status') + 1).setValue('Delivered');
  ui.alert('Sent to ' + email + ', and the order is marked delivered.');
}


/* ------------------------------------------------------- the customer's files */

/**
 * Take one file and put it in its order's folder.
 *
 * One file per request on purpose: a web app has a request size limit and a six
 * minute clock, and a customer sending twelve files should not lose all twelve
 * because the eleventh was large. The page sends them one at a time and can
 * report which arrived.
 *
 * The guard is the order reference: a file is only accepted against a row that
 * is really in the sheet, and only while the order is young. It is not a
 * password - the reference travels in the customer's own confirmation - but it
 * stops this being an open drop box for anyone who finds the URL.
 */
function receiveFile(req) {
  var id = String(req.id || '');
  if (!/^[A-Za-z0-9-]{4,40}$/.test(id)) return { ok: false, error: 'bad reference' };

  var row = findOrderRow(id);
  if (!row) return { ok: false, error: 'unknown order' };

  var received = Object.prototype.toString.call(row[0]) === '[object Date]' ? row[0].getTime() : 0;
  if (!received || Date.now() - received > 6 * 3600 * 1000) {
    return { ok: false, error: 'order closed to uploads' };
  }

  /* Drive has no paths, so a name cannot escape anywhere - but a file called
     ".._.._evil.sh" is still a nuisance to work with. Keep letters, digits,
     spaces and the ordinary punctuation of a filename; collapse the rest. */
  var name = String(req.name || 'file')
    .replace(/[^A-Za-z0-9 ._()À-ɏ-]/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^[._]+/, '')
    .slice(0, 150) || 'file';
  var data = String(req.data || '');
  if (!data) return { ok: false, error: 'empty file' };
  if (data.length > MAX_UPLOAD_BYTES * 1.4) return { ok: false, error: 'file too large' };

  var folder = orderFolder(id, '');
  var count = 0;
  var walk = folder.getFiles();
  while (walk.hasNext() && count <= MAX_FILES) { walk.next(); count++; }
  if (count > MAX_FILES) return { ok: false, error: 'too many files' };

  var blob = Utilities.newBlob(Utilities.base64Decode(data),
    req.mime || 'application/octet-stream', name);
  folder.createFile(blob);
  return { ok: true, stored: name };
}

/** The row of one order, by its reference, or null. */
function findOrderRow(id) {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(SHEET);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][1]) === id) return rows[i];
  }
  return null;
}


/* ───────────────────────────────────────────────────────────────── the sheet */

/*
 * Status is written by hand in the sheet as the job moves, and read back by the
 * back office; the page only ever sets the first value. Files is the link to
 * the order's Drive folder - the customer's documents, the confirmation, and
 * where the finished translation is put when it goes back.
 */
var STATUS_NEW = 'Received';

/*
 * Everything the order carried, in one row - who, where, what, when, how much.
 * The billing address used to be dropped on the floor: the row kept the company
 * and the country and threw away the street, the town, the postcode and the
 * telephone, which is most of an address. An order you cannot invoice is not a
 * record of anything.
 */
var HEADERS = ['Received', 'Reference', 'Status', 'Customer code',
  'First name', 'Last name', 'Email', 'Telephone',
  'Company', 'VAT number', 'Address', 'City', 'Postal code', 'Country',
  'Project', 'Solution', 'From', 'Into', 'Words', 'Subject', 'Desktop publishing', 'Delivery',
  'Net', 'VAT', 'Total', 'Currency', 'VAT treatment', 'Re-priced',
  'Customer reference', 'Notes', 'Files'];

/**
 * A short, stable name for the customer: mara_bianchi for a person, the
 * company slugged for a company. The same customer ordering again gets the
 * same code, which is the whole point of having one.
 *
 * Accents are folded rather than dropped, so Lundqvist and Lundqvíst do not
 * become two customers.
 */
function slug(text) {
  return String(text || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 50);
}

function customerCode(order) {
  var b = order.billing || {};
  if (b.company && b.companyName) return slug(b.companyName);
  var who = splitName(order);
  return slug(who[0] + ' ' + who[1]) || slug(order.customer.email.split('@')[0]);
}

/** First and last name, however the page happened to send them. */
function splitName(order) {
  var c = order.customer || {};
  if (c.firstName || c.lastName) return [c.firstName || '', c.lastName || ''];
  var whole = String(c.name || '').trim();
  var gap = whole.indexOf(' ');
  return gap === -1 ? [whole, ''] : [whole.slice(0, gap), whole.slice(gap + 1)];
}

function record(order, net, vat, gross, repriced, folderUrl) {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(SHEET) || book.insertSheet(SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    /*
     * The columns have grown - the address used to be thrown away - and a sheet
     * started under the old set would keep its old heading for ever, so the back
     * office would read every column by the wrong name.
     *
     * The heading is brought up to date here. Rows written before the change
     * keep the shape they were written in and will read crooked: there is no
     * honest way to invent an address that was never recorded. Move them to a
     * sheet of their own, or delete them if they were only tests.
     */
    var width = Math.max(sheet.getLastColumn(), HEADERS.length);
    var heading = sheet.getRange(1, 1, 1, width).getValues()[0];
    var current = heading.slice(0, HEADERS.length).join('');
    if (current !== HEADERS.join('')) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }
  }
  var b = order.billing || {};
  var who = splitName(order);
  var street = [b.address1, b.address2].filter(function (x) { return x; }).join(', ');

  sheet.appendRow([
    new Date(), order.id, STATUS_NEW, customerCode(order),
    who[0], who[1], order.customer.email, b.telephone || '',
    b.company ? b.companyName : 'individual', b.vatNumber || '',
    street, b.city || '', b.postalCode || '', b.countryName || b.country || '',
    order.project || '', order.solution, order.source, (order.targets || []).join(', '),
    (Number(order.words) || 0) * (order.targets || []).length, order.subject,
    order.layout ? order.layout.pages + ' pages' : 'no',
    order.eta ? new Date(order.eta) : '', net, vat, gross, order.currency,
    order.vatTreatment || '', repriced ? 'yes' : '', order.customer.reference || '',
    order.customer.note || '', folderUrl || ''
  ]);
}

/**
 * The order book, newest first, as rows in the order of HEADERS.
 *
 * Dates come back as ISO strings so the back office can format them in the
 * reader's own locale; everything else is handed over as the sheet holds it.
 */
function orders() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    out.push(rows[i].map(function (cell) {
      if (Object.prototype.toString.call(cell) === '[object Date]') return cell.toISOString();
      return cell;
    }));
  }
  return out;
}

/* ──────────────────────────────────────────────────────────── the paperwork */

/** 5000 -> 5,000: a document a client reads should set its numbers properly. */
function thousands(n) {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

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
    ' - ' + thousands((Number(order.words) || 0) * count) + ' words - ' + escape(order.solution) + '</div></td>' +
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
    '<td><img src="' + LETTERHEAD + '" width="170" alt="DEMA Solutions"></td>' +
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

function deskBody(order, net, vat, gross, repriced, senderTrouble) {
  var b = order.billing || {};
  var rows = [
    ['Reference', order.id],
    ['Project', order.project || '-'],
    ['Solution', order.solution + ' (' + (order.stages || []).join(' > ') + ')'],
    ['From', order.source],
    ['Into', (order.targets || []).join(', ')],
    ['Subject', order.subject],
    ['Words', thousands((Number(order.words) || 0) * (order.targets || []).length)],
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
  if (senderTrouble) {
    rows.unshift(['Check this', senderTrouble]);
  }
  return shell('New order · ' + order.id, table(rows));
}

/* The letterhead, exactly the mark Dema Solutions' invoices carry, kept here as
   data so the document never depends on fetching an image at send time. */
var LETTERHEAD = 'data:image/jpeg;base64,' +
  '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8Q' +
  'EBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQ' +
  'EBAQEBAQEBD/wAARCADIAlgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUF' +
  'BAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVW' +
  'V1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi' +
  '4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAEC' +
  'AxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVm' +
  'Z2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq' +
  '8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKK' +
  'KACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACi' +
  'iigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigA' +
  'ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo' +
  'AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiikyMhcjJ5AoAWiiigAooooAKKKKA' +
  'CiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiii' +
  'gAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAK2palp2jaddavrGoW1jYWML3' +
  'F1dXMqxQwRIpZ5HdiFVVAJJJAABJr5tb9oP4v/Hy8m0v9k7whZQeGY5GguPiN4rikj06QglZBp1oMS3bKc4kbbEHjZWGCGOPHb3v' +
  '7cvje8e+lmh/Z+8I6iba3ghkKDxzqcD/ADyu6n5tPhkXChTtkddxJK7YvqywsLHSrG30zTLKCzs7SJILe3gjEcUMSAKqIq4CqAAA' +
  'AMACvalTo5Ukq0VOt2fww9V9qXdPSOzTd0sburs7L8/+AfO8X7GNp4vRp/2gfjf8RfiZPcoFu7CTVn0fRXYdGjsLIoI+MZ/eNnFf' +
  'DH/BTn9nX4NfACL4Zw/CLwVF4fXWv7Z+3lby5uGn8n7F5WWnkcjb5snQj73Pav1wu7u1sLaW9vrmK3t4EMks0rhERR1ZmPAA9TX5' +
  'Y/8ABWv4mfDj4hyfC2LwB8QPDfiZ9LOui+XR9VgvTal/sOwS+UzbC3lvjdjOxsdDX03BWYZhis6oxlOXs/eulpD4JW91WitdtDmx' +
  'lOnGi3bX8dz4o8B/Gj4t/C+VJPh58SvEnh5UkEhhsNSligkYHI3xBtjjPZlIPcV9s/s+/wDBWjxno13baB+0RocXiDTXYI2vaVbp' +
  'b30Od3zywLiGYZKj5BEQAThzxX550V+u5nkGXZvBxxVJNvqlaS+a1/TujyadepSfus/o/wDAnj3wd8TfCth428A+IrPW9E1OPzba' +
  '8tXyrDurA4ZHU5DIwDKQQwBBFb9fg9+yD+1p4t/Zc8exX8Ul1qPg7VJkTXtGV8iWPp58IYhVnQcg5AYDYxAIK/ud4Y8S6F4z8O6Z' +
  '4t8MalFqGkaxaRXtldRZ2zQyKGRgDgjII4IBHQgGvwrifhmtw7iEr81KXwy/R+a/HddUvcw2JWIj5o06KKK+XOkKKKKACiiigAoo' +
  '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' +
  'bQqSkSAKAMkKmdo9cr57G4inTprBYV3gndy/nl3/AMK2iu129W0bwi2+eW/5f11CiiivLNQooooAKKKKACiiigAooooAKKKKACii' +
  'igAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAo' +
  'oooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigCG7tLW/tZrG+torm2uY2imhlQOkiMMMrKeCCCQQeCDX4r/ALfX7HU/' +
  '7N3jdfFngy0nk+Hnia4Y6exy/wDZd0QWayd+4wGaJm5ZAQSxjZj+11c38R/h54T+K/gjWPh5440tNQ0XW7c29zC2MjkFXQ/wujBX' +
  'VhyrKpHSvo+GeIavD+LVVa05aSj3Xdea6fd1OfE4dYiFuvQ/nFr9Mf8AgjN934u/7vh/+eo18U/tOfs5eMP2ZfideeA/EoN1YS5u' +
  'dF1VU2xajZk4VwP4ZF+66fwsDgspVm+xv+CRHiHRPCegfG3xP4k1ODTtK0q10O7vLudtscEKDUmZ2PoADX6/xhXp47hurWwz5oy5' +
  'LW63nH+rHk4SLhiEpaWv+R+hPxg+Lng34IeAdS+Injm9eKwsFCRQQrvuL25biK2gTI3yu3AHAHLMVVWYeY/s8/CLxZf+Jb79pb47' +
  'Wo/4WN4nthBp2lMxeHwno5JaOwhyP9cc7ppMAlmKgD5y/O/CPQNc/aj+INh+0z8S9FnsPBOhSM/ww8N3iYLKf+Y5dRnrNJgGAHiN' +
  'drrk7ZX+pK/FsRJZZSeEpv8Aey0m+y/kX/t76v3dk+b2I/vXzvbp/n/kFFFFeMbBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQ' +
  'AUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUU' +
  'UAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAeO/tS/s1+Ev2nvhjdeCteWO01a13XWhasFzJp95jAPq0T8LInR' +
  'lwRhlRl+D/2Jf2J/i3qPirxZ4W+M9he6D8PdJ1iz/t7RZV2jxLqFiZXtoc/8tLNPtBlZsmOTdFtDEb4v1Sor6DL+Jcbl2BqYCk/d' +
  'lZq+8Xe949n+T1VmYVMPCpNVHuvxGxxpEixRIqIgCqqjAAHQAU6iivnzcKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigA' +
  'ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo' +
  'AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKK' +
  'KACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACi' +
  'iigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigA' +
  'ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo' +
  'AKKKKACiiigAooooAKKKKAP/2Q==';

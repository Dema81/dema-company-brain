/**
 * Company Brain — waitlist backend (Google Apps Script)
 *
 * Deploy: Extensions > Apps Script from a Google Sheet, paste this file,
 * then Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the /exec URL into WEBHOOK_URL in index.html.
 *
 * GET  -> { taken, total }                      (used to show the real counter on load)
 * POST -> { nome, azienda, email, settore, lang } -> { taken, total }
 */

var SEED   = 6;                                // spots already taken before launch
var TOTAL  = 20;
var NOTIFY = 'g.gambuto@dema-solutions.com';
var SHEET  = 'leads';

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['data', 'nome', 'azienda', 'email', 'settore', 'lingua']);
  }
  return sh;
}

function taken_() {
  var rows = Math.max(0, sheet_().getLastRow() - 1); // minus the header
  return Math.min(SEED + rows, TOTAL);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ taken: taken_(), total: TOTAL });
}

function doPost(e) {
  var d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { d = {}; }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sheet_().appendRow([
      new Date(),
      d.nome    || '',
      d.azienda || '',
      d.email   || '',
      d.settore || '',
      d.lang    || ''
    ]);
    var n = taken_();
  } finally {
    lock.releaseLock();
  }

  try {
    MailApp.sendEmail({
      to: NOTIFY,
      replyTo: d.email || NOTIFY,
      subject: 'Nuovo lead Company Brain — ' + (d.azienda || d.nome || ''),
      htmlBody:
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#22252B">' +
        '<h2 style="font-size:16px;text-transform:uppercase;letter-spacing:.04em">Nuovo lead — Company Brain</h2>' +
        '<table cellpadding="6" cellspacing="0" style="border-collapse:collapse">' +
        row_('Nome', d.nome) + row_('Azienda', d.azienda) +
        '<tr><td style="color:#7D828C">Email</td><td><a href="mailto:' + (d.email || '') + '">' + (d.email || '') + '</a></td></tr>' +
        row_('Settore', d.settore) + row_('Lingua', d.lang) +
        '</table>' +
        '<p style="color:#7D828C;font-size:12px">Posti occupati: <b>' + n + '/' + TOTAL + '</b></p></div>'
    });
  } catch (err) {
    // never fail the response because of the mail: the page still needs its counter
  }

  return json_({ taken: n, total: TOTAL });
}

function row_(label, value) {
  return '<tr><td style="color:#7D828C">' + label + '</td><td><b>' + (value || '') + '</b></td></tr>';
}

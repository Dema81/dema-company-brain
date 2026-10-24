/*
 * The Apps Script, checked without deploying it.
 *
 *     node backend/quote-apps-script.test.js
 *
 * Google's own objects - DriveApp, SpreadsheetApp, MailApp, UrlFetchApp - are
 * stood in for below, so the script's logic can be run on an ordinary machine.
 * That is the point: this file is the only way to know the back office still
 * works before a customer meets it, because nothing here can be run locally
 * against the real Google.
 *
 * Run it after every change to quote-apps-script.gs, and before every
 * redeploy. It touches nothing real.
 */
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const path = require('path');
const target = process.argv[2] || path.join(__dirname, 'quote-apps-script.gs');
const src = fs.readFileSync(target, 'utf8');

let props = {}, sheetRows = [], drive = {}, created = [], heading = [], fetchCalls = [];
let stripeCode = 200, stripeBody = '{"url":"https://checkout.stripe.com/c/pay/abc"}';
const ids = {};
let nextId = 0;
const folderId = (name) => (ids[name] = ids[name] || ('1A' + (nextId++) + 'Zx_dem-A'));
const folderStub = (name) => ({
  _name: name, _files: drive[name] = drive[name] || [],
  getId: () => folderId(name), getUrl: () => 'https://drive.google.com/drive/folders/' + folderId(name),
  setDescription() { return this; },
  getName: () => name.split('/').pop(),
  getFoldersByName: (n) => { const k = name + '/' + n; return { hasNext: () => !!drive[k], next: () => folderStub(k) }; },
  getFolders() { const kids = Object.keys(drive).filter(k => k.startsWith(name + '/') && k.split('/').length === name.split('/').length + 1);
    let i = 0; return { hasNext: () => i < kids.length, next: () => folderStub(kids[i++]) }; },
  createFolder(n) { const k = name + '/' + n; drive[k] = []; return folderStub(k); },
  getFiles() { let i = 0; const f = drive[name]; return { hasNext: () => i < f.length, next: () => f[i++] }; },
  createFile(blob) { drive[name].push(blob); created.push({ folder: name, name: blob.getName() }); return blob; }
});
const sandbox = {
  PropertiesService: { getScriptProperties: () => ({
    getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  DriveApp: {
    getFolderById: (id) => { const name = Object.keys(ids).find(k => ids[k] === id);
      if (!name) throw new Error('gone'); return folderStub(name); },
    getFoldersByName: (n) => ({ hasNext: () => false, next: () => null }),
    createFolder: (n) => folderStub(n)
  },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({
    getSheetByName: () => ({
      getLastRow: () => sheetRows.length + 1,
      getLastColumn: () => heading.length,
      getRange: (r, c, nr, nc) => ({
        getValues: () => (r === 1 ? [heading.slice(0, nc)] : sheetRows.map(x => x.slice(0, nc))),
        setValues: (v) => { heading = v[0].slice(); }
      }),
      appendRow: (row) => sheetRows.push(row), setFrozenRows() {}
    }), insertSheet: () => { throw new Error('no'); } }) },
  Session: { getScriptTimeZone: () => 'Europe/Rome' },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  DriveApp_: null,
  Utilities: {
    formatDate: (d, tz, fmt) => d.toISOString().slice(0, 10),
    base64Decode: (s) => Buffer.from(s, 'base64'),
    newBlob: (bytes, mime, name) => ({ getName: () => name, getBytes: () => bytes, _mime: mime }),
    base64Encode: (b) => Buffer.from(b).toString('base64')
  },
  ContentService: { MimeType: { JSON: 'json' },
    createTextOutput: (t) => ({ setMimeType: () => ({ body: t }) }) },
  MailApp: { sendEmail: () => {} }, UrlFetchApp: { fetch: (u, o) => { fetchCalls.push({ url: u, opts: o }); return { getResponseCode: () => stripeCode, getContentText: () => stripeBody }; } }, HtmlService: {}, Session: { getScriptTimeZone: () => 'Europe/Rome' },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  DriveApp_: null, MimeType: {}
};
const ctx = vm.createContext(sandbox);
vm.runInContext(src, ctx);
const { receiveFile, HEADERS, record, orderFolder, splitName, orderStatus, customerCode, slug, adminLogin, payLink, payParagraph } = ctx;

// ── the sheet keeps everything ──────────────────────────────────────────────
const order = {
  id: 'QD-0042', solution: 'Professional', source: 'en-GB', targets: ['fr-FR', 'de-DE'],
  words: 5000, subject: 'mechanical', layout: { pages: 20 }, eta: '2026-11-05T17:00:00.000Z',
  currency: 'EUR', vatTreatment: 'Italy, 22%',
  project: 'Manuale X', customer: { firstName: 'Mara', lastName: 'Bianchi',
    email: 'm@example.it', reference: 'PO-1', note: 'glossario allegato' },
  billing: { company: true, companyName: 'Esempio Srl', vatNumber: 'IT01234567890',
    telephone: '+39 039 000', address1: 'Via Gramsci 10', address2: 'Scala B',
    city: 'Monza', postalCode: '20900', countryName: 'Italy' }
};
record(order, 600, 132, 732, false, 'https://drive/x');
const row = sheetRows[0], at = (name) => row[HEADERS.indexOf(name)];
for (const [col, want] of [['Reference','QD-0042'], ['Status','Received'],
  ['First name','Mara'], ['Last name','Bianchi'], ['Email','m@example.it'],
  ['Telephone','+39 039 000'], ['Company','Esempio Srl'], ['VAT number','IT01234567890'],
  ['Address','Via Gramsci 10, Scala B'], ['City','Monza'], ['Postal code','20900'],
  ['Country','Italy'], ['Project','Manuale X'], ['Solution','Professional'],
  ['From','en-GB'], ['Into','fr-FR, de-DE'], ['Words',10000], ['Net',600], ['VAT',132],
  ['Total',732], ['Currency','EUR'], ['Customer reference','PO-1'],
  ['Notes','glossario allegato'], ['Files','https://drive/x']]) {
  assert.strictEqual(at(col), want, col + ': atteso ' + want + ', trovato ' + at(col));
}
assert.strictEqual(row.length, HEADERS.length, 'la riga ha una cella per colonna');

// a page that sends only a joined name still fills both columns
const names = (o) => JSON.parse(JSON.stringify(splitName(o)));
assert.deepStrictEqual(names({ customer: { name: 'Hans Peter Richter' } }), ['Hans', 'Peter Richter']);
assert.deepStrictEqual(names({ customer: { name: 'Prince' } }), ['Prince', '']);

// ── the upload ──────────────────────────────────────────────────────────────
sheetRows[0][0] = new Date();            // received just now
const b64 = Buffer.from('ciao').toString('base64');
let r = receiveFile({ id: 'QD-0042', name: 'manuale.docx', data: b64, mime: 'application/x' });
assert.strictEqual(r.ok, true, 'un file di un ordine valido viene accettato: ' + r.error);
assert.strictEqual(created.filter(c => c.name === 'manuale.docx').length, 1, 'ed e scritto su Drive');

assert.strictEqual(receiveFile({ id: 'QD-9999', name: 'x', data: b64 }).error, 'unknown order');
assert.strictEqual(receiveFile({ id: '../../etc', name: 'x', data: b64 }).error, 'bad reference');
assert.strictEqual(receiveFile({ id: 'QD-0042', name: 'x', data: '' }).error, 'empty file');
assert.strictEqual(receiveFile({ id: 'QD-0042', name: 'x', data: 'A'.repeat(40 * 1024 * 1024) }).error,
  'file too large');

// a name that tries to escape its folder is filed, not obeyed
created.length = 0;
receiveFile({ id: 'QD-0042', name: '../../../evil.sh', data: b64 });
assert.ok(!/[\/]/.test(created[0].name) && !/^[._]/.test(created[0].name),
  'il nome file non porta barre ne punti iniziali: ' + created[0].name);
assert.ok(created[0].name.indexOf('evil.sh') !== -1, 'ma resta riconoscibile: ' + created[0].name);

// an old order no longer accepts uploads
sheetRows[0][0] = new Date(Date.now() - 7 * 3600 * 1000);
assert.strictEqual(receiveFile({ id: 'QD-0042', name: 'x', data: b64 }).error, 'order closed to uploads');

// ── a sheet that already exists under the old heading ──────────────────────
heading = ['Received','Reference','Customer','Email','Company','Country','VAT number','Project',
  'Solution','From','Into','Words','Subject','Desktop publishing','Delivery','Net','VAT','Total',
  'Currency','VAT treatment','Re-priced','Reference (customer)','Notes'];
sheetRows = [['old', 'QD-0001']];
record(order, 1, 0, 1, false, '');
assert.deepStrictEqual(JSON.parse(JSON.stringify(heading)), JSON.parse(JSON.stringify(HEADERS)),
  'l intestazione vecchia viene portata a quella nuova');

// and a sheet already current is left alone
const before = heading.slice();
record(order, 1, 0, 1, false, '');
assert.deepStrictEqual(JSON.parse(JSON.stringify(heading)), JSON.parse(JSON.stringify(before)),
  'un intestazione gia corretta non viene riscritta');

// ── what a customer may see ────────────────────────────────────────────────
sheetRows = []; heading = HEADERS.slice();
record(order, 600, 132, 732, false, 'https://drive/x');
sheetRows[0][0] = new Date();

const seen = orderStatus({ id: 'QD-0042', email: 'M@Example.IT' });
assert.strictEqual(seen.ok, true, 'riferimento ed email giusti: ' + seen.error);
assert.strictEqual(seen.order.reference, 'QD-0042');
assert.strictEqual(seen.order.status, 'Received');
assert.strictEqual(seen.order.total, 732);
assert.strictEqual(seen.order.into, 'fr-FR, de-DE');
const shown = Object.keys(JSON.parse(JSON.stringify(seen.order)));
for (const secret of ['address', 'Address', 'telephone', 'vat', 'notes', 'city'])
  assert.ok(!shown.includes(secret), 'la fatturazione non esce: ' + secret);

// the reference alone is not a key, and a miss looks like any other miss
assert.deepStrictEqual(JSON.parse(JSON.stringify(orderStatus({ id: 'QD-0042', email: 'ladro@x.com' }))),
  { ok: false, error: 'not found' }, 'email sbagliata: respinto');
assert.deepStrictEqual(JSON.parse(JSON.stringify(orderStatus({ id: 'QD-9999', email: 'm@example.it' }))),
  { ok: false, error: 'not found' }, 'ordine inesistente: stessa risposta');
assert.deepStrictEqual(JSON.parse(JSON.stringify(orderStatus({ id: 'QD-0042' }))),
  { ok: false, error: 'not found' }, 'senza email: respinto');

// ── the folder is named for the customer, the job and the day ─────────
drive = {}; created.length = 0; props = {}; sheetRows = []; heading = HEADERS.slice();
record(order, 600, 132, 732, false, '');
const code = customerCode(order);
assert.strictEqual(code, 'esempio_srl', 'azienda: ragione sociale slugged, trovato ' + code);
assert.strictEqual(customerCode({ customer: { firstName: 'Mara', lastName: 'Bianchi', email: 'm@x.it' }, billing: {} }),
  'mara_bianchi', 'privato: nome_cognome');
assert.strictEqual(customerCode({ customer: { firstName: 'Ana', lastName: 'Lundqvíst', email: 'a@x.se' }, billing: {} }),
  'ana_lundqvist', 'gli accenti sono piegati, non due clienti diversi');

const made = orderFolder('QD-0042', 'Mara Bianchi', 'mara_bianchi', 'Manuale CNC');
const folderName = made.getName();
assert.ok(folderName.indexOf('mara_bianchi') === 0, 'il cliente viene per primo: ' + folderName);
assert.ok(folderName.indexOf('manuale_cnc') !== -1, 'poi il progetto: ' + folderName);
assert.ok(/\d{4}-\d{2}-\d{2}/.test(folderName), 'e la data: ' + folderName);

// an upload later knows only the reference, and finds the folder through the row
sheetRows[0][HEADERS.indexOf('Files')] = made.getUrl();
sheetRows[0][0] = new Date();
const again = orderFolder('QD-0042', '', '', '');
assert.strictEqual(again.getName(), folderName, 'ritrovata dal link nella riga, non ricreata');

// the sheet carries the code
assert.strictEqual(sheetRows[0][HEADERS.indexOf('Customer code')], 'esempio_srl',
  'il foglio registra il codice cliente');

// ── the admin login ─────────────────────────────────────────
props = {};
assert.strictEqual(adminLogin('desk', 'secret'), false, 'senza credenziali impostate: chiuso');
props = { ADMIN_USER: 'desk', ADMIN_PASS: 'lungo-e-segreto' };
assert.strictEqual(adminLogin('desk', 'lungo-e-segreto'), true, 'la coppia giusta entra');
assert.strictEqual(adminLogin('DESK', 'lungo-e-segreto'), true, 'il nome non guarda le maiuscole');
assert.strictEqual(adminLogin('desk', 'Lungo-e-segreto'), false, 'la password sì');
assert.strictEqual(adminLogin('altro', 'lungo-e-segreto'), false, 'nome sbagliato: fuori');
assert.strictEqual(adminLogin('desk', ''), false, 'password vuota: fuori');

// ── paying by card, which is off until a key exists ──────────────────
props = {}; fetchCalls = [];
assert.strictEqual(payLink(order, 732), '', 'senza chiave: nessun link, nessuna chiamata');
assert.strictEqual(fetchCalls.length, 0, 'e Stripe non viene nemmeno contattato');
assert.strictEqual(payParagraph(''), '', 'e la conferma non nomina la carta');

props = { STRIPE_SECRET: 'sk_test_x' }; fetchCalls = [];
const link = payLink(order, 732);
assert.strictEqual(link, 'https://checkout.stripe.com/c/pay/abc', 'con chiave: arriva il link');
const call = fetchCalls[0];
assert.strictEqual(call.url, 'https://api.stripe.com/v1/checkout/sessions');
assert.strictEqual(call.opts.headers.Authorization, 'Bearer sk_test_x');
assert.strictEqual(call.opts.payload['line_items[0][price_data][unit_amount]'], '73200',
  'l importo e in centesimi, non in euro');
assert.strictEqual(call.opts.payload['line_items[0][price_data][currency]'], 'eur');
assert.strictEqual(call.opts.payload.client_reference_id, 'QD-0042');
assert.ok(payParagraph(link).indexOf(link) !== -1, 'il paragrafo porta il link');

// a refusal must not cost the order
props = { STRIPE_SECRET: 'sk_test_x' }; stripeCode = 402;
stripeBody = '{"error":{"message":"card payments not activated"}}';
assert.strictEqual(payLink(order, 732), '', 'Stripe rifiuta: nessun link, ordine salvo');
stripeCode = 200; stripeBody = '{"url":"https://checkout.stripe.com/c/pay/abc"}';

// nothing to pay, nothing to offer
assert.strictEqual(payLink(order, 0), '', 'totale zero: nessun link');

console.log('49 controlli verdi: il foglio registra tutto, gli upload arrivano e i casi storti sono respinti');

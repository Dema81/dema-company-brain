/*
 * The Apps Script, checked without deploying it.
 *
 *     node backend/quote-apps-script.test.js
 *
 * Google's own objects - SpreadsheetApp, MailApp, UrlFetchApp, HtmlService -
 * are stood in for below, and so is OneDrive: UrlFetchApp routes Graph calls
 * into a small drive kept in memory, so the folder logic can be run on an
 * ordinary machine. That is the point: this file is the only way to know the
 * desk still works before a customer meets it, because nothing here can be run
 * locally against the real Google or the real Microsoft.
 *
 * Run it after every change to quote-apps-script.gs, and before every
 * redeploy. It touches nothing real.
 */
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const path = require('path');
const target = process.argv[2] || path.join(__dirname, 'quote-apps-script.gs');
const src = fs.readFileSync(target, 'utf8');

let props = {}, sheetRows = [], heading = [], fetchCalls = [], sentMail = [];
let stripeCode = 200, stripeBody = '{"url":"https://checkout.stripe.com/c/pay/abc"}';

/*
 * OneDrive, in a dictionary: path -> item. A folder holds nothing itself; what
 * is in it is worked out from the paths, exactly as Graph does it.
 */
let onedrive = {};
let nextItem = 0;
const itemFor = (p, isFolder, size) => {
  const name = p.split('/').pop();
  const above = p.split('/').slice(0, -1).join('/');
  return {
    id: 'ITEM' + (nextItem++),
    name,
    webUrl: 'https://dema-my.sharepoint.com/personal/x/Documents/' + p.replace(/ /g, '%20'),
    parentReference: { path: '/drive/root:' + (above ? '/' + above.split('/').map(encodeURIComponent).join('/') : '') },
    size: size || 0,
    folder: isFolder ? { childCount: 0 } : undefined,
    file: isFolder ? undefined : { mimeType: 'application/octet-stream' }
  };
};
const pathOfId = (id) => Object.keys(onedrive).find((p) => onedrive[p].id === id);
const decodePath = (escaped) => escaped.split('/').map(decodeURIComponent).join('/');

/* What the token endpoint hands back as a refresh token; a test moves it. */
let tokenEndpointGives = 'refresh-1';

/* Set by the test that checks what happens when the tenant has not consented:
   every Graph call answers 403, the way it does with no admin consent. */
let graphRefuses = false;

/* Every Graph call this script makes, answered the way Graph answers it. */
function graphRoute(url, opts) {
  const method = String((opts && opts.method) || 'get').toLowerCase();
  const reply = (code, body) => ({
    getResponseCode: () => code,
    getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body || {})),
    getBlob: () => ({ _name: 'downloaded', setName(n) { this._name = n; return this; },
      getName() { return this._name; }, getBytes: () => Buffer.from('bytes'),
      getContentType: () => 'application/octet-stream' })
  });

  if (url.indexOf('login.microsoftonline.com') !== -1) {
    const grant = (opts && opts.payload && opts.payload.grant_type) || '';
    /* Only the two delegated grants carry a refresh token back. */
    if (grant === 'authorization_code' || grant === 'refresh_token') {
      return reply(200, { access_token: 'graph-tok', expires_in: 3600,
        refresh_token: tokenEndpointGives });
    }
    return reply(200, { access_token: 'graph-tok', expires_in: 3600 });
  }

  if (graphRefuses) {
    return reply(403, { error: { code: 'accessDenied',
      message: 'Either scp or roles claim need to be present in the token.' } });
  }

  // a folder is made: .../root:/<parent>:/children  or  .../root/children
  let m = /\/root(?::\/(.*?))?:?\/children$/.exec(url.split('?')[0]);
  if (m && method === 'post') {
    const parent = m[1] ? decodePath(m[1]) : '';
    const body = JSON.parse(opts.payload);
    const full = parent ? parent + '/' + body.name : body.name;
    if (onedrive[full]) return reply(409, { error: { code: 'nameAlreadyExists' } });
    onedrive[full] = itemFor(full, true);
    return reply(201, onedrive[full]);
  }

  // an upload session
  m = /\/root:\/(.*):\/createUploadSession$/.exec(url);
  if (m) {
    const full = decodePath(m[1]);
    /* Graph echoes the expiry it accepted, so the stub echoes the one it was
       asked for. It is a request, not a guarantee: the real service may hand
       back a shorter one, and activity on the session is what keeps it alive. */
    const asked = JSON.parse((opts && opts.payload) || '{}');
    return reply(200, {
      uploadUrl: 'https://dema-my.sharepoint.com/_api/v2.0/upload/' + encodeURIComponent(full),
      expirationDateTime: asked.expirationDateTime ||
        new Date(Date.now() + 3600000).toISOString()
    });
  }

  // a small file written whole
  m = /\/root:\/(.*):\/content$/.exec(url);
  if (m && method === 'put') {
    const full = decodePath(m[1]);
    onedrive[full] = itemFor(full, false, 1024);
    return reply(201, onedrive[full]);
  }

  // what is in a folder
  m = /\/items\/([^/?]+)\/children/.exec(url);
  if (m) {
    const here = pathOfId(decodeURIComponent(m[1]));
    const kids = Object.keys(onedrive)
      .filter((p) => p.startsWith(here + '/') && p.slice(here.length + 1).indexOf('/') === -1)
      .map((p) => onedrive[p]);
    return reply(200, { value: kids });
  }

  // one item, by id
  m = /\/items\/([^/?]+)(\?|$)/.exec(url);
  if (m) {
    const here = pathOfId(decodeURIComponent(m[1]));
    if (method === 'delete') {
      Object.keys(onedrive).forEach((p) => {
        if (p === here || p.startsWith(here + '/')) delete onedrive[p];
      });
      return reply(204, '');
    }
    if (!here) return reply(404, { error: { code: 'itemNotFound' } });
    return reply(200, Object.assign({}, onedrive[here],
      { '@microsoft.graph.downloadUrl': 'https://dema-my.sharepoint.com/download/' + m[1] }));
  }

  // one item, by path
  m = /\/root:\/([^:?]+)$/.exec(url);
  if (m) {
    const full = decodePath(m[1]);
    return onedrive[full] ? reply(200, onedrive[full])
      : reply(404, { error: { code: 'itemNotFound' } });
  }

  if (url.indexOf('/download/') !== -1) return reply(200, 'bytes');
  return null;
}

/*
 * Google Drive, in memory.
 *
 * The fallback writes here when Graph refuses, so the tests need a Drive that
 * behaves like the real one in the two ways the code depends on: a folder
 * looked up by name is made once and found thereafter, and a file dropped in
 * a folder is listed back with its name and its size.
 */
const drive = { folders: {}, files: {} };

function driveFolderStub(id, name) {
  if (drive.folders[id]) return drive.folders[id];
  const made = {
    _id: id, _name: name, _children: [], _files: [],
    getId() { return this._id; },
    getName() { return this._name; },
    getUrl() { return 'https://drive.google.com/drive/folders/' + this._id; },
    getFoldersByName(wanted) {
      const hit = this._children.map((cid) => drive.folders[cid])
        .filter((f) => f.getName() === wanted);
      let i = 0;
      return { hasNext: () => i < hit.length, next: () => hit[i++] };
    },
    createFolder(childName) {
      const cid = this._id + '/' + childName;
      const child = driveFolderStub(cid, childName);
      child._parent = this._id;
      if (this._children.indexOf(cid) === -1) this._children.push(cid);
      return child;
    },
    getParents() {
      const up = this._parent ? [drive.folders[this._parent]] : [];
      let i = 0;
      return { hasNext: () => i < up.length, next: () => up[i++] };
    },
    createFile(blob) {
      const fid = 'file:' + this._id + '/' + blob.getName();
      drive.files[fid] = { id: fid, name: blob.getName(), blob };
      this._files.push(fid);
      return {
        getId: () => fid,
        setName(n) { drive.files[fid].name = n; return this; },
        getSize: () => 10
      };
    },
    getFiles() {
      const list = this._files.map((fid) => ({
        getId: () => fid,
        getName: () => drive.files[fid].name,
        getSize: () => 10
      }));
      let i = 0;
      return { hasNext: () => i < list.length, next: () => list[i++] };
    },
    getFolders() {
      const list = this._children.map((cid) => drive.folders[cid]);
      let i = 0;
      return { hasNext: () => i < list.length, next: () => list[i++] };
    }
  };
  drive.folders[id] = made;
  return made;
}

driveFolderStub('ORDERS-ROOT', 'Quote Desk — Orders');

const sandbox = {
  DriveApp: {
    getFolderById: (id) => {
      if (!drive.folders[id]) throw new Error('No item with the given ID could be found: ' + id);
      return drive.folders[id];
    },
    getFileById: (id) => {
      if (!drive.files[id]) throw new Error('No file ' + id);
      return { getBlob: () => drive.files[id].blob };
    }
  },
  PropertiesService: { getScriptProperties: () => ({
    getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; },
    deleteProperty: (k) => { delete props[k]; } }) },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({
    getSheetByName: () => ({
      getLastRow: () => sheetRows.length + 1,
      getLastColumn: () => heading.length,
      getRange: (r, c, nr, nc) => ({
        getValues: () => {
          if (r === 1) return [heading.slice(0, nc)];
          return sheetRows.slice(r - 2, r - 2 + nr).map((x) => x.slice(c - 1, c - 1 + nc));
        },
        setValues: (v) => { heading = v[0].slice(); },
        setValue: (v) => { sheetRows[r - 2][c - 1] = v; },
        getValue: () => (sheetRows[r - 2] || [])[c - 1]
      }),
      appendRow: (row) => sheetRows.push(row), setFrozenRows() {}
    }), insertSheet: () => { throw new Error('no'); } }) },
  Session: { getScriptTimeZone: () => 'Europe/Rome' },
  Logger: { log() {} },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/AAA/exec' }) },
  HtmlService: { createHtmlOutput: (html) => ({
    _html: html, getContent() { return this._html; },
    setTitle() { return this; }, addMetaTag() { return this; } }) },
  Utilities: {
    formatDate: (d, tz, fmt) => {
      if (fmt === 'yyyy') return d.toISOString().slice(0, 4);
      if (fmt === 'HH-mm') return d.toISOString().slice(11, 16).replace(':', '-');
      if (fmt === 'EEE d MMM, HH:mm') return d.toISOString().slice(0, 16);
      return d.toISOString().slice(0, 10);
    },
    base64Decode: (s) => Buffer.from(s, 'base64'),
    base64Encode: (b) => Buffer.from(b).toString('base64'),
    newBlob: (bytes, mime, name) => ({
      _name: name, _mime: mime,
      getName() { return this._name; }, setName(n) { this._name = n; return this; },
      getBytes: () => (Buffer.isBuffer(bytes) ? bytes : Buffer.from(String(bytes))),
      getContentType() { return this._mime; },
      getAs() { return this; }
    })
  },
  MimeType: { HTML: 'text/html', PDF: 'application/pdf' },
  ContentService: { MimeType: { JSON: 'json' },
    createTextOutput: (t) => ({ setMimeType: () => ({ body: t }) }) },
  MailApp: { sendEmail: (to, subject, body) => { sentMail.push({ to, subject, body }); } },
  UrlFetchApp: { fetch: (u, o) => {
    fetchCalls.push({ url: u, opts: o });
    const graph = graphRoute(u, o);
    if (graph) return graph;
    return { getResponseCode: () => stripeCode, getContentText: () => stripeBody,
      getBlob: () => ({ getName: () => 'x', getBytes: () => Buffer.from('x') }) };
  } }
};
const ctx = vm.createContext(sandbox);
vm.runInContext(src, ctx);
const { receiveFile, HEADERS, record, orderFolder, splitName, orderStatus, customerCode,
  slug, adminLogin, payLink, payParagraph, localeFor, localesFrom, quoteSerial, folderSafe,
  fileSafe, QUOTE_SERIAL_START, oneDriveTarget, graphPath, mail, mailLang, mailLangs,
  uploadTicket, uploadDone, decide, decisionToken, quoteBody, confirmationPdf, quotePdf,
  orders, doGet, MAILS, QUOTE_VALID_DAYS, MAX_FILES, MAX_FILE_BYTES,
  graphToken, graphCredentials, graphConsentLink, graphConsentLanding,
  graphRedirectUri, GRAPH_SCOPES, doPost, noteTrouble, rowIndexFor } = ctx;

const plain = (x) => JSON.parse(JSON.stringify(x));
let checks = 0;
const ok = (cond, what) => { assert.ok(cond, what); checks++; };
const eq = (got, want, what) => { assert.strictEqual(got, want, what + ': atteso ' +
  want + ', trovato ' + got); checks++; };

/* Graph and OneDrive configured, as they will be in production. */
const configure = () => {
  props.GRAPH_TENANT = 'ten'; props.GRAPH_CLIENT = 'cli'; props.GRAPH_SECRET = 'sec';
  props.ONEDRIVE_TARGET = 'user:g.gambuto@dema-solutions.com';
  /* Where the fallback writes when Graph refuses. */
  props.ORDERS_FOLDER_ID = 'ORDERS-ROOT';
};
configure();
heading = HEADERS.slice();

// ── the sheet keeps everything ──────────────────────────────────────────────
const order = {
  id: 'QD-0042', solution: 'Professional', source: 'en-GB', targets: ['fr-FR', 'de-DE'],
  words: 5000, subject: 'mechanical', layout: { pages: 20 }, eta: '2026-11-05T17:00:00.000Z',
  currency: 'EUR', vatTreatment: 'Italy, 22%', language: 'it',
  project: 'Manuale X', customer: { firstName: 'Mara', lastName: 'Bianchi',
    email: 'm@example.it', reference: 'PO-1', note: 'glossario allegato' },
  billing: { company: true, companyName: 'Esempio Srl', vatNumber: 'IT01234567890',
    telephone: '+39 039 000', address1: 'Via Gramsci 10', address2: 'Scala B',
    city: 'Monza', postalCode: '20900', countryName: 'Italy' }
};
record(order, 600, 132, 732, false, 'https://onedrive/x', '2127/2026',
  { language: 'it', token: 'a'.repeat(32), item: 'ITEM-X' });
const row = sheetRows[0], at = (name) => row[HEADERS.indexOf(name)];
for (const [col, want] of [['Reference', 'QD-0042'], ['Status', 'Quoted'],
  ['First name', 'Mara'], ['Last name', 'Bianchi'], ['Email', 'm@example.it'],
  ['Telephone', '+39 039 000'], ['Company', 'Esempio Srl'], ['VAT number', 'IT01234567890'],
  ['Address', 'Via Gramsci 10, Scala B'], ['City', 'Monza'], ['Postal code', '20900'],
  ['Country', 'Italy'], ['Project', 'Manuale X'], ['Solution', 'Professional'],
  ['From', 'en-GB'], ['Into', 'fr-FR, de-DE'], ['Words', 10000], ['Net', 600], ['VAT', 132],
  ['Total', 732], ['Currency', 'EUR'], ['Customer reference', 'PO-1'],
  ['Notes', 'glossario allegato'], ['Files', 'https://onedrive/x'],
  ['Quote number', '2127/2026'], ['Language', 'it'], ['Drive item', 'ITEM-X']]) {
  eq(at(col), want, col);
}
eq(row.length, HEADERS.length, 'la riga ha una cella per colonna');

/*
 * Lo stato di partenza non e piu "Received": la pagina chiede un preventivo,
 * non registra un ordine. Se questo torna "Received" vuol dire che qualcuno ha
 * rimesso il vecchio flusso e il cliente riceve una conferma per qualcosa che
 * non ha ancora accettato.
 */
eq(at('Status'), 'Quoted', 'una richiesta nuova nasce Quoted');

// il token della decisione sta nel foglio, ma non esce verso il back office
eq(at('Decision token'), 'a'.repeat(32), 'il token e scritto nella riga');
const book = orders();
eq(book[0][HEADERS.indexOf('Decision token')], '(set)',
  'il back office vede che un token c e, non quale sia');
ok(book[0].join('|').indexOf('a'.repeat(32)) === -1,
  'e il token non compare da nessun altra parte nella riga servita');

// ── la coppia linguistica, come la scrive un CAT ────────────────────────────
eq(at('From code'), 'en-GB', 'il codice sorgente finisce in colonna');
eq(at('Into codes'), 'fr-FR, de-DE', 'e i codici target in lista');

const last = (name) => sheetRows[sheetRows.length - 1][HEADERS.indexOf(name)];
record(Object.assign({}, order, { source: 'English (United Kingdom)',
  targets: ['French (France)', 'Chinese (Simplified)'] }), 600, 132, 732, false, '');
eq(last('From'), 'English (United Kingdom)', 'il nome resta leggibile');
eq(last('From code'), 'en-GB', 'e il codice si ricava dal nome');
eq(last('Into codes'), 'fr-FR, zh-CN', 'anche per una lingua non europea');

record(Object.assign({}, order, { source: 'English (United Kingdom)',
  targets: ['French (France)'], sourceCode: 'en-US', targetCodes: ['pt-BR'] }),
  600, 132, 732, false, '');
eq(last('From code'), 'en-US', 'il codice mandato dalla pagina vince sul nome');
eq(last('Into codes'), 'pt-BR', 'e cosi per i target');
sheetRows.length = 1;

eq(localeFor('Klingon (Qo noS)'), '', 'lingua ignota: nessun codice');
eq(localeFor(''), '', 'cella vuota: nessun codice');
eq(localeFor('  de-DE  '), 'de-DE', 'gli spazi attorno non contano');
assert.deepStrictEqual(plain(localesFrom('fr-FR, de-DE ,')), ['fr-FR', 'de-DE'],
  'una lista si legge anche se finisce con una virgola');
checks++;

/*
 * La tabella dei codici e una copia: l originale e LANGUAGES in src/config.js,
 * nell altro repository. Se le due si separano, un progetto nasce con la
 * lingua sbagliata e il foglio non lo mostra.
 */
const configPath = path.join(__dirname, '..', '..', 'ANTIGRAVITY IDE',
  'translation-quote-app', 'src', 'config.js');
if (fs.existsSync(configPath)) {
  const want = {};
  require(configPath).LANGUAGES.forEach((l) => { want[l.name] = l.code; });
  assert.deepStrictEqual(plain(ctx.LOCALES), want,
    'LOCALES non combacia con LANGUAGES in src/config.js');
  checks++;
} else {
  console.log('  (parita con src/config.js non controllata: preventivatore non trovato)');
}

const names = (o) => plain(splitName(o));
assert.deepStrictEqual(names({ customer: { name: 'Hans Peter Richter' } }), ['Hans', 'Peter Richter']);
assert.deepStrictEqual(names({ customer: { name: 'Prince' } }), ['Prince', '']);
checks += 2;

// ── dove vanno i file ───────────────────────────────────────────────────────
props.ONEDRIVE_TARGET = '';
eq(oneDriveTarget(), null, 'senza ONEDRIVE_TARGET non si sa dove scrivere');
props.ONEDRIVE_TARGET = 'user:g.gambuto@dema-solutions.com';
ok(oneDriveTarget().base.indexOf('/users/g.gambuto%40dema-solutions.com/drive') !== -1,
  'un OneDrive personale: ' + oneDriveTarget().base);
props.ONEDRIVE_TARGET = 'site:dema.sharepoint.com,site-1,drive-9';
ok(oneDriveTarget().base.indexOf('/drives/drive-9') !== -1,
  'una libreria SharePoint: si usa l ultimo campo, il drive id');
props.ONEDRIVE_TARGET = 'user:g.gambuto@dema-solutions.com';

/*
 * Il percorso si scappa una volta sola. Scappandolo due volte uno spazio
 * diventa %2520 e Graph risponde 404 per una cartella che e li: un errore che
 * si vede solo in produzione, perche in locale i nomi non hanno spazi.
 */
eq(graphPath(['Quote Desk', 'Esempio Srl']), 'Quote%20Desk/Esempio%20Srl',
  'uno spazio diventa %20');
ok(graphPath(['Quote Desk']).indexOf('%2520') === -1, 'e non %2520');

// ── la cartella porta il cliente, il lavoro, il giorno e i due codici ───────
onedrive = {}; props = {}; configure(); sheetRows = []; heading = HEADERS.slice();
record(order, 600, 132, 732, false, '', '2127/2026', { language: 'it', token: 'b'.repeat(32) });
sheetRows[0][0] = new Date();

eq(customerCode(order), 'esempio_srl', 'azienda: ragione sociale slugged');
eq(customerCode({ customer: { firstName: 'Mara', lastName: 'Bianchi', email: 'm@x.it' }, billing: {} }),
  'mara_bianchi', 'privato: nome_cognome');
eq(customerCode({ customer: { firstName: 'Ana', lastName: 'Lundqvist', email: 'a@x.se' }, billing: {} }),
  'ana_lundqvist', 'gli accenti sono piegati, non due clienti diversi');

const made = orderFolder('QD-0042', 'Mara Bianchi', 'mara_bianchi', 'Manuale CNC',
  '2127/2026', 'Esempio Srl');
const folderName = made.getName();
ok(folderName.indexOf('QUOTE 2127-2026') === 0,
  'il numero apre il nome, con il trattino al posto della barra: ' + folderName);
ok(folderName.indexOf('Esempio Srl') !== -1, 'poi l azienda: ' + folderName);
ok(folderName.indexOf('Mara Bianchi') !== -1, 'poi la persona: ' + folderName);
ok(folderName.indexOf('Manuale CNC') !== -1, 'poi il progetto: ' + folderName);
ok(/\d{4}-\d{2}-\d{2}_\d{2}-\d{2}_QD-0042$/.test(folderName),
  'e in coda giorno, ora e riferimento: ' + folderName);
ok(Object.keys(onedrive).indexOf('Quote Desk') !== -1, 'la radice /Quote Desk esiste');
ok(Object.keys(onedrive).indexOf('Quote Desk/Esempio Srl') !== -1,
  'e sotto c e la cartella del cliente');
assert.deepStrictEqual(plain(made.getPath()),
  ['Quote Desk', 'Esempio Srl', folderName], 'il percorso torna a nomi, non scappato');
checks++;

eq(folderSafe('2127/2026'), '2127-2026', 'la barra diventa trattino');
eq(folderSafe('a:b*c?d"e<f>g|h'), 'a-b-c-d-e-f-g-h',
  'e con lei ogni carattere che Windows rifiuta');

/*
 * OneDrive ha i percorsi, Drive non li aveva: un nome file ora puo contenere
 * un pezzo di percorso, e Windows rifiuta certi nomi qualunque sia l
 * estensione. Un file che si chiama CON.docx non si scrive, e il cliente
 * vedrebbe solo un upload che non arriva.
 */
eq(fileSafe('../../../evil.sh'), 'evil.sh', 'un nome che prova a scappare viene archiviato');
eq(fileSafe('CON.docx'), '_CON.docx', 'un nome riservato da Windows viene scansato');
eq(fileSafe('nul'), '_nul', 'anche senza estensione, e anche minuscolo');
eq(fileSafe('manuale .'), 'manuale', 'punti e spazi in coda non sopravvivono');
eq(fileSafe('Manuale CNC (rev 2).docx'), 'Manuale CNC (rev 2).docx',
  'ma un nome normale resta identico');

// senza numero il preventivo prende comunque una cartella
const noQuote = orderFolder('QD-0099', 'Mara Bianchi', 'mara_bianchi', 'Senza numero');
ok(noQuote.getName().indexOf('QD-0099') === 0,
  'senza numero vale il riferimento: ' + noQuote.getName());

// il contatore parte da QUOTE_SERIAL_START e poi sale di uno
props = {}; configure();
const first = quoteSerial();
ok(first.indexOf(QUOTE_SERIAL_START + '/') === 0,
  'il primo numero e ' + QUOTE_SERIAL_START + ', trovato ' + first);
eq(quoteSerial().split('/')[0], String(QUOTE_SERIAL_START + 1), 'il secondo e quello dopo');
ok(Object.keys(props).some((k) => /^QUOTE_SERIAL_\d{4}$/.test(k)),
  'il contatore e tenuto per anno');

/*
 * Una cartella fatta una volta si ritrova dall id nella riga, non dal nome:
 * un nome si puo modificare e un percorso si puo spostare, un id no. Se questo
 * controllo cade, il secondo file di un ordine finisce in una cartella nuova.
 */
sheetRows[0][HEADERS.indexOf('Drive item')] = made.getId();
const again = orderFolder('QD-0042', '', '', '', '', '');
eq(again.getName(), folderName, 'ritrovata dall id nella riga, non ricreata');
assert.deepStrictEqual(plain(again.getPath()), ['Quote Desk', 'Esempio Srl', folderName],
  'e il percorso si ricostruisce da parentReference');
checks++;

/*
 * Una riga con un link a Google Drive nella colonna Files si apre su Drive.
 *
 * Prima lanciava un errore - "before the move to OneDrive" - e aveva senso
 * finche Drive era fuori dal progetto. Ora il ripiego ci scrive davvero,
 * quindi quell'errore avrebbe rifiutato di toccare proprio le richieste che
 * gia avevano avuto un problema. L id si prende dalla colonna 'Drive item'
 * se c'e, altrimenti si legge dall URL, che e il caso degli ordini vecchi.
 */
driveFolderStub('1AbC', 'cartella vecchia');
sheetRows[0][HEADERS.indexOf('Drive item')] = '';
sheetRows[0][HEADERS.indexOf('Files')] = 'https://drive.google.com/drive/folders/1AbC';
const vecchia = orderFolder('QD-0042', '', '', '', '', '');
eq(vecchia.isDrive, true, 'una riga con link Drive si apre su Drive');
eq(vecchia.getId(), '1AbC', "e l id si legge dall URL quando la colonna e vuota");
checks += 2;
sheetRows[0][HEADERS.indexOf('Files')] = '';
sheetRows[0][HEADERS.indexOf('Drive item')] = made.getId();

// ── chiedere il permesso di caricare ───────────────────────────────────────
onedrive = {}; sheetRows = []; heading = HEADERS.slice();
record(order, 600, 132, 732, false, '', '2127/2026', { language: 'it', token: 'c'.repeat(32) });
sheetRows[0][0] = new Date();

const ticket = uploadTicket({ id: 'QD-0042', email: 'M@Example.IT',
  name: 'manuale.docx', size: 40 * 1024 * 1024, mime: 'application/x' });
eq(ticket.ok, true, 'un file di una richiesta valida ottiene una sessione: ' + ticket.error);
ok(/^https:\/\/dema-my\.sharepoint\.com/.test(ticket.uploadUrl),
  'e l URL e quello di SharePoint, non di Graph: ' + ticket.uploadUrl);
eq(ticket.chunk, 10 * 1024 * 1024, 'il blocco e 10 MiB, multiplo dei 320 KiB che Graph vuole');
ok(ticket.chunk % (320 * 1024) === 0, 'e multiplo esatto: Graph rifiuta il resto');
ok(JSON.stringify(ticket).indexOf('graph-tok') === -1,
  'il token dell app non esce mai verso il browser');

/*
 * Le cinque condizioni. Il riferimento ORD-YYYYMMDD-XXXX e indovinabile -
 * quattro caratteri su una data nota - e il premio per chi indovina sarebbe il
 * permesso di scrivere in una cartella. Serve anche l email, e un rifiuto dice
 * sempre la stessa cosa.
 */
eq(uploadTicket({ id: 'QD-0042', email: 'ladro@x.com', name: 'x', size: 10 }).error,
  'not found', 'email sbagliata: respinto');
eq(uploadTicket({ id: 'QD-9999', email: 'm@example.it', name: 'x', size: 10 }).error,
  'not found', 'richiesta inesistente: stessa risposta');
eq(uploadTicket({ id: 'QD-0042', name: 'x', size: 10 }).error,
  'not found', 'senza email: respinto');
eq(uploadTicket({ id: '../../etc', email: 'm@example.it', name: 'x', size: 10 }).error,
  'bad reference', 'un riferimento che prova a scappare non arriva nemmeno al foglio');
eq(uploadTicket({ id: 'QD-0042', email: 'm@example.it', name: 'x', size: 0 }).error,
  'empty file', 'un file vuoto non ha bisogno di una sessione');
eq(uploadTicket({ id: 'QD-0042', email: 'm@example.it', name: 'x',
  size: MAX_FILE_BYTES + 1 }).error, 'file too large',
  'un tetto c e, perche senza tetto non c e niente da superare');

/*
 * Due giga devono passare: e la dimensione per cui questo pezzo esiste, il
 * posto di WeTransfer. Se questo controllo diventa rosso, qualcuno ha rimesso
 * un tetto da trasporto su un canale che non trasporta.
 */
eq(MAX_FILE_BYTES, 2 * 1024 * 1024 * 1024, 'il tetto e 2 GB');
const big = uploadTicket({ id: 'QD-0042', email: 'm@example.it',
  name: 'lavoro completo.zip', size: 2 * 1024 * 1024 * 1024 });
eq(big.ok, true, 'uno zip da 2 GB ottiene la sua sessione: ' + big.error);
eq(big.name, 'lavoro completo.zip', 'e lo zip non e un formato da rifiutare');

/*
 * L'endpoint e aperto per scelta: nessun account, nessuna password. Quel che
 * un estraneo non deve poter lasciare sul disco del desk e qualcosa che si
 * esegue.
 */
for (const bad of ['fattura.exe', 'setup.msi', 'script.bat', 'x.ps1', 'a.vbs', 'run.sh']) {
  eq(uploadTicket({ id: 'QD-0042', email: 'm@example.it', name: bad, size: 10 }).error,
    'file type not accepted', bad + ' respinto');
}
for (const good of ['job.zip', 'manuale.pdf', 'testo.docx', 'memoria.tmx', 'strano.xyz']) {
  eq(uploadTicket({ id: 'QD-0042', email: 'm@example.it', name: good, size: 10 }).ok, true,
    good + ' accettato');
}

/*
 * La sessione non vive piu della finestra in cui la richiesta accetta file.
 * Un'ora non bastava: 2 GB su una linea da 3 Mbit in salita sono un'ora e
 * mezza, e il cliente avrebbe visto fallire al novanta per cento.
 */
const askedFor = JSON.parse(fetchCalls.filter(
  (c) => /createUploadSession/.test(c.url)).pop().opts.payload);
const life = (new Date(askedFor.expirationDateTime).getTime() - Date.now()) / 3600000;
ok(life > 5.5 && life <= 6.5,
  'si chiede una sessione lunga quanto la finestra, sei ore: ' + life.toFixed(1));
ok(new Date(big.expires).getTime() > Date.now(),
  'e quel che si riporta al browser e la scadenza che Graph ha concesso, non quella chiesta');

// trenta file e non uno di piu
const folderNow = orderFolder('QD-0042', '', '', '', '', '');
for (let i = 0; i < MAX_FILES; i++) {
  onedrive[folderNow.getPath().join('/') + '/f' + i + '.docx'] =
    itemFor(folderNow.getPath().join('/') + '/f' + i + '.docx', false, 10);
}
eq(uploadTicket({ id: 'QD-0042', email: 'm@example.it', name: 'x', size: 10 }).error,
  'too many files', 'al trentunesimo si chiude');
Object.keys(onedrive).forEach((p) => { if (/\/f\d+\.docx$/.test(p)) delete onedrive[p]; });

// una richiesta di ieri non prende piu file
sheetRows[0][0] = new Date(Date.now() - 7 * 3600 * 1000);
eq(uploadTicket({ id: 'QD-0042', email: 'm@example.it', name: 'x', size: 10 }).error,
  'order closed to uploads', 'fuori dalle sei ore: chiuso');
sheetRows[0][0] = new Date();

// la conferma dell upload scrive dove e la cartella
sheetRows[0][HEADERS.indexOf('Files')] = '';
sheetRows[0][HEADERS.indexOf('Drive item')] = '';
const done = uploadDone({ id: 'QD-0042', email: 'm@example.it' });
eq(done.ok, true, 'la conferma passa: ' + done.error);
ok(String(sheetRows[0][HEADERS.indexOf('Files')]).indexOf('sharepoint.com') !== -1,
  'e la riga impara il link OneDrive: ' + sheetRows[0][HEADERS.indexOf('Files')]);
ok(String(sheetRows[0][HEADERS.indexOf('Drive item')]).indexOf('ITEM') === 0,
  'e l id della cartella');

/*
 * La mail che il desk aspetta: parte quando i file sono arrivati davvero, non
 * quando il preventivo e uscito. Un archivio da 2 GB sta ancora salendo un ora
 * dopo che il cliente ha chiuso la scheda, e il desk deve sapere quando c e.
 */
const arrival = fetchCalls.filter((c) => /sendMail/.test(c.url)).pop();
ok(arrival, 'una mail parte all arrivo dei file');
const told = JSON.parse(arrival.opts.payload).message;
eq(told.toRecipients[0].emailAddress.address, 'g.gambuto@dema-solutions.com',
  'e va al desk');
ok(/Files uploaded/.test(told.subject), 'e si capisce dall oggetto: ' + told.subject);
ok(told.subject.indexOf('Esempio Srl') !== -1, 'con il nome del cliente nell oggetto');
ok(told.body.content.indexOf('sharepoint.com') !== -1,
  'il corpo porta il link alla cartella su OneDrive');
ok(told.body.content.indexOf('vecchio.docx') !== -1 ||
   told.body.content.indexOf('Open the folder') !== -1,
  'e dice cosa c e dentro');
ok(!told.attachments, 'nessun allegato: si manda il link, non i due giga');

/*
 * La vecchia tratta base64 resta finche la pagina pubblicata la usa, ma scrive
 * su OneDrive come tutto il resto: una sola destinazione, non due.
 */
const b64 = Buffer.from('ciao').toString('base64');
const old = receiveFile({ id: 'QD-0042', name: 'vecchio.docx', data: b64, mime: 'application/x' });
eq(old.ok, true, 'la tratta vecchia accetta ancora: ' + old.error);
ok(Object.keys(onedrive).some((p) => /vecchio\.docx$/.test(p)),
  'e il file finisce su OneDrive, non su Drive');
eq(receiveFile({ id: 'QD-0042', name: 'x', data: '' }).error, 'empty file', 'file vuoto');
eq(receiveFile({ id: 'QD-9999', name: 'x', data: b64 }).error, 'unknown order', 'ordine ignoto');
eq(receiveFile({ id: 'QD-0042', name: 'x',
  data: 'A'.repeat(40 * 1024 * 1024) }).error, 'file too large',
  'oltre quel che una richiesta Apps Script regge');

// ── le lingue in cui si scrive al cliente ──────────────────────────────────
const langs = mailLangs();
ok(langs.indexOf('en') !== -1 && langs.indexOf('it') !== -1,
  'inglese e italiano ci sono: ' + langs.join(' '));
eq(mailLang('it'), 'it', 'la lingua chiesta si rispetta');
eq(mailLang('IT'), 'it', 'maiuscole o minuscole non contano');
eq(mailLang('it-IT'), 'it', 'e nemmeno la variante regionale');
eq(mailLang('fi'), 'en', 'una lingua senza dizionario ricade sull inglese');
eq(mailLang(''), 'en', 'e senza indicazione, inglese');
ok(mail('quote_title', 'it').length > 0, 'una chiave esistente ha un testo italiano');
eq(mail('quote_valid', 'it', { date: '2026-12-14' }),
  'Il preventivo è valido fino al 2026-12-14.', 'i segnaposto vengono riempiti');

/*
 * Ogni chiave deve avere almeno l inglese, altrimenti un cliente riceve il
 * nome della chiave. E ogni lingua dichiarata deve coprire ogni chiave,
 * altrimenti una mail esce meta in una lingua e meta in un altra - che e
 * peggio di una mail in una lingua sola che non e la sua.
 */
let holes = [];
Object.keys(MAILS).forEach((key) => {
  if (!MAILS[key].en) holes.push(key + ' (en)');
  langs.forEach((lang) => { if (!MAILS[key][lang]) holes.push(key + ' (' + lang + ')'); });
});
eq(holes.join(', '), '', 'nessun buco nel dizionario');

// ── il preventivo, e i due link ────────────────────────────────────────────
onedrive = {}; sheetRows = []; heading = HEADERS.slice(); sentMail = [];
const token = decisionToken();
ok(/^[0-9a-f]{32}$/.test(token), 'il token e 32 esadecimali: ' + token);
ok(decisionToken() !== decisionToken(), 'e non e due volte lo stesso');

record(order, 600, 132, 732, false, '', '2127/2026',
  { language: 'it', token: token, item: '' });
sheetRows[0][0] = new Date();

const body = quoteBody(order, 600, 132, 732, 'it', '2127/2026', token);
ok(body.indexOf('a=yes') !== -1 && body.indexOf('a=no') !== -1,
  'la mail porta i due link, accetta e rifiuta');
ok(body.indexOf('t=' + token) !== -1, 'e ciascuno porta il segreto');
ok(body.indexOf('Accetto il preventivo') !== -1,
  'con le etichette nella lingua del cliente');
ok(body.indexOf('lang=it') !== -1, 'e la lingua viaggia nel link');

/*
 * Il controllo piu importante di tutto il file. Outlook e Gmail seguono i link
 * di un messaggio per scansionarli: se un GET accettasse, il preventivo
 * verrebbe accettato da un antivirus pochi minuti dopo l invio e il cliente non
 * lo vedrebbe mai. doGet deve solo disegnare.
 */
const page = doGet({ parameter: { a: 'yes', id: 'QD-0042', t: token, lang: 'it' } });
eq(sheetRows[0][HEADERS.indexOf('Status')], 'Quoted',
  'aprire il link di accettazione NON accetta');
ok(page.getContent().indexOf('noindex') !== -1, 'la pagina non va indicizzata');
ok(page.getContent().indexOf('Accetta il preventivo') !== -1,
  'e parla la lingua del cliente');
ok(page.getContent().indexOf('button') !== -1, 'e il bottone c e');
ok(page.getContent().indexOf(token) !== -1, 'il bottone sa quale preventivo conferma');

// un link con il token sbagliato non dice nemmeno se il preventivo esiste
const wrong = doGet({ parameter: { a: 'yes', id: 'QD-0042', t: 'f'.repeat(32), lang: 'it' } });
ok(wrong.getContent().indexOf('non si trova') !== -1, 'token sbagliato: non si trova');
eq(decide({ id: 'QD-0042', token: 'f'.repeat(32), decision: 'accept' }).error, 'not found',
  'e la decisione viene respinta');
eq(decide({ id: 'QD-0042', decision: 'accept' }).error, 'not found', 'senza token: respinto');

// ── la decisione, scritta una volta sola ───────────────────────────────────
sentMail = []; fetchCalls = [];
const said = decide({ id: 'QD-0042', token: token, decision: 'accept' });
eq(said.ok, true, 'con il token giusto si accetta: ' + said.error);
eq(sheetRows[0][HEADERS.indexOf('Status')], 'Accepted', 'la riga diventa Accepted');
/* instanceof non serve: una Date nata dentro il sandbox ha il suo
   costruttore, e per il realm di fuori non e una Date. */
const isDate = (x) => Object.prototype.toString.call(x) === '[object Date]';
ok(isDate(sheetRows[0][HEADERS.indexOf('Decided')]), 'con la data della decisione');

const toCustomer = fetchCalls.filter((c) => /sendMail/.test(c.url));
ok(toCustomer.length >= 2, 'partono due mail: al cliente e al desk');
const payload = JSON.parse(toCustomer[0].opts.payload);
eq(payload.message.toRecipients[0].emailAddress.address, 'm@example.it',
  'la prima va al cliente');
ok(/Ordine confermato/.test(payload.message.subject),
  'con oggetto in italiano: ' + payload.message.subject);
ok((payload.message.attachments || []).length === 1, 'e il PDF allegato');
ok(/^order-confirmation-/.test(payload.message.attachments[0].name),
  'che e la conferma d ordine, non il preventivo: ' + payload.message.attachments[0].name);

// un secondo clic non ordina due volte
const twice = decide({ id: 'QD-0042', token: token, decision: 'accept' });
eq(twice.error, 'already decided', 'premere due volte non accetta due volte');
eq(decide({ id: 'QD-0042', token: token, decision: 'decline' }).error, 'already decided',
  'e non si rifiuta per link quel che si e accettato');

// il rifiuto lascia tutto dove e
sheetRows = []; heading = HEADERS.slice(); onedrive = {}; fetchCalls = [];
const t2 = decisionToken();
record(order, 600, 132, 732, false, '', '2128/2026', { language: 'fr', token: t2 });
sheetRows[0][0] = new Date();
const no = decide({ id: 'QD-0042', token: t2, decision: 'decline' });
eq(no.ok, true, 'si rifiuta: ' + no.error);
eq(sheetRows[0][HEADERS.indexOf('Status')], 'Declined', 'la riga diventa Declined');
const declined = JSON.parse(fetchCalls.filter((c) => /sendMail/.test(c.url))[0].opts.payload);
ok(/Devis .* refus/.test(declined.message.subject),
  'e la mail e in francese: ' + declined.message.subject);
ok(!declined.message.attachments, 'un rifiuto non porta PDF');
ok(Object.keys(onedrive).length === 0 ||
  !Object.keys(onedrive).some((p) => /order-confirmation/.test(p)),
  'e non si deposita nessuna conferma');

// un preventivo scaduto non si accetta piu
sheetRows = []; heading = HEADERS.slice();
const t3 = decisionToken();
record(order, 600, 132, 732, false, '', '2129/2026', { language: 'en', token: t3 });
sheetRows[0][0] = new Date(Date.now() - (QUOTE_VALID_DAYS + 1) * 24 * 3600 * 1000);
const late = decide({ id: 'QD-0042', token: t3, decision: 'accept' });
eq(late.error, 'expired', 'oltre i ' + QUOTE_VALID_DAYS + ' giorni: scaduto');
eq(sheetRows[0][HEADERS.indexOf('Status')], 'Expired', 'e la riga lo registra');

// ── quel che un cliente puo vedere ─────────────────────────────────────────
sheetRows = []; heading = HEADERS.slice();
record(order, 600, 132, 732, false, 'https://onedrive/x', '2130/2026',
  { language: 'it', token: 'd'.repeat(32) });
sheetRows[0][0] = new Date();

const seen = orderStatus({ id: 'QD-0042', email: 'M@Example.IT' });
eq(seen.ok, true, 'riferimento ed email giusti: ' + seen.error);
eq(seen.order.reference, 'QD-0042', 'il riferimento torna');
eq(seen.order.status, 'Quoted', 'e lo stato e quello nuovo');
eq(seen.order.total, 732, 'e il totale');
const shown = Object.keys(plain(seen.order));
for (const secret of ['address', 'Address', 'telephone', 'vat', 'notes', 'city',
  'token', 'Decision token']) {
  ok(!shown.includes(secret), 'non esce dalla finestra del cliente: ' + secret);
}
assert.deepStrictEqual(plain(orderStatus({ id: 'QD-0042', email: 'ladro@x.com' })),
  { ok: false, error: 'not found' }, 'email sbagliata: respinto');
assert.deepStrictEqual(plain(orderStatus({ id: 'QD-9999', email: 'm@example.it' })),
  { ok: false, error: 'not found' }, 'ordine inesistente: stessa risposta');
checks += 2;

// ── il PDF, nelle sue due vesti ────────────────────────────────────────────
const quoteDoc = quotePdf(order, 600, 132, 732, 'it', '2127/2026');
ok(/^quote-2127-2026\.pdf$/.test(quoteDoc.getName()),
  'il preventivo si chiama col suo numero: ' + quoteDoc.getName());
const quoteHtml = String(quoteDoc.getBytes());
ok(quoteHtml.indexOf('PREVENTIVO') !== -1, 'e si intitola PREVENTIVO in italiano');
ok(quoteHtml.indexOf('non una fattura') !== -1, 'e dice di non essere una fattura');
ok(quoteHtml.indexOf('Valido fino al') !== -1, 'e porta la scadenza');
ok(quoteHtml.indexOf('Italy, 22%') !== -1,
  'la formula fiscale resta come e: e testo legale, non si traduce');

const orderDoc = confirmationPdf(order, 600, 132, 732, 'de', 'order', '2127/2026');
const orderHtml = String(orderDoc.getBytes());
ok(orderHtml.indexOf('AUFTRAGSBEST') !== -1, 'la conferma in tedesco si intitola cosi');
ok(orderHtml.indexOf('Valido fino al') === -1, 'e una conferma non ha scadenza');

// ── l accesso al back office ───────────────────────────────────────────────
props = {}; configure();
eq(adminLogin('desk', 'secret'), false, 'senza credenziali impostate: chiuso');
props.ADMIN_USER = 'desk'; props.ADMIN_PASS = 'lungo-e-segreto';
eq(adminLogin('desk', 'lungo-e-segreto'), true, 'la coppia giusta entra');
eq(adminLogin('DESK', 'lungo-e-segreto'), true, 'il nome non guarda le maiuscole');
eq(adminLogin('desk', 'Lungo-e-segreto'), false, 'la password si');
eq(adminLogin('altro', 'lungo-e-segreto'), false, 'nome sbagliato: fuori');
eq(adminLogin('desk', ''), false, 'password vuota: fuori');

// ── pagare con la carta, che resta spento finche non c e una chiave ────────
props = {}; configure(); fetchCalls = [];
eq(payLink(order, 732), '', 'senza chiave: nessun link, nessuna chiamata');
eq(fetchCalls.filter((c) => /stripe/.test(c.url)).length, 0, 'e Stripe non viene contattato');
eq(payParagraph(''), '', 'e la conferma non nomina la carta');

props.STRIPE_SECRET = 'sk_test_x'; fetchCalls = [];
const link = payLink(order, 732);
eq(link, 'https://checkout.stripe.com/c/pay/abc', 'con chiave: arriva il link');
const call = fetchCalls.filter((c) => /stripe/.test(c.url))[0];
eq(call.opts.headers.Authorization, 'Bearer sk_test_x', 'con la chiave in testa');
eq(call.opts.payload['line_items[0][price_data][unit_amount]'], '73200',
  'l importo e in centesimi, non in euro');
eq(call.opts.payload.client_reference_id, 'QD-0042', 'e porta il riferimento');
ok(payParagraph(link).indexOf(link) !== -1, 'il paragrafo porta il link');

stripeCode = 402; stripeBody = '{"error":{"message":"card payments not activated"}}';
eq(payLink(order, 732), '', 'Stripe rifiuta: nessun link, ordine salvo');
stripeCode = 200; stripeBody = '{"url":"https://checkout.stripe.com/c/pay/abc"}';
eq(payLink(order, 0), '', 'totale zero: nessun link');

// ── Trados, che non parte finche non ha le sue property ────────────────────
const { tradosCredentials, tradosToken, tradosDirections, tradosCall,
  TRADOS_AUDIENCE, TRADOS_API } = ctx;

props = {}; configure(); fetchCalls = [];
eq(tradosCredentials(), null, 'senza property Trados resta spento');
props.TRADOS_CLIENT = 'cid'; props.TRADOS_SECRET = 'shhh'; props.TRADOS_TENANT = 'ten-1';
eq(tradosCredentials().tenant, 'ten-1', 'con le tre property si accende');
delete props.TRADOS_SECRET;
eq(tradosCredentials(), null, 'due su tre non bastano: meglio spento che a meta');
props.TRADOS_SECRET = 'shhh';

/*
 * Il limite piu stretto di tutta l integrazione: Auth0 concede sedici token al
 * giorno. Un token vale 24 ore, quindi ne basta uno - ma solo se viene tenuto.
 * Se questo controllo diventa rosso l integrazione non rallenta: smette del
 * tutto a meta giornata.
 */
fetchCalls = [];
stripeBody = '{"access_token":"tok-abc","expires_in":86400}';
const creds = tradosCredentials();
eq(tradosToken(creds), 'tok-abc', 'il primo token arriva da Auth0');
const auth0 = fetchCalls.filter((c) => /auth0/.test(c.url));
eq(auth0.length, 1, 'una sola chiamata per prenderlo');
const asked = JSON.parse(auth0[0].opts.payload);
eq(asked.audience, 'https://api.sdl.com',
  'audience api.sdl.com: con api.trados.com il token esce ma l API lo rifiuta');
eq(TRADOS_AUDIENCE, asked.audience, 'e la costante dice la stessa cosa');
eq(tradosToken(creds), 'tok-abc', 'il secondo lo trova gia in casa');
eq(fetchCalls.filter((c) => /auth0/.test(c.url)).length, 1,
  'e non ne chiede un altro: sono sedici al giorno');

props.TRADOS_TOKEN_UNTIL = String(Date.now() + 30 * 60 * 1000);
stripeBody = '{"access_token":"tok-nuovo","expires_in":86400}';
eq(tradosToken(creds), 'tok-nuovo', 'vicino alla scadenza se ne prende uno nuovo');

// ── le lingue non si indovinano mai ───────────────────────────────────────
const tradosRow = (over) => {
  const all = Object.assign({ 'From': 'English (United Kingdom)',
    'Into': 'French (France)', 'From code': 'en-GB', 'Into codes': 'fr-FR, de-DE' },
    over || {});
  return (name) => (name in all ? all[name] : '');
};
const pair = tradosDirections(tradosRow());
eq(pair.length, 2, 'due target, due direzioni');
eq(pair[0].sourceLanguage.languageCode, 'en-GB', 'il sorgente');
eq(pair[1].targetLanguage.languageCode, 'de-DE', 'lo spazio dopo la virgola cade');
const older = tradosDirections(tradosRow({ 'From code': '', 'Into codes': '' }));
eq(older[0].targetLanguage.languageCode, 'fr-FR',
  'una riga scritta prima delle colonne codice si ricava dai nomi');
eq(tradosDirections(tradosRow({ 'From code': '', 'Into codes': '', 'From': 'Klingon' })), null,
  'lingua sconosciuta: nessuna direzione, nessun progetto, qualcuno guarda');
eq(tradosDirections(tradosRow({ 'Into codes': '', 'Into': '' })), null, 'nessun target: idem');

fetchCalls = []; stripeCode = 200; stripeBody = '{"id":"prj-1"}';
const tradosMade = tradosCall(creds, 'post', '/projects', { name: 'x' });
eq(tradosMade.id, 'prj-1', 'il progetto torna con il suo id');
const tradosHit = fetchCalls.filter((c) => c.url.indexOf(TRADOS_API) === 0)[0];
eq(tradosHit.opts.headers['X-LC-Tenant'], 'ten-1',
  'X-LC-Tenant con i trattini: scritto X-LC_Tenant risponde 401 e non dice perche');

stripeCode = 400; stripeBody = '{"message":"Input validation error"}';
assert.throws(() => tradosCall(creds, 'post', '/projects', { name: 'x' }),
  /HTTP 400/, 'un rifiuto diventa un errore, non un progetto fantasma');
checks++;
stripeCode = 200; stripeBody = '{"url":"https://checkout.stripe.com/c/pay/abc"}';

/*
 * -- la richiesta finisce nel foglio ANCHE se Microsoft rifiuta --------------
 *
 * E il controllo che vale piu di tutti per chi vende: la riga nel foglio e la
 * dashboard sono l unica traccia della richiesta di un cliente. Prima di
 * questo giro la cartella OneDrive veniva creata PRIMA della riga, quindi un
 * 403 di Graph - un tenant che non ha dato il consenso, un segreto scaduto -
 * faceva sparire tutto: nessuna riga, niente in back office, e un cliente che
 * aveva compilato il modulo e non esisteva per nessuno.
 */
onedrive = {}; sheetRows = []; heading = HEADERS.slice(); sentMail = []; fetchCalls = [];
graphRefuses = true;
const refused = JSON.parse(doPost({ postData: { contents: JSON.stringify(
  Object.assign({}, order, { id: 'QD-0099', action: 'quote' })) } }).body);
graphRefuses = false;

eq(refused.ok, true, 'la richiesta e accettata anche se Microsoft rifiuta');
eq(refused.files, true, 'e i file hanno dove andare lo stesso: il ripiego su Drive');
ok(/drive\.google\.com/.test(String(sheetRows[0][HEADERS.indexOf('Files')])),
  'la colonna Files punta alla cartella Drive: ' +
  sheetRows[0][HEADERS.indexOf('Files')]);
eq(sheetRows.length, 1, 'e soprattutto: la riga nel foglio ce');
eq(sheetRows[0][HEADERS.indexOf('Reference')], 'QD-0099', 'con il suo riferimento');
eq(sheetRows[0][HEADERS.indexOf('Status')], 'Quoted', 'e lo stato giusto');
eq(sheetRows[0][HEADERS.indexOf('Email')], 'm@example.it', 'e il cliente dentro');
ok(Number(sheetRows[0][HEADERS.indexOf('Total')]) >= 732,
  'e il prezzo dentro, mai sotto il listino: ' + sheetRows[0][HEADERS.indexOf('Total')]);
ok(/Google Drive/.test(String(sheetRows[0][HEADERS.indexOf('Trouble')])),
  'e Trouble dice che sono sul disco sbagliato, cosi qualcuno li sposta: ' +
  sheetRows[0][HEADERS.indexOf('Trouble')]);
ok(sentMail.length >= 2,
  'e le mail partono lo stesso, dal conto Gmail: ' + sentMail.length);

/* La dashboard legge dallo stesso posto, quindi deve vederla. */
const refusedBook = orders();
eq(refusedBook.length, 1, 'il back office vede la richiesta');
eq(refusedBook[0][HEADERS.indexOf('Reference')], 'QD-0099', 'ed e la sua');
eq(refusedBook[0][HEADERS.indexOf('Decision token')], '(set)',
  'col segreto mascherato, come sempre');

/*
 * -- la cartella si ricostruisce dal foglio ---------------------------------
 *
 * Chiamata con il solo riferimento - e lo fanno sia un upload sia un secondo
 * tentativo dopo un rifiuto - deve rifare il nome dalla riga, altrimenti la
 * stessa richiesta finisce in due cartelle con due nomi diversi.
 */
const rebuilt = orderFolder('QD-0099', '');
ok(/Esempio Srl/.test(rebuilt.getPath().join('/')),
  'la cartella sta sotto il nome dell azienda: ' + rebuilt.getPath().join('/'));
ok(/QUOTE/.test(rebuilt.getPath()[2]) || /2127/.test(rebuilt.getPath()[2]),
  'e porta il numero di preventivo nel nome: ' + rebuilt.getPath()[2]);
ok(/Manuale X/.test(rebuilt.getPath()[2]),
  'e il progetto, come le altre: ' + rebuilt.getPath()[2]);
eq(String(sheetRows[0][HEADERS.indexOf('Trouble')]).indexOf('OneDrive') !== -1, true,
  'e il guaio di prima resta scritto: non si cancella da se');

/*
 * -- quello che il desk puo fare su una richiesta gia presa -----------------
 *
 * Tre azioni, tutte dietro il login che apre il libro e mai dietro il solo
 * riferimento. Esistono perche il 14 novembre 2026 una richiesta era
 * irraggiungibile: Graph rifiutava, la cartella non c'era, e del preventivo
 * restava solo l allegato in una mail. Il PDF pero si ridisegna dalla riga,
 * quindi la tracciabilita non ha mai avuto bisogno di un archivio.
 */
const deskCall = (body) => JSON.parse(doPost({ postData: { contents: JSON.stringify(body) } }).body);

props.ADMIN_USER = 'desk@example.it';
props.ADMIN_PASS = 'segreta';
const asDesk = (extra) => Object.assign(
  { user: 'desk@example.it', pass: 'segreta' }, extra);

// senza credenziali non si passa, e la risposta non distingue i casi
eq(deskCall({ action: 'quotePdf', id: 'QD-0099' }).error, 'refused',
  'il PDF non si scarica senza login');
eq(deskCall({ action: 'resend', id: 'QD-0099', user: 'desk@example.it', pass: 'sbagliata' }).error,
  'refused', 'ne con la password sbagliata');
eq(deskCall(asDesk({ action: 'quotePdf', id: 'NON-ESISTE' })).error, 'not found',
  'e una richiesta inventata non esiste');
checks += 3;

// il preventivo si ricostruisce dalla riga, senza che nulla sia stato archiviato
const rebuiltPdf = deskCall(asDesk({ action: 'quotePdf', id: 'QD-0099' }));
eq(rebuiltPdf.ok, true, 'il PDF del preventivo si rigenera dalla riga');
eq(rebuiltPdf.mime, 'application/pdf', 'ed e un PDF');
ok(rebuiltPdf.data && rebuiltPdf.data.length > 0, 'con dei byte dentro');
ok(/2127|QD-0099/.test(String(rebuiltPdf.name)),
  'e il nome porta il numero di preventivo o il riferimento: ' + rebuiltPdf.name);
checks += 4;

// rimandare la mail: stesso documento, stessi due link, stessa lingua
sentMail = []; fetchCalls = [];
const again2 = deskCall(asDesk({ action: 'resend', id: 'QD-0099' }));
eq(again2.ok, true, 'la mail di preventivo si rimanda');
eq(again2.sent, 'm@example.it', 'al cliente della riga, non a chi lo chiede');
checks += 2;

// una decisione presa non si rimanda: si parla col desk, non si ripreme un bottone
sheetRows[0][HEADERS.indexOf('Status')] = 'Accepted';
eq(deskCall(asDesk({ action: 'resend', id: 'QD-0099' })).error, 'already decided',
  'un preventivo gia deciso non si rimanda');
checks++;
sheetRows[0][HEADERS.indexOf('Status')] = 'Quoted';

/*
 * Riprovare la cartella. Con Graph ancora in rifiuto il ripiego la fa su
 * Drive e lo dice; la colonna Trouble non si svuota, perche una cartella sul
 * disco sbagliato e meta recupero e sparire sarebbe peggio che restare.
 */
graphRefuses = true;
const retried = deskCall(asDesk({ action: 'retryFiles', id: 'QD-0099' }));
graphRefuses = false;
eq(retried.ok, true, 'il secondo tentativo risponde');
eq(retried.drive, true, 'e dice che e finito su Drive, non su OneDrive');
ok(/drive\.google\.com/.test(String(retried.url)), 'con il suo link: ' + retried.url);
ok(String(sheetRows[0][HEADERS.indexOf('Trouble')]).length > 0,
  'e il guaio resta scritto finche i file non sono dove devono stare');
checks += 4;

// con Graph che risponde, la cartella si fa davvero e la riga si ripulisce
sheetRows[0][HEADERS.indexOf('Files')] = '';
sheetRows[0][HEADERS.indexOf('Drive item')] = '';
const recovered = deskCall(asDesk({ action: 'retryFiles', id: 'QD-0099' }));
eq(recovered.ok, true, 'con Graph che risponde il recupero riesce');
eq(recovered.drive, false, 'e la cartella e su OneDrive');
eq(recovered.filed, true, 'col preventivo archiviato accanto');
eq(String(sheetRows[0][HEADERS.indexOf('Trouble')]), '',
  'e solo allora la riga smette di chiedere aiuto');
checks += 4;

// -- which grant Microsoft is asked for ------------------------------------
const tokenCalls = () => fetchCalls.filter(
  (c) => c.url.indexOf('login.microsoftonline.com') !== -1);

delete props.GRAPH_REFRESH;
fetchCalls = [];
graphToken(graphCredentials());
eq(tokenCalls()[0].opts.payload.grant_type, 'client_credentials',
  'senza GRAPH_REFRESH il token si chiede come applicazione');

props.GRAPH_REFRESH = 'refresh-1';
fetchCalls = [];
graphToken(graphCredentials());
const asPerson = tokenCalls()[0].opts.payload;
eq(asPerson.grant_type, 'refresh_token',
  'con GRAPH_REFRESH il token si chiede come persona');
eq(asPerson.refresh_token, 'refresh-1', 'e si manda il refresh conservato');
ok(GRAPH_SCOPES.indexOf('offline_access') !== -1,
  'offline_access tra gli scope: senza, Microsoft non rende un refresh token');

/* Microsoft rende spesso un refresh nuovo: va riscritto, o il giorno in cui il
   vecchio smette di valere il desk resta muto senza sapere perche. */
tokenEndpointGives = 'refresh-2';
graphToken(graphCredentials());
eq(props.GRAPH_REFRESH, 'refresh-2', 'il refresh ruotato viene riscritto');
tokenEndpointGives = 'refresh-1';

// -- il consenso: lo state non e' decorativo --------------------------------
props.GRAPH_REFRESH = 'refresh-2';
const consentLink = graphConsentLink();
ok(consentLink.indexOf('response_type=code') !== -1, 'il link chiede un code');
ok(consentLink.indexOf(encodeURIComponent(graphRedirectUri())) !== -1,
  'e porta il redirect URI che va registrato in Entra');
ok(!!props.GRAPH_STATE, 'e lascia uno state da riconoscere');

const consentRefused = graphConsentLanding('any-code', 'non-lo-state');
ok(consentRefused.getContent().indexOf('Not expected') !== -1,
  'uno state sbagliato non scambia niente');
eq(props.GRAPH_REFRESH, 'refresh-2',
  'e soprattutto non sovrascrive la connessione buona');

tokenEndpointGives = 'refresh-3';
const consentDone = graphConsentLanding('any-code', props.GRAPH_STATE);
ok(consentDone.getContent().indexOf('Connected') !== -1, 'lo state giusto connette');
eq(props.GRAPH_REFRESH, 'refresh-3', 'e conserva il refresh token nuovo');
eq(props.GRAPH_STATE, undefined, 'lo state vale una volta sola');
tokenEndpointGives = 'refresh-1';
delete props.GRAPH_REFRESH;

console.log(checks + ' controlli verdi: il foglio registra tutto, i file vanno su OneDrive, ' +
  'aprire un link non accetta nulla, una decisione si prende una volta sola, le mail ' +
  'parlano la lingua del cliente e i casi storti sono respinti');

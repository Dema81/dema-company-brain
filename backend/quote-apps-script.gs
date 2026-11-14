/**
 * Quote Desk — the quoting desk behind ai-dema-solutions.com/quote
 *
 * The quote page is a static file on GitHub Pages, so it has no server. This web
 * app is its back office: it takes the request, checks the money against its own
 * copy of the rate card, writes a row in the sheet, files the customer's
 * documents on OneDrive, and sends two emails — the quote with its PDF to the
 * customer, and the whole record to the desk.
 *
 * ── A quote, not an order ───────────────────────────────────────────────────
 *  Nothing is agreed when the page is submitted. The customer receives a quote
 *  carrying two links, accept and decline, and the job begins only when one of
 *  them is confirmed. Until then the row says 'Quoted' and the files sit where
 *  they were uploaded. A quote expires after QUOTE_VALID_DAYS: a price from
 *  last spring must not still be claimable next year.
 *
 *  The links do not act when they are fetched. Outlook and Gmail follow links
 *  in a message to scan them, so a link that accepted on sight would be
 *  accepted by a virus scanner rather than by a person. doGet shows a page;
 *  only its button, which posts, changes anything. Do not "simplify" that.
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
 * number: a request arrived with no documents to translate, which is most of a
 * translation job missing. The files now follow the request into a folder of
 * their own, one per request, and the paperwork is filed beside them so there
 * is a record of what was sent and to whom.
 *
 * That folder is on OneDrive, not on Google Drive. The desk works in Microsoft
 * 365; files parked in a Google account belonging to the script were a second
 * place to look and a second thing to keep in step. Google Drive is no longer
 * written to at all — the folders of orders taken before this change stay where
 * they are, and their rows keep their old link.
 *
 * ONEDRIVE_TARGET, a script property, says which drive:
 *
 *     user:g.gambuto@dema-solutions.com          a person's own OneDrive
 *     site:<hostname>,<siteId>,<driveId>         a SharePoint document library
 *
 * It is a property and not a constant on purpose. Reaching a personal OneDrive
 * needs Files.ReadWrite.All, which is read and write over every drive in the
 * tenant; a dedicated library can be granted with Sites.Selected instead, which
 * is that library and nothing else. Moving is then one line of configuration.
 */
var ONEDRIVE_ROOT = 'Quote Desk';

/*
 * The most a file may be when the bytes have to pass through Apps Script
 * rather than go straight to OneDrive. A web app request will not carry much
 * beyond this, and base64 costs a third on top of the file itself, so a
 * larger file is refused with a reason rather than failing halfway.
 */
var RELAY_MAX_BYTES = 25 * 1024 * 1024;

/*
 * How big one file may be: 2 GB, which is the WeTransfer-sized archive this
 * desk exists to receive. It is a decision, not a limit - OneDrive itself
 * takes 250 GB per file, and nothing between here and there cares, because the
 * bytes go from the browser to OneDrive and never through this script.
 *
 * Thirty files at this size is sixty gigabytes, so the drive's own quota is the
 * thing to watch, not this number.
 */
var MAX_FILE_BYTES = 2 * 1024 * 1024 * 1024;
var MAX_FILES = 30;

/*
 * 10 MiB a block, a multiple of the 320 KiB Graph insists on (32 x 327680).
 * At 5 MiB a 2 GB archive is four hundred requests; at 10 MiB it is two
 * hundred, and a block is still small enough that losing one costs little.
 */
var UPLOAD_CHUNK = 10 * 1024 * 1024;

/*
 * How long a request keeps taking files. It is also how long an upload URL is
 * good for: no permission to write outlives the window it was granted in.
 *
 * Six hours is not generous, it is necessary. A 2 GB archive on a 10 Mbit line
 * is half an hour of uploading, and on the 3 Mbit upstream a small office
 * actually has, an hour and a half. The hour this used to be would have expired
 * mid-file, and the customer would have watched it fail at ninety per cent.
 */
var UPLOAD_WINDOW_HOURS = 6;

/*
 * How long a quote stands. A price is a snapshot - the page itself only
 * re-prices when the customer asks it to - so the quote says when it lapses,
 * and the accept link stops working on that day rather than never.
 */
var QUOTE_VALID_DAYS = 30;

/**
 * The quote number, as the trade writes it: 2127/2026.
 *
 * It sits beside the reference, it does not replace it. ORD-20261102-JTJA is
 * what the page, the uploads and the status check pass about; the quote number
 * is for people - it names the folder, heads the confirmation and fills its own
 * column in the sheet.
 *
 * The count is kept per year under QUOTE_SERIAL_<year> in the script
 * properties, so it starts again at 1 each January. The first number of all is
 * QUOTE_SERIAL_START, 2127, carrying on from the numbering already in use.
 *
 * Two orders placed in the same second would otherwise be given the same
 * number: read, add one, write is not atomic on its own. The lock makes it so.
 * If the lock cannot be had in ten seconds the order is still taken - a missing
 * quote number is a blemish, a refused order is lost work - and the folder is
 * named after the reference instead.
 */
var QUOTE_SERIAL_START = 2127;

function quoteSerial() {
  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  var key = 'QUOTE_SERIAL_' + year;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return '';
  }
  try {
    var props = PropertiesService.getScriptProperties();
    var last = Number(props.getProperty(key) || 0);
    var next = last ? last + 1 : QUOTE_SERIAL_START;
    props.setProperty(key, String(next));
    return next + '/' + year;
  } finally {
    lock.releaseLock();
  }
}

/**
 * A name Drive and the desktop client both show the same way.
 *
 * Drive itself takes a slash in a name, but Drive for desktop rewrites it when
 * it syncs, because Windows and macOS read it as a path separator - so the
 * folder would be called one thing in the browser and another on disk. The
 * slash becomes a hyphen here, once, rather than being discovered later.
 */
function folderSafe(text) {
  return String(text || '')
    .replace(/[\/\\:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
}

/**
 * A file name OneDrive will take, and a person will still recognise.
 *
 * Drive had no paths, so a stray slash was only untidy. OneDrive has paths, and
 * Windows has names it refuses outright: CON, PRN, AUX, NUL, COM1..9, LPT1..9,
 * whatever the extension. It also will not keep a trailing dot or space. What
 * is left of the name after all that is still the customer's own word for the
 * file, which is the point: the desk has to recognise it.
 */
var RESERVED_NAMES = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)/i;

/*
 * The only file types this desk turns away.
 *
 * Everything else is accepted, whatever it is and however big: a ZIP of the
 * whole job is the ordinary case, not the exception. These are not documents
 * under any reading, and the upload endpoint is open by design - no account, no
 * password - so a stranger must not be able to leave something runnable on the
 * desk's own drive. Microsoft 365 scans what lands; this is the door, not the
 * scanner.
 */
var RUNNABLE = /\.(exe|msi|scr|bat|cmd|com|pif|cpl|dll|vbs|vbe|jse?|wsf|wsh|ps1|lnk|reg|hta|jar|apk|app|deb|rpm|sh)$/i;

function fileSafe(text) {
  var name = String(text || 'file')
    .replace(/[^A-Za-z0-9 ._()À-ɏ-]/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^[._ ]+/, '')
    .replace(/[. ]+$/, '')
    .slice(0, 150) || 'file';
  if (RESERVED_NAMES.test(name)) name = '_' + name;
  return name;
}

/* ───────────────────────────────────────────── OneDrive, through Graph */

/**
 * Which drive the files go to, from the ONEDRIVE_TARGET script property.
 *
 * Returns the Graph prefix every call hangs off, or null when the property is
 * missing - and then nothing is uploaded and selfTest says so, rather than
 * files going quietly nowhere.
 */
function oneDriveTarget() {
  var raw = String(PropertiesService.getScriptProperties()
    .getProperty('ONEDRIVE_TARGET') || '').trim();
  var cut = raw.indexOf(':');
  if (cut < 1) return null;
  var kind = raw.slice(0, cut).toLowerCase();
  var rest = raw.slice(cut + 1).trim();
  if (!rest) return null;

  if (kind === 'user') {
    return { base: 'https://graph.microsoft.com/v1.0/users/' +
      encodeURIComponent(rest) + '/drive', label: rest };
  }
  if (kind === 'site' || kind === 'drive') {
    /* A library is addressed by its drive id. The full
       hostname,siteId,driveId triple is accepted because that is the shape
       Graph hands back when you look a site up, and retyping it invites a
       transcription error; only the last field is used. */
    var parts = rest.split(',');
    var driveId = String(parts[parts.length - 1] || '').trim();
    if (!driveId) return null;
    return { base: 'https://graph.microsoft.com/v1.0/drives/' +
      encodeURIComponent(driveId), label: rest };
  }
  return null;
}

/**
 * A Graph path from plain folder names: root:/Quote%20Desk/Acme%20Srl/…
 *
 * Every function below takes the names, never the built path, so a path is
 * escaped exactly once. Escaping twice turns a space into %2520 and Graph
 * answers 404 for a folder that is plainly there.
 */
function graphPath(segments) {
  var kept = [];
  for (var i = 0; i < segments.length; i++) {
    var part = String(segments[i] || '');
    if (part) kept.push(encodeURIComponent(part));
  }
  return kept.join('/');
}

/**
 * One call to Graph, with the app's own token.
 *
 * A 403 here almost always means one thing: the application permission was
 * added to the app registration and nobody pressed "Grant admin consent". It
 * says so, because the symptom otherwise - an upload that returns no error and
 * produces no file - is the kind that gets blamed on the customer's browser.
 */
function graphCall(method, url, options) {
  var creds = graphCredentials();
  if (!creds) {
    throw new Error('Microsoft 365 is not configured: set GRAPH_TENANT, ' +
      'GRAPH_CLIENT and GRAPH_SECRET in the script properties.');
  }
  var opts = options || {};
  var params = {
    method: method,
    headers: { Authorization: 'Bearer ' + graphToken(creds) },
    muteHttpExceptions: true
  };
  if (opts.json !== undefined) {
    params.contentType = 'application/json';
    params.payload = JSON.stringify(opts.json);
  } else if (opts.blob) {
    params.contentType = opts.blob.getContentType ?
      (opts.blob.getContentType() || 'application/octet-stream') : 'application/octet-stream';
    params.payload = opts.blob.getBytes();
  }

  var res = UrlFetchApp.fetch(url, params);
  var code = res.getResponseCode();
  var text = res.getContentText();
  var body = null;
  if (text) { try { body = JSON.parse(text); } catch (err) { body = null; } }

  if (code === 403) {
    throw new Error('Graph refused with 403. The app registration most likely ' +
      'has Files.ReadWrite.All added but not admin-consented. ' + text.slice(0, 200));
  }
  if (code === 404 && opts.allow404) return { code: code, body: null };
  if (code < 200 || code >= 300) {
    throw new Error('Graph HTTP ' + code + ' on ' + method.toUpperCase() + ' ' +
      url.replace(/\/\/[^/]+/, '//graph') + ' ' + text.slice(0, 300));
  }
  return { code: code, body: body, text: text };
}

/** The drive, or a refusal that names the property to set. */
function requireDrive() {
  var drive = oneDriveTarget();
  if (!drive) {
    throw new Error('ONEDRIVE_TARGET is not set in the script properties. ' +
      'Expected user:someone@dema-solutions.com or site:<hostname>,<siteId>,<driveId>.');
  }
  return drive;
}

/** The item at a path given as plain names, or null when there is nothing there. */
function graphItemByPath(segments) {
  var answer = graphCall('get',
    requireDrive().base + '/root:/' + graphPath(segments), { allow404: true });
  return answer.body;
}

/**
 * A folder at a path, made if it is not there yet.
 *
 * conflictBehavior is 'fail' and the 409 is read as "someone else made it a
 * moment ago, good" - which is what two uploads arriving together look like.
 * 'rename' would answer the same 201 and quietly leave two folders called
 * almost the same thing, and the second customer's files in the wrong one.
 */
function graphFolder(parentSegments, name) {
  var drive = requireDrive();
  var parent = graphPath(parentSegments || []);
  var url = parent ?
    drive.base + '/root:/' + parent + ':/children' :
    drive.base + '/root/children';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + graphToken(graphCredentials()) },
    payload: JSON.stringify({
      name: name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail'
    }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code === 409) {
    var found = graphItemByPath((parentSegments || []).concat([name]));
    if (found) return found;
  }
  if (code === 403) {
    throw new Error('Graph refused with 403 making folder "' + name + '". The app ' +
      'registration most likely has Files.ReadWrite.All added but not ' +
      'admin-consented. ' + res.getContentText().slice(0, 200));
  }
  if (code < 200 || code >= 300) {
    throw new Error('Graph HTTP ' + code + ' making folder "' + name + '" ' +
      res.getContentText().slice(0, 300));
  }
  return JSON.parse(res.getContentText());
}

/** Everything in a folder, following Graph's paging to the end. */
function graphChildren(itemId) {
  var drive = requireDrive();
  var url = drive.base + '/items/' + encodeURIComponent(itemId) +
    '/children?$select=id,name,size,folder,file,webUrl';
  var all = [];
  while (url) {
    var answer = graphCall('get', url, {});
    var body = answer.body || {};
    all = all.concat(body.value || []);
    url = body['@odata.nextLink'] || '';
  }
  return all;
}

/**
 * One file's bytes.
 *
 * Fetched from the pre-authenticated download URL rather than from
 * /items/{id}/content, so the app's bearer token is never sent to the storage
 * host the redirect leads to.
 */
function graphDownload(itemId, name) {
  var drive = requireDrive();
  var meta = graphCall('get', drive.base + '/items/' + encodeURIComponent(itemId) +
    '?$select=id,name,file,@microsoft.graph.downloadUrl', {}).body || {};
  var link = meta['@microsoft.graph.downloadUrl'];
  if (!link) throw new Error('no download URL for ' + (name || itemId));
  var res = UrlFetchApp.fetch(link, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('download HTTP ' + res.getResponseCode() + ' for ' + (name || itemId));
  }
  var blob = res.getBlob();
  blob.setName(name || meta.name || 'file');
  return blob;
}

/** A small file, written whole. Graph takes up to 4 MB this way; ours are PDFs. */
function graphUpload(segments, blob) {
  return graphCall('put',
    requireDrive().base + '/root:/' + graphPath(segments) + ':/content',
    { blob: blob }).body;
}

/**
 * An upload session: the permission the browser uses to put a file on OneDrive
 * by itself, in blocks, without the bytes ever passing through here.
 *
 * The URL it returns is a write permission in the clear, so it is asked for one
 * file at a time, at the moment that file is uploaded, and it is made to lapse
 * within the hour. Graph would give it seven days.
 */
function graphUploadSession(segments, name, size) {
  var until = new Date(Date.now() + UPLOAD_WINDOW_HOURS * 3600 * 1000).toISOString();
  var answer = graphCall('post',
    requireDrive().base + '/root:/' + graphPath(segments) + ':/createUploadSession', {
    json: {
      item: {
        '@microsoft.graph.conflictBehavior': 'rename',
        name: name,
        description: 'Quote Desk upload'
      },
      deferCommit: false,
      expirationDateTime: until
    }
  });
  var body = answer.body || {};
  if (!body.uploadUrl) throw new Error('Graph gave no uploadUrl for ' + name);
  return { uploadUrl: body.uploadUrl, expires: body.expirationDateTime || until,
    chunk: UPLOAD_CHUNK, size: size };
}

/** Throw something away. Used by selfTest, and by nothing that touches a customer. */
function graphDelete(itemId) {
  graphCall('delete',
    requireDrive().base + '/items/' + encodeURIComponent(itemId), {});
}

/**
 * The folder of one request, made the first time it is asked for.
 *
 * Two levels under /Quote Desk/: a folder per customer, and inside it a folder
 * per quote.
 *
 *   Quote Desk/
 *   |-- Dema Solutions Srl/
 *       |-- QUOTE 2127-2026_Dema Solutions Srl_Prova Trados_2026-11-02_15-31_ORD-20261102-JTJA/
 *
 * The customer's folder is reused across requests, so one customer's work sits
 * together and the listing stays short. The quote folder carries the number,
 * the company, the person, the day, the time and the reference - the time
 * because two quotes for the same customer on the same day are ordinary, and
 * the reference because the number is what a person quotes and the reference is
 * what the page, the links and the uploads pass about. Either one finds the
 * folder.
 *
 * `quote` is the number from quoteSerial(), and may be empty: when it is, the
 * folder is named after the reference alone, so a request never goes without a
 * folder for the sake of a cosmetic number.
 *
 * Returned as a small object with the shape the rest of the file already used,
 * so a call site reads the same whether the folder is on Drive or on OneDrive.
 */
/**
 * The plain names of the folders leading to an item, from the drive root.
 *
 * Graph writes the parent as "/drive/root:/Quote Desk/Acme Srl", escaped. It is
 * decoded back to names here because every call below escapes once, itself.
 */
function itemPath(item) {
  var parent = String(((item || {}).parentReference || {}).path || '');
  var cut = parent.indexOf('root:');
  var above = cut === -1 ? '' : parent.slice(cut + 'root:'.length);
  var segments = [];
  above.split('/').forEach(function (part) {
    if (part) segments.push(decodeURIComponent(part));
  });
  if (item && item.name) segments.push(item.name);
  return segments;
}

/*
 * A folder, whichever drive it is on.
 *
 * There are two backings now. OneDrive is the one the project chose and the
 * one every finished request should end up on. Google Drive is the fallback
 * that catches a request while Graph is refusing - without it a customer's
 * documents have nowhere at all to go, which is what happened on 14 November
 * 2026: the folder could not be made, so the files stayed on the customer's
 * own machine and the quote PDF existed only as an attachment in a mail.
 *
 * The two handles answer the same four questions and do the same two things,
 * so the callers - the delivery menu, Trados, the PDF filing, the uploads -
 * do not ask which drive they are on. `isDrive` is there for the one place
 * that has to care: the browser cannot be handed an upload session for a
 * Google folder, so it is told to send the bytes through here instead.
 */
function folderHandle(item, segments) {
  return {
    isDrive: false,
    getId: function () { return item.id; },
    getUrl: function () { return item.webUrl || ''; },
    getName: function () { return item.name || ''; },
    getPath: function () { return segments; },
    put: function (blob, name) {
      return graphUpload(segments.concat([name || blob.getName()]), blob);
    },
    list: function () { return graphChildren(item.id); },
    item: item
  };
}

/** The Google folder the fallback writes into, from ORDERS_FOLDER_ID. */
function driveRoot() {
  var id = PropertiesService.getScriptProperties().getProperty('ORDERS_FOLDER_ID');
  if (!id) throw new Error('ORDERS_FOLDER_ID is not set, so there is no fallback folder.');
  return DriveApp.getFolderById(id);
}

/** Find a child folder by name, or make it. Drive has no path addressing. */
function driveChild(parent, name) {
  var found = parent.getFoldersByName(name);
  return found.hasNext() ? found.next() : parent.createFolder(name);
}

/*
 * The same shape as folderHandle, over a Google folder.
 *
 * `list()` answers driveItem-shaped objects on purpose: `name`, `id`, `size`
 * and a `folder` that is truthy for folders, which is what every caller was
 * already written against. The extra `drive: true` is how fetchChild knows
 * which side to download from.
 */
function driveHandle(folder) {
  return {
    isDrive: true,
    getId: function () { return folder.getId(); },
    getUrl: function () { return folder.getUrl(); },
    getName: function () { return folder.getName(); },
    getPath: function () { return drivePath(folder); },
    put: function (blob, name) {
      var made = folder.createFile(blob);
      if (name) made.setName(name);
      return made;
    },
    list: function () { return driveChildren(folder); },
    folder: folder
  };
}

/*
 * The three names of a Google folder, outermost first.
 *
 * Google addresses a folder by id and has no path, but everything downstream -
 * the naming checks, what the desk is shown - was written against the three
 * OneDrive segments. Walking up to the fallback root gives the same shape:
 * the root's own name, the customer, the request.
 */
function drivePath(folder) {
  var stop = PropertiesService.getScriptProperties().getProperty('ORDERS_FOLDER_ID') || '';
  var segments = [folder.getName()];
  var up = folder;
  for (var i = 0; i < 5; i++) {
    var parents = up.getParents();
    if (!parents || !parents.hasNext()) break;
    up = parents.next();
    segments.unshift(up.getName());
    if (up.getId() === stop) break;
  }
  return segments;
}

function driveChildren(folder) {
  var out = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    out.push({ id: f.getId(), name: f.getName(), size: f.getSize(), drive: true });
  }
  var folders = folder.getFolders();
  while (folders.hasNext()) {
    var d = folders.next();
    out.push({ id: d.getId(), name: d.getName(), folder: {}, drive: true });
  }
  return out;
}

/** The children of a child folder, from either side. */
function childrenIn(child) {
  return child && child.drive
    ? driveChildren(DriveApp.getFolderById(child.id))
    : graphChildren(child.id);
}

/** One child's bytes, from either side. */
function fetchChild(child) {
  return child && child.drive
    ? DriveApp.getFileById(child.id).getBlob().setName(child.name)
    : graphDownload(child.id, child.name);
}

function orderFolder(id, customerName, code, project, quote, company) {
  /* The uploads and the decision links that follow a request know only its
     reference, and the folder is not named after it alone - so a folder once
     made is found through the row, which carries its item id. A name can be
     edited and a path can be moved; an id cannot drift. */
  var row = findOrderRow(id);
  if (row) {
    var known = String(row[HEADERS.indexOf('Drive item')] || '').trim();
    var filedOnDrive = String(row[HEADERS.indexOf('Files')] || '')
      .indexOf('drive.google.com') !== -1;

    /*
     * The Files link is read before the id is tried, because the id alone
     * does not say which drive it belongs to: a Google id asked of Graph is
     * a wasted call at best and a confusing error at worst.
     *
     * And the call is wrapped, because this is the one place a refusal must
     * not end the journey. Graph answering 403 for the whole tenant is
     * exactly the moment the fallback is needed; letting the throw out here
     * would mean the request with a folder already made is the one that
     * cannot be reopened.
     */
    if (known && !filedOnDrive) {
      try {
        var found = graphCall('get', requireDrive().base + '/items/' +
          encodeURIComponent(known) + '?$select=id,name,webUrl,parentReference',
          { allow404: true }).body;
        if (found) return folderHandle(found, itemPath(found));
      } catch (err) {}
    }
    /*
     * A Google link in `Files` used to be a dead end - the function threw and
     * said the order predated the move to OneDrive. It is not a dead end any
     * more: the fallback writes Google folders too, and an order filed there
     * has to be openable, or the delivery menu and Trados would refuse to
     * touch exactly the requests that already needed a hand. Old orders from
     * before the move open the same way, which is a gain and not a loss.
     */
    var old = String(row[HEADERS.indexOf('Files')] || '');
    if (filedOnDrive) {
      var fromUrl = (old.match(/\/folders\/([A-Za-z0-9_-]+)/) || [])[1];
      var driveId = known || fromUrl;
      if (!driveId) {
        throw new Error('Order ' + id + ' is filed on Google Drive at ' + old +
          ', but the row does not say which folder.');
      }
      return driveHandle(DriveApp.getFolderById(driveId));
    }

    /*
     * Called with a reference and nothing else - which is what an upload and a
     * second attempt after a refusal both do - the name is rebuilt from the
     * row. Without this the folder would be called after the reference alone
     * and filed under "unknown customer", so one request could end up with two
     * folders under two different names.
     */
    if (!customerName && !company && !quote) {
      var cell = function (name) { return String(row[HEADERS.indexOf(name)] || '').trim(); };
      customerName = (cell('First name') + ' ' + cell('Last name')).trim();
      code = code || cell('Customer code');
      project = project || cell('Project');
      quote = cell('Quote number');
      var firm = cell('Company');
      /* 'individual' is what the sheet holds when there is no company: it is a
         note to a reader, not a name to file anybody under. */
      company = firm && firm !== 'individual' ? firm : '';
    }
  }

  /* The day the request came in, not the day the folder happened to be made:
     a folder built on the second attempt must still carry the request's date. */
  var now = row && Object.prototype.toString.call(row[0]) === '[object Date]'
    ? row[0] : new Date();
  var tz = Session.getScriptTimeZone();
  var day = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var time = Utilities.formatDate(now, tz, 'HH-mm');

  /* The customer's own folder, under the name a person would look for: the
     company when there is one, otherwise whoever placed the request. */
  var who = folderSafe(company || customerName) || folderSafe(code) || 'unknown customer';

  var head = quote ? 'QUOTE ' + folderSafe(quote) : folderSafe(id);
  var name = [head, folderSafe(company), folderSafe(customerName),
    folderSafe(project), day, time, quote ? folderSafe(id) : '']
    .filter(function (part) { return part; })
    .join('_');

  /*
   * Three folders, outermost first: a path cannot be made in one call, and
   * asking for a child of a folder that is not there yet answers 404.
   *
   * If Graph refuses - which it does for the whole tenant until an
   * administrator consents - the same three folders are made on Google Drive
   * instead. That is deliberately a fallback and not a second home: the row
   * records which drive it landed on through the `Files` link, the desk is
   * told in `Trouble`, and when consent arrives the folders are moved across
   * by hand. One request is on one drive, never on two.
   */
  try {
    graphFolder([], ONEDRIVE_ROOT);
    graphFolder([ONEDRIVE_ROOT], who);
    var made = graphFolder([ONEDRIVE_ROOT, who], name);
    return folderHandle(made, [ONEDRIVE_ROOT, who, name]);
  } catch (err) {
    var fallback = driveChild(driveChild(driveRoot(), who), name);
    fallback.refused = String(err);
    return driveHandle(fallback);
  }
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
 * It needs one thing done once:
 *
 *  1. Entra admin centre > App registrations > New registration. Any name.
 *  2. Certificates & secrets > New client secret. Copy the VALUE, not the id.
 *  3. Overview gives the Application (client) ID and the Directory (tenant) ID.
 *  4. Put all three in Project Settings > Script properties, never in this
 *     file: GRAPH_TENANT, GRAPH_CLIENT, GRAPH_SECRET.
 *
 * Then the permissions, and there are two roads. Which one is open depends on
 * whether anyone here administers the tenant.
 *
 * APPLICATION permissions (Mail.Send, Files.ReadWrite.All) plus "Grant admin
 * consent" give the script rights of its own, with no person behind it. They
 * are also rights over EVERY mailbox and EVERY drive in the tenant, and only a
 * Global Administrator or Privileged Role Administrator can grant them.
 *
 * DELEGATED permissions are the road for the rest of us. The script acts as
 * one named person - g.gambuto@dema-solutions.com - and can reach exactly what
 * that person can reach, nothing wider. Mail.Send and Files.ReadWrite are both
 * consentable by the user themselves, so no administrator is needed:
 *
 *  a. API permissions > Microsoft Graph > Delegated permissions >
 *     Mail.Send, Files.ReadWrite, offline_access. No consent button to press.
 *  b. Authentication > Add a platform > Web, and as the redirect URI the web
 *     app's own /exec address - the one graphRedirectUri() prints.
 *  c. In the sheet, Quote Desk > Connect Microsoft 365, open the link it gives,
 *     sign in AS g.gambuto and accept. That writes GRAPH_REFRESH.
 *
 * GRAPH_REFRESH is what tells the two apart: set, the script works as the
 * person; absent, it falls back to the application grant. The refresh token
 * renews itself on every use and only dies after ninety days unused - a desk
 * that quotes weekly will never see that, and reconnecting is step (c) again.
 *
 * With tenant, client or secret missing nothing breaks: the mail goes out
 * through the Gmail account as before, and the desk's copy says so. Uploads,
 * which have no such fallback, say plainly that Graph is not configured.
 */
var GRAPH_SCOPES = 'offline_access https://graph.microsoft.com/Mail.Send ' +
  'https://graph.microsoft.com/Files.ReadWrite';

function graphCredentials() {
  try {
    var props = PropertiesService.getScriptProperties();
    var tenant = props.getProperty('GRAPH_TENANT');
    var client = props.getProperty('GRAPH_CLIENT');
    var secret = props.getProperty('GRAPH_SECRET');
    if (tenant && client && secret) {
      return {
        tenant: tenant,
        client: client,
        secret: secret,
        /* Set only on the delegated road; see the note above. */
        refresh: props.getProperty('GRAPH_REFRESH') || ''
      };
    }
  } catch (err) {}
  return null;
}

/** Where both grants are asked for. */
function graphTokenUrl(creds) {
  return 'https://login.microsoftonline.com/' + encodeURIComponent(creds.tenant) +
    '/oauth2/v2.0/token';
}

/*
 * The redirect URI, which must match what is registered in Entra character for
 * character. Run from the editor, getUrl() answers the /dev address, which is
 * not the one the world uses - hence the rewrite.
 */
function graphRedirectUri() {
  return String(ScriptApp.getService().getUrl() || '').replace(/\/dev$/, '/exec');
}

/*
 * Step (c): print the consent link, and remember a one-shot secret to
 * recognise the answer by.
 *
 * Without that secret anyone could call /exec?code=... with a code minted for
 * their own account and quietly move the desk's uploads onto their drive.
 */
function graphConsentLink() {
  var creds = graphCredentials();
  if (!creds) {
    throw new Error('Set GRAPH_TENANT, GRAPH_CLIENT and GRAPH_SECRET first.');
  }
  var state = decisionToken();
  PropertiesService.getScriptProperties().setProperty('GRAPH_STATE', state);
  var url = 'https://login.microsoftonline.com/' + encodeURIComponent(creds.tenant) +
    '/oauth2/v2.0/authorize' +
    '?client_id=' + encodeURIComponent(creds.client) +
    '&response_type=code&response_mode=query' +
    '&redirect_uri=' + encodeURIComponent(graphRedirectUri()) +
    '&scope=' + encodeURIComponent(GRAPH_SCOPES) +
    '&prompt=consent' +
    '&state=' + encodeURIComponent(state);
  Logger.log('Redirect URI to register in Entra: ' + graphRedirectUri());
  Logger.log('Consent link: ' + url);
  return url;
}

/** The same thing from the sheet's menu, for someone who is not in the editor. */
function graphConnect() {
  var url = graphConsentLink();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      '<p style="font:14px/1.5 system-ui">Register this redirect URI in Entra ' +
      '(Authentication &gt; Add a platform &gt; Web):</p>' +
      '<p style="font:12px/1.4 ui-monospace;word-break:break-all">' +
      escape(graphRedirectUri()) + '</p>' +
      '<p style="font:14px/1.5 system-ui">Then open this link and sign in as ' +
      escape(SENDER_FROM) + ':</p>' +
      '<p><a href="' + escape(url) + '" target="_blank" ' +
      'style="font:14px/1.5 system-ui">Connect Microsoft 365</a></p>')
      .setWidth(560).setHeight(260),
    'Connect Microsoft 365');
}

/*
 * The answer to that link, arriving at doGet as ?code=&state=.
 *
 * What it writes is a credential, so the state check is not optional: a
 * mismatch is answered with nothing at all.
 */
function graphConsentLanding(code, state) {
  var props = PropertiesService.getScriptProperties();
  var want = props.getProperty('GRAPH_STATE');
  var page = function (text) {
    return HtmlService.createHtmlOutput(
      '<meta name="robots" content="noindex,nofollow">' +
      '<p style="font:16px/1.5 system-ui;padding:2rem">' + escape(text) + '</p>');
  };
  if (!want || !sameSecret(String(state || ''), want)) {
    return page('Not expected.');
  }
  var creds = graphCredentials();
  if (!creds) return page('Microsoft 365 is not configured.');

  var res = UrlFetchApp.fetch(graphTokenUrl(creds), {
    method: 'post',
    payload: {
      client_id: creds.client,
      client_secret: creds.secret,
      grant_type: 'authorization_code',
      code: String(code || ''),
      redirect_uri: graphRedirectUri(),
      scope: GRAPH_SCOPES
    },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    return page('Microsoft refused: HTTP ' + res.getResponseCode() + '. ' +
      'Check that the redirect URI registered in Entra is ' + graphRedirectUri());
  }
  var body = JSON.parse(res.getContentText());
  if (!body.refresh_token) {
    return page('No refresh token came back: add offline_access to the delegated ' +
      'permissions and try again.');
  }
  props.setProperty('GRAPH_REFRESH', body.refresh_token);
  /* One use only. */
  props.deleteProperty('GRAPH_STATE');
  CacheService.getScriptCache().put('graph_token', body.access_token,
    Math.max(60, (Number(body.expires_in) || 3600) - 60));
  return page('Connected. Quote Desk now works as ' + SENDER_FROM + '.');
}

/*
 * A token, cached for the hour Microsoft gives it.
 *
 * Which grant is used depends on GRAPH_REFRESH: present, the script acts as
 * the person who consented; absent, as the application itself. Microsoft hands
 * back a fresh refresh token on most refreshes and the old one keeps working
 * for a while, but only the newest is certain to - so it is written back.
 */
function graphToken(creds) {
  var cache = CacheService.getScriptCache();
  var kept = cache.get('graph_token');
  if (kept) return kept;

  var payload = creds.refresh
    ? {
        client_id: creds.client,
        client_secret: creds.secret,
        grant_type: 'refresh_token',
        refresh_token: creds.refresh,
        scope: GRAPH_SCOPES
      }
    : {
        client_id: creds.client,
        client_secret: creds.secret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      };

  var res = UrlFetchApp.fetch(graphTokenUrl(creds), {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('token HTTP ' + res.getResponseCode() + ' ' +
      res.getContentText().slice(0, 200) + (creds.refresh
        ? ' - the Microsoft 365 connection has lapsed: run Quote Desk > ' +
          'Connect Microsoft 365 again'
        : ''));
  }
  var body = JSON.parse(res.getContentText());
  if (creds.refresh && body.refresh_token && body.refresh_token !== creds.refresh) {
    PropertiesService.getScriptProperties().setProperty('GRAPH_REFRESH', body.refresh_token);
    creds.refresh = body.refresh_token;
  }
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

/* ───────────────────────────────────── what the customer is written to in */

/*
 * The customer chooses the language of their own correspondence, and the page
 * passes it. These are the languages with a dictionary; the page speaks
 * fifteen, and the ten without one are written to in English rather than in
 * half a language.
 *
 * Adding a language is filling in a column below and nothing else: MAIL_LANGS
 * is derived from the table, so selfTest and the fallback both follow along.
 *
 * Keep the table small. Every key is five strings today and fifteen in the end,
 * so a sentence that could be left out should be left out.
 */
var MAILS = {
  quote_subject: {
    en: 'Quote {quote} - DEMA Solutions', it: 'Preventivo {quote} - DEMA Solutions',
    fr: 'Devis {quote} - DEMA Solutions', de: 'Angebot {quote} - DEMA Solutions',
    es: 'Presupuesto {quote} - DEMA Solutions' },
  quote_title: {
    en: 'Your quote', it: 'Il tuo preventivo', fr: 'Votre devis',
    de: 'Ihr Angebot', es: 'Tu presupuesto' },
  quote_intro: {
    en: 'Thank you for your request. The quote is attached, and in short it reads:',
    it: 'Grazie per la richiesta. Il preventivo è in allegato, e in breve dice:',
    fr: 'Merci pour votre demande. Le devis est en pièce jointe ; en résumé :',
    de: 'Vielen Dank für Ihre Anfrage. Das Angebot liegt bei, in Kurzform:',
    es: 'Gracias por tu solicitud. El presupuesto va adjunto; en resumen:' },
  quote_valid: {
    en: 'This quote stands until {date}.', it: 'Il preventivo è valido fino al {date}.',
    fr: 'Ce devis est valable jusqu\'au {date}.', de: 'Dieses Angebot gilt bis {date}.',
    es: 'Este presupuesto es válido hasta el {date}.' },
  quote_ask: {
    en: 'If it suits you, accept it and we start. If it does not, decline it and nothing happens.',
    it: 'Se ti va bene, accettalo e partiamo. Se non ti va bene, rifiutalo e non succede nulla.',
    fr: 'S\'il vous convient, acceptez-le et nous commençons. Sinon, refusez-le et rien ne se passe.',
    de: 'Passt es, nehmen Sie es an und wir beginnen. Passt es nicht, lehnen Sie ab und nichts geschieht.',
    es: 'Si te conviene, acéptalo y empezamos. Si no, recházalo y no ocurre nada.' },
  btn_accept: {
    en: 'Accept the quote', it: 'Accetto il preventivo', fr: 'J\'accepte le devis',
    de: 'Angebot annehmen', es: 'Acepto el presupuesto' },
  btn_decline: {
    en: 'Decline', it: 'Rifiuto', fr: 'Refuser', de: 'Ablehnen', es: 'Rechazar' },
  quote_files: {
    en: 'Your documents are with us and nothing is done with them until you answer.',
    it: 'I tuoi documenti sono al sicuro e nessuno li tocca finché non rispondi.',
    fr: 'Vos documents sont chez nous et personne n\'y touche avant votre réponse.',
    de: 'Ihre Dokumente liegen bei uns und bleiben unberührt, bis Sie antworten.',
    es: 'Tus documentos están con nosotros y nadie los toca hasta que respondas.' },
  any_question: {
    en: 'Any question, answer this email and it reaches the desk handling the job.',
    it: 'Per qualsiasi domanda, rispondi a questa mail: arriva a chi segue il lavoro.',
    fr: 'Pour toute question, répondez à cet e-mail : il arrive chez la personne qui suit le dossier.',
    de: 'Bei Fragen antworten Sie einfach auf diese E-Mail; sie erreicht die zuständige Person.',
    es: 'Para cualquier duda, responde a este correo: llega a quien lleva el trabajo.' },
  row_quote: {
    en: 'Quote', it: 'Preventivo', fr: 'Devis', de: 'Angebot', es: 'Presupuesto' },
  row_reference: {
    en: 'Reference', it: 'Riferimento', fr: 'Référence', de: 'Referenz', es: 'Referencia' },
  row_solution: {
    en: 'Solution', it: 'Soluzione', fr: 'Solution', de: 'Leistung', es: 'Solución' },
  row_from: { en: 'From', it: 'Da', fr: 'De', de: 'Aus', es: 'De' },
  row_into: { en: 'Into', it: 'Verso', fr: 'Vers', de: 'Nach', es: 'A' },
  row_words: { en: 'Words', it: 'Parole', fr: 'Mots', de: 'Wörter', es: 'Palabras' },
  row_delivery: {
    en: 'Delivery', it: 'Consegna', fr: 'Livraison', de: 'Lieferung', es: 'Entrega' },
  row_total: {
    en: 'Total', it: 'Totale', fr: 'Total', de: 'Gesamt', es: 'Total' },
  page_title_accept: {
    en: 'Accept quote {quote}', it: 'Accetta il preventivo {quote}',
    fr: 'Accepter le devis {quote}', de: 'Angebot {quote} annehmen',
    es: 'Aceptar el presupuesto {quote}' },
  page_title_decline: {
    en: 'Decline quote {quote}', it: 'Rifiuta il preventivo {quote}',
    fr: 'Refuser le devis {quote}', de: 'Angebot {quote} ablehnen',
    es: 'Rechazar el presupuesto {quote}' },
  page_confirm_accept: {
    en: 'Nothing is agreed yet. Press the button to accept this quote.',
    it: 'Non è ancora accettato nulla. Premi il bottone per accettare il preventivo.',
    fr: 'Rien n\'est encore accepté. Appuyez sur le bouton pour accepter ce devis.',
    de: 'Noch ist nichts angenommen. Drücken Sie die Taste, um das Angebot anzunehmen.',
    es: 'Todavía no se ha aceptado nada. Pulsa el botón para aceptar el presupuesto.' },
  page_confirm_decline: {
    en: 'Press the button to decline this quote. Nothing will be done.',
    it: 'Premi il bottone per rifiutare il preventivo. Non verrà fatto nulla.',
    fr: 'Appuyez sur le bouton pour refuser ce devis. Rien ne sera fait.',
    de: 'Drücken Sie die Taste, um abzulehnen. Es wird nichts unternommen.',
    es: 'Pulsa el botón para rechazar el presupuesto. No se hará nada.' },
  page_done_accept: {
    en: 'Accepted. The confirmation is on its way by email.',
    it: 'Accettato. La conferma sta arrivando per mail.',
    fr: 'Accepté. La confirmation arrive par e-mail.',
    de: 'Angenommen. Die Bestätigung kommt per E-Mail.',
    es: 'Aceptado. La confirmación llega por correo.' },
  page_done_decline: {
    en: 'Declined. Nothing will be done, and your files stay as they are.',
    it: 'Rifiutato. Non verrà fatto nulla e i tuoi file restano come sono.',
    fr: 'Refusé. Rien ne sera fait et vos fichiers restent en place.',
    de: 'Abgelehnt. Es wird nichts unternommen, Ihre Dateien bleiben unverändert.',
    es: 'Rechazado. No se hará nada y tus archivos quedan como están.' },
  page_already: {
    en: 'This quote was already answered on {date}.',
    it: 'Questo preventivo ha già avuto una risposta il {date}.',
    fr: 'Ce devis a déjà reçu une réponse le {date}.',
    de: 'Dieses Angebot wurde am {date} schon beantwortet.',
    es: 'Este presupuesto ya se respondió el {date}.' },
  page_expired: {
    en: 'This quote lapsed on {date}. Write to {email} and we will make a new one.',
    it: 'Questo preventivo è scaduto il {date}. Scrivi a {email} e ne facciamo uno nuovo.',
    fr: 'Ce devis a expiré le {date}. Écrivez à {email} et nous en ferons un nouveau.',
    de: 'Dieses Angebot ist am {date} abgelaufen. Schreiben Sie an {email}, wir machen ein neues.',
    es: 'Este presupuesto caducó el {date}. Escribe a {email} y haremos uno nuevo.' },
  page_notfound: {
    en: 'We cannot find that quote.', it: 'Questo preventivo non si trova.',
    fr: 'Ce devis est introuvable.', de: 'Dieses Angebot ist nicht zu finden.',
    es: 'No encontramos ese presupuesto.' },
  confirm_subject: {
    en: 'Order confirmed - {quote}', it: 'Ordine confermato - {quote}',
    fr: 'Commande confirmée - {quote}', de: 'Auftrag bestätigt - {quote}',
    es: 'Pedido confirmado - {quote}' },
  confirm_title: {
    en: 'Your order is confirmed', it: 'Il tuo ordine è confermato',
    fr: 'Votre commande est confirmée', de: 'Ihr Auftrag ist bestätigt',
    es: 'Tu pedido está confirmado' },
  confirm_intro: {
    en: 'Thank you. The order confirmation is attached and the work starts.',
    it: 'Grazie. La conferma d\'ordine è in allegato e il lavoro parte.',
    fr: 'Merci. La confirmation de commande est en pièce jointe et le travail démarre.',
    de: 'Vielen Dank. Die Auftragsbestätigung liegt bei, die Arbeit beginnt.',
    es: 'Gracias. La confirmación de pedido va adjunta y el trabajo empieza.' },
  decline_subject: {
    en: 'Quote {quote} declined', it: 'Preventivo {quote} rifiutato',
    fr: 'Devis {quote} refusé', de: 'Angebot {quote} abgelehnt',
    es: 'Presupuesto {quote} rechazado' },
  decline_title: {
    en: 'Your quote is declined', it: 'Preventivo rifiutato',
    fr: 'Devis refusé', de: 'Angebot abgelehnt', es: 'Presupuesto rechazado' },
  decline_intro: {
    en: 'Noted - nothing will be done. Your documents stay with us untouched; say the word and we remove them.',
    it: 'Preso nota: non verrà fatto nulla. I tuoi documenti restano intatti; se vuoi li cancelliamo.',
    fr: 'C\'est noté : rien ne sera fait. Vos documents restent intacts ; dites-le et nous les supprimons.',
    de: 'Notiert - es wird nichts unternommen. Ihre Dokumente bleiben unberührt; auf Wunsch löschen wir sie.',
    es: 'Anotado: no se hará nada. Tus documentos quedan intactos; dilo y los borramos.' },
  pdf_doc_quote: {
    en: 'QUOTE', it: 'PREVENTIVO', fr: 'DEVIS', de: 'ANGEBOT', es: 'PRESUPUESTO' },
  pdf_doc_order: {
    en: 'ORDER CONFIRMATION', it: 'CONFERMA D\'ORDINE', fr: 'CONFIRMATION DE COMMANDE',
    de: 'AUFTRAGSBESTÄTIGUNG', es: 'CONFIRMACIÓN DE PEDIDO' },
  pdf_no: { en: 'no.', it: 'n.', fr: 'n°', de: 'Nr.', es: 'n.º' },
  pdf_of: { en: 'of', it: 'del', fr: 'du', de: 'vom', es: 'del' },
  pdf_vatid: {
    en: 'VAT ID', it: 'PARTITA IVA', fr: 'N° TVA', de: 'USt-IDNR.', es: 'NIF' },
  pdf_email: { en: 'EMAIL', it: 'EMAIL', fr: 'E-MAIL', de: 'E-MAIL', es: 'CORREO' },
  pdf_recipient: {
    en: 'RECIPIENT', it: 'DESTINATARIO', fr: 'DESTINATAIRE',
    de: 'EMPFÄNGER', es: 'DESTINATARIO' },
  pdf_subject: {
    en: 'SUBJECT', it: 'OGGETTO', fr: 'OBJET', de: 'BETREFF', es: 'ASUNTO' },
  pdf_description: {
    en: 'DESCRIPTION', it: 'DESCRIZIONE', fr: 'DESCRIPTION',
    de: 'BESCHREIBUNG', es: 'DESCRIPCIÓN' },
  pdf_quantity: {
    en: 'QUANTITY', it: 'QUANTITÀ', fr: 'QUANTITÉ', de: 'MENGE', es: 'CANTIDAD' },
  pdf_amount: {
    en: 'AMOUNT', it: 'IMPORTO', fr: 'MONTANT', de: 'BETRAG', es: 'IMPORTE' },
  pdf_payment: {
    en: 'PAYMENT METHOD', it: 'MODALITÀ DI PAGAMENTO', fr: 'MODE DE PAIEMENT',
    de: 'ZAHLUNGSWEISE', es: 'FORMA DE PAGO' },
  pdf_delivery: {
    en: 'DELIVERY', it: 'CONSEGNA', fr: 'LIVRAISON', de: 'LIEFERUNG', es: 'ENTREGA' },
  pdf_vatsummary: {
    en: 'VAT SUMMARY', it: 'RIEPILOGO IVA', fr: 'RÉCAPITULATIF TVA',
    de: 'UST-ÜBERSICHT', es: 'RESUMEN DE IVA' },
  pdf_totalamount: {
    en: 'TOTAL AMOUNT', it: 'IMPORTO TOTALE', fr: 'MONTANT TOTAL',
    de: 'GESAMTBETRAG', es: 'IMPORTE TOTAL' },
  pdf_taxes: { en: 'TAXES', it: 'IMPOSTE', fr: 'TAXES', de: 'STEUERN', es: 'IMPUESTOS' },
  pdf_total: {
    en: 'Total amount', it: 'Importo totale', fr: 'Montant total',
    de: 'Gesamtbetrag', es: 'Importe total' },
  pdf_services: {
    en: 'Translation services', it: 'Servizi di traduzione',
    fr: 'Services de traduction', de: 'Übersetzungsleistungen',
    es: 'Servicios de traducción' },
  pdf_language: {
    en: 'language', it: 'lingua', fr: 'langue', de: 'Sprache', es: 'idioma' },
  pdf_languages: {
    en: 'languages', it: 'lingue', fr: 'langues', de: 'Sprachen', es: 'idiomas' },
  pdf_words: { en: 'words', it: 'parole', fr: 'mots', de: 'Wörter', es: 'palabras' },
  pdf_dtp: {
    en: 'Desktop publishing', it: 'Impaginazione', fr: 'Mise en page',
    de: 'Layout', es: 'Maquetación' },
  pdf_dtp_note: {
    en: '{n} pages per target language', it: '{n} pagine per lingua di arrivo',
    fr: '{n} pages par langue cible', de: '{n} Seiten pro Zielsprache',
    es: '{n} páginas por idioma de destino' },
  pdf_valid: {
    en: 'Valid until {date}', it: 'Valido fino al {date}', fr: 'Valable jusqu\'au {date}',
    de: 'Gültig bis {date}', es: 'Válido hasta el {date}' },
  pdf_note_quote: {
    en: 'This is a quote, not a fiscal invoice.',
    it: 'Questo è un preventivo, non una fattura.',
    fr: 'Ceci est un devis, pas une facture.',
    de: 'Dies ist ein Angebot, keine Rechnung.',
    es: 'Este es un presupuesto, no una factura.' },
  pdf_note_order: {
    en: 'This is an order confirmation, not a fiscal invoice.',
    it: 'Questa è una conferma d\'ordine, non una fattura.',
    fr: 'Ceci est une confirmation de commande, pas une facture.',
    de: 'Dies ist eine Auftragsbestätigung, keine Rechnung.',
    es: 'Esta es una confirmación de pedido, no una factura.' }
};

/** The languages a dictionary actually covers, read off the table itself. */
function mailLangs() {
  var found = {};
  for (var key in MAILS) {
    if (!Object.prototype.hasOwnProperty.call(MAILS, key)) continue;
    for (var lang in MAILS[key]) {
      if (Object.prototype.hasOwnProperty.call(MAILS[key], lang)) found[lang] = true;
    }
  }
  return Object.keys(found).sort();
}

/**
 * The language to write in: what the customer asked for, if we have it.
 *
 * The page speaks fifteen languages and this table five, so a Finnish customer
 * is written to in English. That is a choice: a message half in one language
 * and half in another reads worse than a message in one that is not theirs.
 */
function mailLang(asked) {
  var want = String(asked || '').toLowerCase().slice(0, 2);
  return (want && MAILS.quote_title[want]) ? want : 'en';
}

/** One line, in one language, with {placeholders} filled in. */
function mail(key, lang, fill) {
  var entry = MAILS[key];
  if (!entry) return key;
  var text = entry[mailLang(lang)] || entry.en || '';
  if (fill) {
    for (var name in fill) {
      if (!Object.prototype.hasOwnProperty.call(fill, name)) continue;
      text = text.split('{' + name + '}').join(String(fill[name]));
    }
  }
  return text;
}

/* ──────────────────────────────── the quote, and the answer to it */

/** The secret in the two links. The reference alone is guessable; this is not. */
function decisionToken() {
  var hex = '0123456789abcdef';
  var out = '';
  for (var i = 0; i < 32; i++) {
    out += hex.charAt(Math.floor(Math.random() * 16));
  }
  return out;
}

/** When a quote taken at `from` stops standing. */
function quoteExpiry(from) {
  return new Date(from.getTime() + QUOTE_VALID_DAYS * 24 * 3600 * 1000).toISOString();
}

/**
 * One of the two links.
 *
 * It carries the answer, the reference, the secret and the language - and it
 * carries them in a query string, which is the one place a secret should not
 * be. It is tolerable only because following this URL changes nothing: the
 * page it opens has a button, and the button posts. See decisionPage.
 */
function decisionUrl(id, token, answer, lang) {
  return ScriptApp.getService().getUrl() +
    '?a=' + encodeURIComponent(answer) +
    '&id=' + encodeURIComponent(id) +
    '&t=' + encodeURIComponent(token) +
    '&lang=' + encodeURIComponent(lang || 'en');
}

/**
 * Why this page exists at all, in one paragraph.
 *
 * Outlook, Gmail and corporate mail gateways fetch the links in a message to
 * see where they go. A link that accepted a quote on being fetched would be
 * accepted by a scanner minutes after it was sent, and the customer would never
 * have seen it. So a GET only draws the quote and a button; the button posts
 * action 'decide', and only that writes anything. Do not shorten this.
 *
 * The page is also told not to be indexed, and loads nothing from anywhere
 * else, so the secret in its URL is not handed to a third party in a Referer.
 */
function decisionPage(decision, id, token, lang) {
  var at = quoteRow(id, token);
  var language = mailLang(lang || (at ? at('Language') : ''));
  var title, body;

  if (!at) {
    title = mail('page_notfound', language);
    body = '<p>' + escape(mail('page_notfound', language)) + '</p>';
    return decisionHtml(title, body, language);
  }

  var quote = at('Quote number') || at('Reference');
  var status = String(at('Status') || '');
  var decided = at('Decided');

  if (status === 'Accepted' || status === 'Declined') {
    title = escape(quote);
    body = '<p>' + escape(mail('page_already', language,
      { date: decided ? when(new Date(decided).toISOString()) : '-' })) + '</p>' +
      '<p>' + escape(status === 'Accepted'
        ? mail('page_done_accept', language) : mail('page_done_decline', language)) + '</p>';
    return decisionHtml(title, body, language);
  }

  if (quoteLapsed(at)) {
    title = escape(quote);
    body = '<p>' + escape(mail('page_expired', language,
      { date: when(quoteExpiry(new Date(at('Received')))), email: DESK })) + '</p>';
    return decisionHtml(title, body, language);
  }

  title = mail(decision === 'accept' ? 'page_title_accept' : 'page_title_decline',
    language, { quote: quote });
  body =
    '<p>' + escape(mail(decision === 'accept'
      ? 'page_confirm_accept' : 'page_confirm_decline', language)) + '</p>' +
    table([
      [mail('row_quote', language), quote],
      [mail('row_solution', language), at('Solution')],
      [mail('row_into', language), at('Into')],
      [mail('row_delivery', language), at('Delivery')],
      [mail('row_total', language), at('Total') + ' ' + at('Currency')]
    ]) +
    '<form method="post" action="' + ScriptApp.getService().getUrl() + '" id="f">' +
    '<input type="hidden" name="payload" value="">' +
    '</form>' +
    '<p style="margin:22px 0"><button type="button" onclick="go()" ' +
    'style="background:#1b2a6b;color:#fff;border:0;padding:12px 20px;border-radius:6px;' +
    'font:700 15px Helvetica,Arial,sans-serif;cursor:pointer">' +
    escape(mail(decision === 'accept' ? 'btn_accept' : 'btn_decline', language)) +
    '</button></p>' +
    '<div id="said" style="font-weight:700"></div>' +
    '<script>' +
    'function go(){' +
    'var b=document.querySelector("button");b.disabled=true;' +
    'google.script.run.withSuccessHandler(function(r){' +
    'document.getElementById("said").textContent=r;b.style.display="none";})' +
    '.decideFromPage(' + JSON.stringify({ id: id, token: token, decision: decision }) + ');' +
    '}</script>';
  return decisionHtml(title, body, language);
}

/** The page itself: no outside requests, and not for search engines. */
function decisionHtml(title, body, language) {
  var page = '<!doctype html><html lang="' + escape(language) + '"><head>' +
    '<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<meta name="robots" content="noindex,nofollow">' +
    '<title>' + escape(title) + '</title></head>' +
    '<body style="margin:0;padding:32px;font:15px/1.6 Helvetica,Arial,sans-serif;color:#131a38">' +
    '<div style="max-width:620px;margin:0 auto">' +
    '<h1 style="font-size:22px;margin:0 0 18px">' + escape(title) + '</h1>' + body +
    '<p style="margin-top:28px;color:#667;font-size:12px">' + escape(COMPANY.name) +
    ' · ' + COMPANY.site + '</p></div></body></html>';
  return HtmlService.createHtmlOutput(page)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * The button's own call.
 *
 * google.script.run reaches a named function, not doPost, and it runs as the
 * deployment rather than as whoever is looking - which is what we want: the
 * customer has no Google account and must not need one.
 */
function decideFromPage(req) {
  var answer = decide({ id: (req || {}).id, token: (req || {}).token,
    decision: (req || {}).decision });
  var language = mailLang((req || {}).lang);
  if (answer.ok) {
    return answer.decision === 'accept'
      ? mail('page_done_accept', language) : mail('page_done_decline', language);
  }
  if (answer.error === 'expired') return mail('page_expired', language,
    { date: answer.until || '-', email: DESK });
  if (answer.error === 'already decided') return mail('page_already', language,
    { date: answer.decided || '-' });
  return mail('page_notfound', language);
}

/** A reader for one row, found by reference and secret, or null. */
function quoteRow(id, token) {
  if (!/^[A-Za-z0-9-]{4,40}$/.test(String(id || ''))) return null;
  var row = findOrderRow(id);
  if (!row) return null;
  var want = String(row[HEADERS.indexOf('Decision token')] || '');
  if (!want || !sameSecret(String(token || ''), want)) return null;
  return function (name) { return row[HEADERS.indexOf(name)]; };
}

/** Has this quote stood longer than it was meant to? */
function quoteLapsed(at) {
  var received = at('Received');
  var taken = Object.prototype.toString.call(received) === '[object Date]'
    ? received.getTime() : Date.parse(received);
  if (!taken) return false;
  return Date.now() > taken + QUOTE_VALID_DAYS * 24 * 3600 * 1000;
}

/**
 * The decision, written down once.
 *
 * Checks the secret, the state and the date, then writes Status and Decided and
 * sends what follows. A second press finds the row already decided and changes
 * nothing: a customer who clicks accept twice has not ordered twice, and one
 * who clicks decline after accepting has not cancelled by link - that is a
 * conversation with the desk, not a URL.
 */
function decide(req) {
  var id = String((req || {}).id || '');
  var decision = (req || {}).decision === 'accept' ? 'accept' : 'decline';
  var at = quoteRow(id, (req || {}).token);
  if (!at) return { ok: false, error: 'not found' };

  var status = String(at('Status') || '');
  if (status === 'Accepted' || status === 'Declined') {
    var was = at('Decided');
    return { ok: false, error: 'already decided', status: status,
      decided: was ? when(new Date(was).toISOString()) : '' };
  }
  if (quoteLapsed(at)) {
    markStatus(id, 'Expired', null);
    return { ok: false, error: 'expired',
      until: when(quoteExpiry(new Date(at('Received')))) };
  }

  var now = new Date();
  markStatus(id, decision === 'accept' ? 'Accepted' : 'Declined', now);

  var language = mailLang(at('Language'));
  var quote = at('Quote number') || id;
  var order = orderFromRow(at);
  var net = Number(at('Net')) || 0;
  var vat = Number(at('VAT')) || 0;
  var gross = Number(at('Total')) || round2(net + vat);

  try {
    if (decision === 'accept') {
      /* Asked for only now: a card link belongs with a confirmed order, not
         with a price the customer has not yet agreed to. */
      var payUrl = payLink(order, gross);
      var pdf = confirmationPdf(order, net, vat, gross, language, 'order', quote);
      try {
        var folder = orderFolder(id, '');
        folder.put(pdf);
      } catch (err) {}
      send({
        to: at('Email'), name: SENDER_NAME, replyTo: DESK,
        subject: mail('confirm_subject', language, { quote: quote, id: id }),
        htmlBody: confirmBody(order, net, vat, gross, language, quote, payUrl),
        attachments: [pdf]
      });
    } else {
      send({
        to: at('Email'), name: SENDER_NAME, replyTo: DESK,
        subject: mail('decline_subject', language, { quote: quote, id: id }),
        htmlBody: declineBody(order, language, quote)
      });
    }
    send({
      to: DESK, name: SENDER_NAME, replyTo: at('Email'),
      subject: (decision === 'accept' ? 'ACCEPTED ' : 'declined ') + quote + ' - ' + id,
      htmlBody: deskDecisionBody(null, at, decision)
    });
  } catch (err) {
    /* The decision is already written, and that is the part that matters. A
       mail that did not go out is told to the desk, not hidden from it. */
    try {
      MailApp.sendEmail(DESK, 'Quote Desk: ' + id + ' was ' + decision +
        'ed but a message failed', String(err));
    } catch (ignored) {}
  }

  return { ok: true, decision: decision, id: id };
}

/** Status and the date of the decision, in the row of one request. */
function markStatus(id, status, at) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sheet) return;
  var last = sheet.getLastRow();
  if (last < 2) return;
  var refs = sheet.getRange(2, HEADERS.indexOf('Reference') + 1, last - 1, 1).getValues();
  for (var i = refs.length - 1; i >= 0; i--) {
    if (String(refs[i][0]) === String(id)) {
      sheet.getRange(i + 2, HEADERS.indexOf('Status') + 1).setValue(status);
      if (at) sheet.getRange(i + 2, HEADERS.indexOf('Decided') + 1).setValue(at);
      return;
    }
  }
}

/**
 * A row read back into the shape the mail and the PDF builders expect.
 *
 * They were written for the object the page posts, and they stay written for
 * it: the acceptance is a different moment in time, not a different document.
 * What the sheet cannot give back - the stage names, the per-file detail - is
 * not in these documents anyway.
 */
function orderFromRow(at) {
  var pages = String(at('Desktop publishing') || '');
  var count = localesFrom(at('Into')).length || String(at('Into') || '').split(',').length;
  return {
    id: at('Reference'),
    solution: at('Solution'),
    source: at('From'),
    targets: String(at('Into') || '').split(',').map(function (s) { return s.trim(); })
      .filter(String),
    words: count ? Math.round((Number(at('Words')) || 0) / count) : Number(at('Words')) || 0,
    subject: at('Subject'),
    layout: /^\d+$/.test(pages) ? { pages: Number(pages) } : null,
    eta: at('Delivery') ? String(at('Delivery')) : '',
    currency: at('Currency') || 'EUR',
    vatTreatment: at('VAT treatment') || '',
    project: at('Project') || '',
    stages: [],
    customer: {
      firstName: at('First name'), lastName: at('Last name'),
      name: [at('First name'), at('Last name')].filter(String).join(' '),
      email: at('Email'), reference: at('Customer reference'), note: at('Notes')
    },
    billing: {
      company: !!at('Company'), companyName: at('Company'),
      name: [at('First name'), at('Last name')].filter(String).join(' '),
      vatNumber: at('VAT number'), taxCode: at('Tax code'),
      address1: at('Address'), city: at('City'), postalCode: at('Postal code'),
      countryName: at('Country'), telephone: at('Telephone')
    }
  };
}

/* ──────────────────────────────────────────────────────────────── the web app */

/**
 * The decision links land here, and nothing else does.
 *
 * A GET only ever *shows*. The page it serves carries a button, and that button
 * posts. This is the whole defence against a mail client that follows links to
 * scan them: a scanner sees the summary, a person presses the button.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.a === 'yes' || p.a === 'no') {
    return decisionPage(p.a === 'yes' ? 'accept' : 'decline', p.id, p.t, p.lang);
  }
  /* Microsoft coming back from the consent screen; see graphConsentLink(). */
  if (p.code && p.state) {
    return graphConsentLanding(p.code, p.state);
  }
  return json({ ok: true, service: 'quote-desk-orders' });
}

function doPost(e) {
  try {
    var order = JSON.parse(e.postData.contents);

    /*
     * The same /exec serves several callers. The quote page posts a request and
     * sends action 'quote' - or nothing at all, which is what the page
     * published before this change sends, and it must keep working until the
     * new page is pushed. The back office posts {action: 'orders', token} and
     * gets the book back. The token travels in the body on purpose: a query
     * string ends up in browser history and in referrers, and Apps Script
     * cannot read a custom header (a custom header would force a CORS
     * preflight, and the web app answers no OPTIONS).
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

    /*
     * The desk's own actions. They read and rewrite other people's requests,
     * so they are behind exactly the login that opens the book, and never
     * behind the reference alone.
     */
    if (order.action === 'quotePdf' || order.action === 'resend' ||
        order.action === 'retryFiles') {
      var mayDesk = order.user ? adminLogin(order.user, order.pass) : sameToken(order.token);
      if (!mayDesk) return json({ ok: false, error: 'refused' });
      if (order.action === 'quotePdf') return json(deskQuotePdf(order));
      if (order.action === 'resend') return json(deskResend(order));
      return json(deskRetryFiles(order));
    }

    if (order.action === 'upload') return json(receiveFile(order));
    if (order.action === 'uploadUrl') return json(uploadTicket(order));
    if (order.action === 'uploadDone') return json(uploadDone(order));
    if (order.action === 'status') return json(orderStatus(order));
    if (order.action === 'decide') return json(decide(order));

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

    var quote = quoteSerial();
    var who = splitName(order);
    var billing = order.billing || {};
    var lang = mailLang(order.language);
    var token = decisionToken();

    /*
     * THE ROW IS WRITTEN FIRST, AND BEFORE ANYTHING THAT CAN BE REFUSED.
     *
     * It used to be written after the OneDrive folder, and that was the wrong
     * way round: the folder needs Microsoft Graph, Graph needs a tenant that
     * has said yes, and the sheet needs nothing at all. One refusal from Graph
     * threw before the row existed, and the whole request vanished - no line in
     * the book, nothing in the back office, a customer who had filled in the
     * form and a desk that never heard of them.
     *
     * So: the request is recorded always, and every step after it is allowed to
     * fail on its own. What failed is written into Trouble on the same row and
     * repeated in the desk's copy, where somebody will read it.
     */
    var trouble = [];
    record(order, net, vat, gross, repriced, '', quote,
      { language: lang, token: token, item: '' });

    var folder = null;
    try {
      folder = orderFolder(order.id, (who[0] + ' ' + who[1]).trim(),
        customerCode(order), order.project, quote,
        billing.company ? billing.companyName : '');
      /* The row learns where its folder is the moment there is one. */
      rememberFolder(order.id, folder);
      /* A folder on Google Drive is a folder, and the request is safe - but it
         is not where this work belongs, and nobody would notice from a link
         alone. The desk is told, so the folder gets moved once Graph answers
         again rather than being left behind quietly. */
      if (folder.isDrive) {
        trouble.push('Microsoft refused the folder, so the files went to Google Drive ' +
          'instead (' + folder.getUrl() + '). Move them to OneDrive once the ' +
          'administrator has given consent.');
      }
    } catch (err) {
      trouble.push('No folder could be made for this request, on either drive (' +
        err + '). The customer was asked to send the documents by mail.');
    }

    var pdf = null;
    try {
      pdf = quotePdf(order, net, vat, gross, lang, quote);
    } catch (err) {
      trouble.push('The quote PDF could not be built (' + err + ').');
    }
    /* Filed, not only attached: a quote nobody kept a copy of is a document
       that cannot be produced when it is asked for. */
    if (folder && pdf) {
      try {
        folder.put(pdf);
      } catch (err) {
        trouble.push('The quote PDF was not filed beside the request (' + err + ').');
      }
    }

    var senderTrouble = null;
    try {
      senderTrouble = send({
        to: order.customer.email,
        name: SENDER_NAME,
        // The account that sends is not the account that handles the job, so an
        // answer has to be pointed at the desk rather than at the sending mailbox.
        replyTo: DESK,
        subject: mail('quote_subject', lang, { id: order.id, quote: quote || order.id }),
        htmlBody: quoteBody(order, net, vat, gross, lang, quote, token),
        attachments: pdf ? [pdf] : []
      });
    } catch (err) {
      trouble.push('The quote did not reach the customer (' + err + ').');
    }
    try {
      send({
        to: DESK,
        name: SENDER_NAME,
        replyTo: order.customer.email,
        subject: 'New quote ' + (quote || order.id) + ' - ' + order.id + ' - ' +
          order.solution + ' - ' + gross.toFixed(2) + ' ' + order.currency,
        htmlBody: deskBody(order, net, vat, gross, repriced,
          [senderTrouble].concat(trouble).filter(function (x) { return x; }).join(' '), lang),
        attachments: pdf ? [pdf] : []
      });
    } catch (err) {
      trouble.push('The desk copy did not go out (' + err + ').');
    }

    if (trouble.length) noteTrouble(order.id, trouble.join(' '));

    /* `files` tells the page whether there is anywhere to upload to. With no
       folder it stops asking for upload tickets instead of failing twelve
       times, and the request is on the desk either way. */
    return json({ ok: true, id: order.id, quote: quote, files: !!folder });
  } catch (err) {
    // The request must not vanish because the mail or the sheet misbehaved.
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
/* ---------------------------------------------------------------- paying */

/**
 * A card link for one order, when a card is wanted.
 *
 * Be clear about the cost before switching it on: there is no free way to take
 * a card. Stripe keeps roughly 1.5% + 0.25 EUR of a European card payment and
 * charges no monthly fee; every other processor charges something comparable.
 * Bank transfer, which the confirmation already carries, stays free and stays
 * the default.
 *
 * One thing to do once: make an account, then put its SECRET key (sk_live_...,
 * or sk_test_... while trying it) in Project Settings > Script properties under
 * STRIPE_SECRET. Never in this file. Without it nothing changes and no link is
 * offered - the customer simply transfers as before.
 *
 * The amount is taken from what the script itself worked out, never from the
 * page: a price the browser proposed is not a price to charge.
 */
function stripeKey() {
  try {
    return PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET') || '';
  } catch (err) {
    return '';
  }
}

/**
 * @returns {string} a payment page for this order, or '' when cards are off
 */
function payLink(order, gross) {
  var key = stripeKey();
  if (!key || !(gross > 0)) return '';

  var currency = String(order.currency || 'EUR').toLowerCase();
  /* Stripe counts in the currency's smallest unit: cents, not euros. */
  var amount = Math.round(gross * 100);

  var form = {
    mode: 'payment',
    'payment_method_types[0]': 'card',
    client_reference_id: order.id,
    customer_email: order.customer.email,
    'line_items[0][quantity]': '1',
    'line_items[0][price_data][currency]': currency,
    'line_items[0][price_data][unit_amount]': String(amount),
    'line_items[0][price_data][product_data][name]':
      'Translation ' + order.id + ' - ' + (order.solution || ''),
    success_url: 'https://ai-dema-solutions.com/quote/status/?ref=' + encodeURIComponent(order.id),
    cancel_url: 'https://ai-dema-solutions.com/quote/'
  };

  try {
    var res = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', {
      method: 'post',
      payload: form,
      headers: { Authorization: 'Bearer ' + key },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error('HTTP ' + code + ' ' + res.getContentText().slice(0, 200));
    }
    return JSON.parse(res.getContentText()).url || '';
  } catch (err) {
    /* A card link that cannot be made must not lose the order: the customer
       still has the IBAN on the confirmation, and the desk is told. */
    try {
      MailApp.sendEmail(DESK, 'Quote Desk: no card link for ' + order.id, String(err));
    } catch (ignored) {}
    return '';
  }
}

/** The paragraph the confirmation carries when a card link exists. */
function payParagraph(link) {
  if (!link) return '';
  return '<p style="margin:18px 0"><a href="' + link + '" ' +
    'style="background:#1c244b;color:#ffffff;padding:12px 22px;' +
    'text-decoration:none;border-radius:2px;display:inline-block">Pay by card</a></p>' +
    '<p style="color:#6e768c;font-size:13px">Or transfer to the IBAN below, quoting the ' +
    'order reference. Work starts when payment is received, either way.</p>';
}


/* ------------------------------------------------------------- the self-test */

/**
 * Check this deployment from the inside, and say plainly what works.
 *
 * Everything here was written against stubs, never against the real Google:
 * the first time it meets Drive, the sheet, a token and a mailbox is the first
 * order a customer places. That is a poor moment to discover a missing
 * permission. Run this from the editor instead - Run > selfTest - and read the
 * log. It writes a row and deletes it, makes a folder and deletes it, and
 * sends one email to the desk.
 *
 * Nothing here touches a real order.
 */
function selfTest() {
  var out = [];
  var say = function (ok, what, detail) {
    out.push((ok ? '[ok]   ' : '[FAIL] ') + what + (detail ? ' - ' + detail : ''));
  };

  // 1. the sheet
  try {
    var book = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = book.getSheetByName(SHEET);
    if (!sheet) {
      say(false, 'sheet "' + SHEET + '"', 'not found - it is made on the first order');
    } else {
      var heading = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
      var mark = String.fromCharCode(1);
      var same = heading.join(mark) === HEADERS.join(mark);
      say(same, 'sheet heading', same ? HEADERS.length + ' columns, current'
        : 'out of date - it is rewritten on the next order, but rows written before ' +
          'the change will read crooked');
      say(true, 'orders on file', Math.max(0, sheet.getLastRow() - 1) + '');
    }
  } catch (err) { say(false, 'sheet', String(err)); }

  // 2. OneDrive, through Graph
  var drive = oneDriveTarget();
  if (!drive) {
    say(false, 'OneDrive', 'ONEDRIVE_TARGET is not set - no file can be stored. ' +
      'Expected user:someone@dema-solutions.com or site:<hostname>,<siteId>,<driveId>');
  } else if (!graphCredentials()) {
    say(false, 'OneDrive', 'GRAPH_TENANT / GRAPH_CLIENT / GRAPH_SECRET not all set, ' +
      'so Graph cannot be reached at all');
  } else {
    try {
      var probeName = 'selftest-' + Date.now();
      graphFolder([], ONEDRIVE_ROOT);
      var probe = graphFolder([ONEDRIVE_ROOT], probeName);
      graphUpload([ONEDRIVE_ROOT, probeName, 'selftest.txt'],
        Utilities.newBlob('ok', 'text/plain', 'selftest.txt'));
      graphDelete(probe.id);
      say(true, 'OneDrive', 'can write in "' + ONEDRIVE_ROOT + '" on ' + drive.label);
    } catch (err) { say(false, 'OneDrive', String(err)); }
  }

  // 2b. the languages the customer can be written to in
  try {
    var langs = mailLangs();
    say(langs.length > 0, 'mail languages', langs.join(' ') +
      ' - anyone asking for another language is written to in English');
  } catch (err) { say(false, 'mail languages', String(err)); }

  // 3. the back office
  try {
    var props = PropertiesService.getScriptProperties();
    var user = props.getProperty('ADMIN_USER');
    var pass = props.getProperty('ADMIN_PASS');
    say(!!(user && pass), 'back office login',
      user && pass ? 'ADMIN_USER and ADMIN_PASS are set'
        : 'not set - /quote/admin/ will refuse every login');
  } catch (err) { say(false, 'back office login', String(err)); }

  // 4. the sender
  var creds = graphCredentials();
  if (!creds) {
    say(false, 'Microsoft 365 sender',
      'GRAPH_TENANT / GRAPH_CLIENT / GRAPH_SECRET not all set - mail will go out ' +
      'from the Gmail account instead');
  } else {
    say(true, 'Microsoft 365 grant', creds.refresh
      ? 'delegated - the script acts as ' + SENDER_FROM
      : 'application - GRAPH_REFRESH is not set, so admin consent on ' +
        'Mail.Send and Files.ReadWrite.All is required');
    try {
      graphToken(creds);
      say(true, 'Microsoft 365 token', 'granted');
      graphSend(creds, {
        to: DESK, name: SENDER_NAME, replyTo: DESK,
        subject: 'Quote Desk self-test',
        htmlBody: shell('Self-test', '<p>This deployment can send as ' +
          escape(SENDER_FROM) + '. Nothing else was changed.</p>')
      });
      say(true, 'Microsoft 365 send', 'a test message is on its way to ' + DESK);
    } catch (err) {
      say(false, 'Microsoft 365', String(err) +
        ' - check that Mail.Send is an APPLICATION permission and that admin consent was granted');
    }
  }

  // 5. the deployed code, as the page will meet it
  say(typeof receiveFile === 'function' && typeof orderStatus === 'function' &&
      typeof deliverSelected === 'function',
    'endpoints', 'upload, status and delivery are present in this version');

  // 6. Trados, which does nothing at all until its properties are set
  say(typeof createTradosProject === 'function',
    'trados', tradosCredentials()
      ? 'configured; the menu will create projects'
      : 'not configured - the menu will say so and create nothing');

  var report = out.join(String.fromCharCode(10));
  Logger.log(report);
  try { SpreadsheetApp.getUi().alert('Quote Desk self-test', report,
    SpreadsheetApp.getUi().ButtonSet.OK); } catch (err) {}
  return report;
}


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quote Desk')
    .addItem('Create the Trados project for the selected order', 'createTradosProject')
    .addItem('Deliver the selected order', 'deliverSelected')
    .addSeparator()
    .addItem('Connect Microsoft 365', 'graphConnect')
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
  var out = null;
  folder.list().forEach(function (child) {
    if (child.folder && child.name === DELIVERY_FOLDER) out = child;
  });
  if (!out) {
    ui.alert('Put the finished files in a "' + DELIVERY_FOLDER + '" folder inside ' +
      id + ' first.' + String.fromCharCode(10, 10) + folder.getUrl());
    return;
  }

  /* Weighed before anything is fetched: the mail limit is the smallest limit in
     the chain, so a folder too heavy to send is also a folder there is no
     reason to download. */
  var listed = childrenIn(out).filter(function (child) { return !child.folder; });
  var bytes = 0;
  listed.forEach(function (child) { bytes += Number(child.size) || 0; });
  if (!listed.length) { ui.alert('That delivery folder is empty.'); return; }
  if (bytes > MAIL_ATTACHMENT_LIMIT) {
    ui.alert('Those files come to ' + Math.round(bytes / 1048576) + ' MB, too much for one ' +
      'email. Send them from OneDrive and set the status by hand.' +
      String.fromCharCode(10, 10) + folder.getUrl());
    return;
  }


  var answer = ui.alert('Send ' + listed.length + ' file(s) to ' + email + ' for ' + id + '?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  /* Fetched only now: a no at that prompt should not have cost a download. */
  var files = listed.map(function (child) {
    return fetchChild(child);
  });

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


/* ---------------------------------------------------------------- Trados */

/**
 * The order, as a project in Trados Team.
 *
 * Why it is a menu item and not something the order does by itself: today
 * nothing is charged at checkout and the status is moved by hand, so an order
 * placed to try the page out would otherwise mint a real project, burn real
 * machine translation and sit in the list for someone to clean up. The desk
 * decides; the code only does the typing.
 *
 * Each tier is its own project template, made once in the Trados UI, because
 * the API cannot set project settings or custom fields - only a template can.
 * The template is what carries the engine and the AI step; none of that is
 * configured from here.
 *
 * There is nothing to pay for this. API access carries no separate charge on
 * the plan: the only cost is the machine translation and AI the project
 * consumes, exactly as if the desk had built the project by hand in the
 * browser. The automation does not add to that bill, it only arrives sooner.
 *
 * Four things to put in Project Settings > Script properties, never in this
 * file. With any of them missing the menu item says so and does nothing:
 *
 *   TRADOS_CLIENT     the service user's Client ID
 *   TRADOS_SECRET     its Client Secret, shown once at creation and never again
 *   TRADOS_TENANT     the account id, from Profile > Manage Account >
 *                     Account Information > Trados Account ID
 *   TRADOS_LOCATION   the folder the projects are created in
 *
 * And one template id per tier: TRADOS_TEMPLATE_BASIC,
 * TRADOS_TEMPLATE_GOOD_ENOUGH, TRADOS_TEMPLATE_PROFESSIONAL.
 */
var TRADOS_API = 'https://api.eu.cloud.trados.com/public-api/v1';
var TRADOS_TOKEN_URL = 'https://sdl-prod.eu.auth0.com/oauth/token';
/* Not api.trados.com: that is the audience the browser UI asks for, and a
 * token minted against it is refused by the public API. */
var TRADOS_AUDIENCE = 'https://api.sdl.com';
var TRADOS_TEMPLATE_PROPS = {
  'basic': 'TRADOS_TEMPLATE_BASIC',
  'good-enough': 'TRADOS_TEMPLATE_GOOD_ENOUGH',
  'professional': 'TRADOS_TEMPLATE_PROFESSIONAL'
};
/* A project of thirty large files would run past the six minutes a menu item
 * is given, and Apps Script kills it mid-upload with no way to know which
 * files arrived. Until there is a queue behind a timed trigger, the desk is
 * told to add the rest in the browser. */
var TRADOS_MAX_FILES = 10;
var TRADOS_MAX_BYTES = 40 * 1024 * 1024;

function tradosCredentials() {
  try {
    var props = PropertiesService.getScriptProperties();
    var creds = {
      client: props.getProperty('TRADOS_CLIENT'),
      secret: props.getProperty('TRADOS_SECRET'),
      tenant: props.getProperty('TRADOS_TENANT'),
      location: props.getProperty('TRADOS_LOCATION')
    };
    if (creds.client && creds.secret && creds.tenant) return creds;
  } catch (err) {}
  return null;
}

function tradosTemplate(solution) {
  var name = TRADOS_TEMPLATE_PROPS[String(solution)];
  if (!name) return '';
  try {
    return PropertiesService.getScriptProperties().getProperty(name) || '';
  } catch (err) {
    return '';
  }
}

/**
 * A token, kept for the day it lasts.
 *
 * This is the tightest limit in the whole integration and it is not a rate
 * limit on the API: Auth0 grants this application sixteen tokens a day and
 * then starts refusing the address. A token lives 24 hours, so one a day is
 * all that is needed - but only if it is kept somewhere that lasts. It is kept
 * in the script properties and not in CacheService, which holds nothing beyond
 * six hours and would quietly ask for four tokens a day, then more on a busy
 * one.
 *
 * It is stored in clear, like the secret beside it; anyone who can open the
 * script can read both, so the service user should be given rights over the
 * one location it needs and no more.
 */
function tradosToken(creds) {
  var props = PropertiesService.getScriptProperties();
  var kept = props.getProperty('TRADOS_TOKEN');
  var until = Number(props.getProperty('TRADOS_TOKEN_UNTIL') || 0);
  /* An hour of margin, so a token never expires halfway through a project. */
  if (kept && until > Date.now() + 3600000) return kept;

  var res = UrlFetchApp.fetch(TRADOS_TOKEN_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      client_id: creds.client,
      client_secret: creds.secret,
      grant_type: 'client_credentials',
      audience: TRADOS_AUDIENCE
    }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Trados token HTTP ' + res.getResponseCode() + ' ' +
      res.getContentText().slice(0, 200));
  }
  var body = JSON.parse(res.getContentText());
  props.setProperty('TRADOS_TOKEN', body.access_token);
  props.setProperty('TRADOS_TOKEN_UNTIL',
    String(Date.now() + (Number(body.expires_in) || 86400) * 1000));
  return body.access_token;
}

/**
 * One call to the public API.
 *
 * The tenant header is X-LC-Tenant, with hyphens throughout: written
 * X-LC_Tenant it answers 401 and says nothing about why.
 */
function tradosCall(creds, method, path, payload) {
  var options = {
    method: method,
    headers: {
      Authorization: 'Bearer ' + tradosToken(creds),
      'X-LC-Tenant': creds.tenant
    },
    muteHttpExceptions: true
  };
  if (payload !== undefined && payload !== null) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  var res = UrlFetchApp.fetch(TRADOS_API + path, options);
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Trados ' + method.toUpperCase() + ' ' + path +
      ' HTTP ' + code + ' ' + text.slice(0, 300));
  }
  return text ? JSON.parse(text) : {};
}

/**
 * The language pair, as Trados spells it.
 *
 * Read from the two code columns the page fills. A row written before those
 * columns existed falls back to LOCALES, and a name LOCALES does not know is
 * dropped rather than guessed at: a project created into the wrong language is
 * worse than one that was never created at all.
 */
function tradosDirections(at) {
  var from = String(at('From code') || localeFor(String(at('From'))) || '').trim();
  var into = String(at('Into codes') || '').split(',')
    .map(function (code) { return code.trim(); })
    .filter(function (code) { return code; });
  if (!into.length) {
    into = String(at('Into') || '').split(',')
      .map(function (name) { return localeFor(name.trim()); })
      .filter(function (code) { return code; });
  }
  if (!from || !into.length) return null;
  return into.map(function (code) {
    return {
      sourceLanguage: { languageCode: from },
      targetLanguage: { languageCode: code }
    };
  });
}

/**
 * Hand one file to a project.
 *
 * The multipart field names are the one part of this the documentation does
 * not write down - it says only "add the file to the request body". These are
 * the names the API answers to; if a release ever changes them the symptom is
 * a 400 with no detail in it, and this is the place to look.
 *
 * A ZIP becomes a reference file whatever is asked of it, so it is sent as one
 * openly rather than being counted as words that will never be translated.
 */
function tradosAddFile(creds, projectId, blob, language) {
  var name = blob.getName();
  var zip = /\.zip$/i.test(name);
  var res = UrlFetchApp.fetch(
    TRADOS_API + '/projects/' + encodeURIComponent(projectId) + '/source-files', {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + tradosToken(creds),
        'X-LC-Tenant': creds.tenant
      },
      payload: {
        file: blob,
        properties: Utilities.newBlob(JSON.stringify({
          name: name,
          role: zip ? 'reference' : 'translatable',
          type: 'native',
          language: language
        }), 'application/json', 'properties')
      },
      muteHttpExceptions: true
    });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Trados upload ' + name + ' HTTP ' + code + ' ' +
      res.getContentText().slice(0, 300));
  }
}

/**
 * Build the Trados project for the order the desk has selected.
 *
 * The order of the three calls matters and is not negotiable: the project is
 * created, then its files are added, then it is started. A project started
 * before its files have arrived has nothing to work on, and a project that is
 * never started stays a draft no workflow will touch.
 */
function createTradosProject() {
  var ui = SpreadsheetApp.getUi();
  var creds = tradosCredentials();
  if (!creds) {
    ui.alert('Trados is not configured. Put TRADOS_CLIENT, TRADOS_SECRET and ' +
      'TRADOS_TENANT in Project Settings > Script properties first.');
    return;
  }

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
  if (!id) { ui.alert('That row carries no reference.'); return; }

  /* Only a quote the customer has accepted goes into production. A declined or
     unanswered quote put into Trados is work nobody has agreed to pay for, and
     the row says plainly which it is. */
  var status = String(at('Status') || '');
  if (status !== 'Accepted' && status !== 'Delivered') {
    ui.alert('That row is "' + (status || 'blank') + '", not "Accepted". ' +
      'Only a quote the customer has accepted is sent to Trados. If they accepted ' +
      'by telephone, set the status by hand first - deliberately.');
    return;
  }

  if (String(at('Trados project'))) {
    ui.alert(id + ' already has the Trados project ' + at('Trados project') +
      '. Clear that cell if you really mean to create a second one.');
    return;
  }

  var solution = String(at('Solution'));
  var template = tradosTemplate(solution);
  if (!template) {
    ui.alert('No project template is set for the "' + solution + '" tier. Put its id in ' +
      (TRADOS_TEMPLATE_PROPS[solution] || 'a script property') + ' under Script properties.');
    return;
  }

  var directions = tradosDirections(at);
  if (!directions) {
    ui.alert('That row carries no language codes I can trust, and I will not guess them. ' +
      'Fill "From code" and "Into codes" by hand, or build this one in Trados.');
    return;
  }

  var folder = orderFolder(id, '');
  /* Listed and weighed before anything is fetched: UrlFetchApp will not carry
     much over 50 MB, and a folder too heavy to send is also a folder there is
     no point downloading one file at a time to find that out. */
  var listed = folder.list().filter(function (child) {
    if (child.folder) return false;
    /* The paperwork the customer was sent lives in the same folder and is not
       part of the job. */
    return !/^(quote|order-confirmation|confirmation)/i.test(child.name);
  });
  var bytes = 0;
  listed.forEach(function (child) { bytes += Number(child.size) || 0; });
  if (!listed.length) {
    ui.alert('There are no customer files in ' + id + '.' +
      String.fromCharCode(10, 10) + folder.getUrl());
    return;
  }
  if (listed.length > TRADOS_MAX_FILES || bytes > TRADOS_MAX_BYTES) {
    ui.alert(listed.length + ' file(s), ' + Math.round(bytes / 1048576) + ' MB. That is more ' +
      'than one menu run can fetch from OneDrive and upload inside the six minutes Apps ' +
      'Script allows - and a single file over about 50 MB cannot pass through here at all. ' +
      'Create this project in Trados and add the files there, from OneDrive:' +
      String.fromCharCode(10, 10) + folder.getUrl());
    return;
  }

  var answer = ui.alert('Create a ' + solution + ' project in Trados for ' + id + '? ' +
    directions.length + ' language direction(s), ' + listed.length + ' file(s).',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  /* Fetched only now, for the same reason as in the delivery menu: a no at that
     prompt should not have cost a download. */
  var files = listed.map(function (child) {
    return fetchChild(child);
  });

  var body = {
    name: id + ' - ' + (String(at('Project')) || solution),
    description: 'Quote Desk order ' + id + ', ' + solution,
    projectTemplate: { id: template },
    languageDirections: directions
  };
  var due = at('Delivery');
  if (due instanceof Date) body.dueBy = due.toISOString();
  /* Without a location the API tries the root folder and answers forbidden. */
  if (creds.location) body.location = creds.location;

  var project = tradosCall(creds, 'post', '/projects', body);
  var projectId = String(project.id || '');
  if (!projectId) {
    throw new Error('Trados created something with no id: ' +
      JSON.stringify(project).slice(0, 200));
  }

  /*
   * From here the project exists, and a failure must not read as though
   * nothing had happened. The id is written to the row before the files go up,
   * so a crash halfway through leaves a project the desk can find and finish
   * by hand rather than an orphan nobody knows about.
   */
  sheet.getRange(rowIndex, HEADERS.indexOf('Trados project') + 1).setValue(projectId);
  SpreadsheetApp.flush();

  var source = directions[0].sourceLanguage.languageCode;
  files.forEach(function (blob) { tradosAddFile(creds, projectId, blob, source); });
  tradosCall(creds, 'put', '/projects/' + encodeURIComponent(projectId) + '/start', null);

  sheet.getRange(rowIndex, HEADERS.indexOf('Status') + 1).setValue('In Trados');
  ui.alert('Trados project ' + projectId + ' created and started for ' + id + '.');
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
/**
 * May this request still be given somewhere to put a file?
 *
 * The three conditions every upload answers to, in one place so the old path
 * and the new one cannot drift apart:
 *
 *  - the reference and the email must agree, as they must for a status lookup.
 *    The reference alone is four random characters on a known date, which is
 *    guessable, and the prize for guessing would be write access to a folder.
 *  - the request must be younger than UPLOAD_WINDOW_HOURS.
 *  - there must be room: MAX_FILES and no more.
 *
 * A refusal says as little as the status lookup does, for the same reason.
 */
function uploadAllowed(req) {
  var id = String((req || {}).id || '');
  if (!/^[A-Za-z0-9-]{4,40}$/.test(id)) return { ok: false, error: 'bad reference' };

  var row = findOrderRow(id);
  if (!row) return { ok: false, error: 'not found' };

  var given = String((req || {}).email || '').trim().toLowerCase();
  var want = String(row[HEADERS.indexOf('Email')] || '').trim().toLowerCase();
  if (!given || given !== want) return { ok: false, error: 'not found' };

  var received = Object.prototype.toString.call(row[0]) === '[object Date]' ? row[0].getTime() : 0;
  if (!received || Date.now() - received > UPLOAD_WINDOW_HOURS * 3600 * 1000) {
    return { ok: false, error: 'order closed to uploads' };
  }
  return { ok: true, id: id, row: row };
}

/** How many files are in a request's folder already. */
function uploadCount(folder) {
  try {
    return folder.list().length;
  } catch (err) {
    /* A folder that cannot be listed must not be filled blindly: that is how
       thirty turns into three hundred. */
    throw new Error('cannot list the folder of this order: ' + err);
  }
}

/**
 * The permission the browser uses to put one file on OneDrive by itself.
 *
 * The bytes never come through here. What goes back is an upload URL, good for
 * one file, for an hour, in one folder - and nothing else: never the token the
 * app holds, which would be good for every drive in the tenant.
 */
function uploadTicket(req) {
  var allowed = uploadAllowed(req);
  if (!allowed.ok) return allowed;

  var size = Number((req || {}).size) || 0;
  if (size <= 0) return { ok: false, error: 'empty file' };
  if (size > MAX_FILE_BYTES) return { ok: false, error: 'file too large' };

  var name = fileSafe((req || {}).name);
  /* Any format a customer sends is translatable work - ZIP above all, which is
     how a whole job arrives. The handful refused are the ones that are not
     documents in any reading: this endpoint is open by design, and a stranger
     should not be able to put something runnable on the desk's own drive. */
  if (RUNNABLE.test(name)) return { ok: false, error: 'file type not accepted' };
  var folder = orderFolder(allowed.id, '');
  if (uploadCount(folder) >= MAX_FILES) return { ok: false, error: 'too many files' };

  /*
   * An upload session is a OneDrive thing. When the fallback has put this
   * request on Google Drive there is no URL to hand the browser, so the
   * browser is told to send the bytes through here instead, the old way:
   * base64, one file per request, action 'upload'. It is slower and it is
   * bounded by what a single Apps Script request will hold, which is why it
   * is the fallback and not the road - but a customer who cannot upload at
   * all is worse than a customer who uploads slowly.
   */
  if (folder.isDrive) {
    return { ok: true, name: name, relay: true, limit: RELAY_MAX_BYTES };
  }

  var session = graphUploadSession(folder.getPath().concat([name]), name, size);
  return { ok: true, name: name, uploadUrl: session.uploadUrl,
    chunk: session.chunk, expires: session.expires };
}

/**
 * The browser's word that a file is up.
 *
 * Nothing here is trusted as proof - the file is on OneDrive or it is not, and
 * the folder says which. It is the moment the row learns where its folder is,
 * which it needs whether one file arrived or twelve did.
 */
function uploadDone(req) {
  var allowed = uploadAllowed(req);
  if (!allowed.ok) return allowed;
  var folder = orderFolder(allowed.id, '');
  rememberFolder(allowed.id, folder);

  var listed = [];
  try {
    listed = folder.list().filter(function (child) { return !child.folder; });
  } catch (err) {}

  /* The desk is told the moment the files are actually there, which is not the
     moment the quote went out: a 2 GB archive is still climbing for an hour
     after the customer has closed the tab. Said once, after the last file, with
     what arrived and where it is. */
  try { tellDeskFilesArrived(allowed.row, folder, listed); } catch (err) {}

  return { ok: true, files: listed.length };
}

/** 2411724 -> 2.3 MB. A size a person reads, not a number to count zeros in. */
function readableSize(bytes) {
  var n = Number(bytes) || 0;
  if (n >= 1024 * 1024 * 1024) return (n / (1024 * 1024 * 1024)).toFixed(2) + ' GB';
  if (n >= 1024 * 1024) return (n / (1024 * 1024)).toFixed(1) + ' MB';
  if (n >= 1024) return Math.round(n / 1024) + ' KB';
  return n + ' B';
}

/**
 * "The customer has uploaded their files, and here is where they are."
 *
 * One message to the desk, with the folder link and what is in it. The link is
 * a OneDrive folder on the desk's own drive, so it opens for whoever is signed
 * in and needs no sharing link to be minted - and no link to the files is ever
 * put anywhere a customer could reach.
 */
function tellDeskFilesArrived(row, folder, listed) {
  var at = function (name) { return row[HEADERS.indexOf(name)]; };
  var who = [at('First name'), at('Last name')].filter(String).join(' ');
  var company = String(at('Company') || '');
  var bytes = 0;
  listed.forEach(function (child) { bytes += Number(child.size) || 0; });

  var rows = listed.map(function (child) {
    return [child.name, readableSize(child.size)];
  });
  if (!rows.length) rows.push(['(nothing in the folder)', '-']);

  send({
    to: DESK,
    name: SENDER_NAME,
    replyTo: String(at('Email') || ''),
    subject: 'Files uploaded - ' + (company || who || at('Reference')) + ' - ' +
      listed.length + ' file(s), ' + readableSize(bytes),
    htmlBody: shell('The customer has uploaded their files',
      table([
        ['Customer', company ? company + ' (' + who + ')' : who],
        ['Email', at('Email')],
        ['Quote number', at('Quote number') || '-'],
        ['Reference', at('Reference')],
        ['Project', at('Project') || '-'],
        ['Requested', at('Received') ? when(new Date(at('Received')).toISOString()) : '-'],
        ['Files', listed.length + ', ' + readableSize(bytes)]
      ]) +
      '<p style="margin:22px 0"><a href="' + folder.getUrl() + '" ' +
      'style="background:#1b2a6b;color:#fff;text-decoration:none;padding:12px 20px;' +
      'border-radius:6px;font-weight:700">Open the folder on OneDrive</a></p>' +
      table(rows) +
      '<p>Nothing is agreed yet: this says the documents are here, not that the ' +
      'quote has been accepted.</p>')
  });
}

/** The folder's link and id, written into the row that owns it. */
/**
 * Where a reference sits in the sheet, counting from 1 as the sheet does.
 *
 * Searched from the bottom: the row wanted is nearly always the newest, and a
 * reference that somehow appears twice should resolve to the live one.
 *
 * @returns {{sheet: Object, row: number}|null}
 */
function rowIndexFor(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sheet) return null;
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var refs = sheet.getRange(2, HEADERS.indexOf('Reference') + 1, last - 1, 1).getValues();
  for (var i = refs.length - 1; i >= 0; i--) {
    if (String(refs[i][0]) === String(id)) return { sheet: sheet, row: i + 2 };
  }
  return null;
}

function rememberFolder(id, folder) {
  var at = rowIndexFor(id);
  if (!at) return;
  at.sheet.getRange(at.row, HEADERS.indexOf('Files') + 1).setValue(folder.getUrl());
  at.sheet.getRange(at.row, HEADERS.indexOf('Drive item') + 1).setValue(folder.getId());
}

/**
 * What went wrong after the request was recorded, written beside it.
 *
 * Added to rather than replaced: a second attempt that fails differently
 * should not erase the first reason. It is the back office that reads this,
 * so it is written in sentences and not in codes.
 */
function noteTrouble(id, text) {
  var at = rowIndexFor(id);
  if (!at) return;
  var cell = at.sheet.getRange(at.row, HEADERS.indexOf('Trouble') + 1);
  var had = String(cell.getValue() || '').trim();
  cell.setValue(had ? had + ' ' + text : text);
}

/**
 * The older upload: the whole file, base64, in one request.
 *
 * It is kept because the page now live still uses it, and a redeploy must not
 * leave production unable to receive a document for as long as it takes to
 * build and push the new page. It writes to OneDrive like everything else, so
 * there is one destination and not two.
 *
 * Delete it once the page that sends action 'upload' is gone - the chunked path
 * above has no size limit worth naming, and this one cannot carry more than a
 * single Apps Script request will hold.
 */
function receiveFile(req) {
  /* The old page sends no email with the upload, so this path cannot ask for
     one. It keeps the window and the ceiling, and it is on its way out. */
  var id = String(req.id || '');
  if (!/^[A-Za-z0-9-]{4,40}$/.test(id)) return { ok: false, error: 'bad reference' };

  var row = findOrderRow(id);
  if (!row) return { ok: false, error: 'unknown order' };

  /*
   * The page now live sends its email with the upload, because this path is
   * no longer only the old one: it is also the relay the fallback uses when
   * the files are going to Google Drive. When an email is sent it has to
   * match the row, exactly as a status lookup does - the reference on its own
   * is four random characters on a known date. An upload that carries no
   * email is the old page, and keeps the old rule.
   */
  if (req.email) {
    var owner = String(row[HEADERS.indexOf('Email')] || '').trim().toLowerCase();
    if (owner !== String(req.email).trim().toLowerCase()) {
      return { ok: false, error: 'unknown order' };
    }
  }

  var received = Object.prototype.toString.call(row[0]) === '[object Date]' ? row[0].getTime() : 0;
  if (!received || Date.now() - received > UPLOAD_WINDOW_HOURS * 3600 * 1000) {
    return { ok: false, error: 'order closed to uploads' };
  }

  var name = fileSafe(req.name);
  var data = String(req.data || '');
  if (!data) return { ok: false, error: 'empty file' };
  /* What a single Apps Script request can hold, not what OneDrive can: this
     path is the limited one, which is why the chunked path exists. */
  if (data.length > RELAY_MAX_BYTES * 1.4) return { ok: false, error: 'file too large' };

  var folder = orderFolder(id, '');
  if (uploadCount(folder) >= MAX_FILES) return { ok: false, error: 'too many files' };

  var blob = Utilities.newBlob(Utilities.base64Decode(data),
    req.mime || 'application/octet-stream', name);
  folder.put(blob, name);
  rememberFolder(id, folder);
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
/*
 * Status, now that the page quotes rather than orders:
 *
 *   Quoted     the price has gone out and nobody has answered
 *   Accepted   the customer pressed the button; the work is on
 *   Declined   the customer said no; nothing is done and the files stay put
 *   Expired    nobody answered within QUOTE_VALID_DAYS
 *   Delivered  written by the delivery menu when the translation goes back
 *
 * The first four are written by this script. 'Delivered' and anything the desk
 * invents in between are written by hand, as they always were.
 */
var STATUS_NEW = 'Quoted';

/*
 * Everything the order carried, in one row - who, where, what, when, how much.
 * The billing address used to be dropped on the floor: the row kept the company
 * and the country and threw away the street, the town, the postcode and the
 * telephone, which is most of an address. An order you cannot invoice is not a
 * record of anything.
 */
var HEADERS = ['Received', 'Reference', 'Status', 'Customer code',
  'First name', 'Last name', 'Email', 'Telephone',
  'Company', 'VAT number', 'Tax code', 'Address', 'City', 'Postal code', 'Country',
  'Project', 'Solution', 'From', 'Into', 'Words', 'Subject', 'Desktop publishing', 'Delivery',
  'Net', 'VAT', 'Total', 'Currency', 'VAT treatment', 'Re-priced',
  'Customer reference', 'Notes', 'Files',
  'From code', 'Into codes', 'Trados project', 'Quote number',
  /* Added with the move to quoting. New columns go on the end and nowhere else:
     the heading repairs itself, but a column inserted in the middle would make
     every row written before it read one cell to the left. */
  'Language', 'Decision token', 'Decided', 'Drive item',
  /* Whatever went wrong after the row was written: a folder Graph refused, a
     PDF that would not build, a mail that did not leave. The request is in the
     book either way - this says what still needs a hand. */
  'Trouble'];

/*
 * The language pair, as a CAT tool spells it.
 *
 * `From` and `Into` hold names, because a person reads this sheet. Trados reads
 * locale codes, and the two must not be guessed at from one another at the
 * moment a project is created: that is the one moment there is nobody to
 * correct a wrong guess.
 *
 * So the page now sends the codes it already had, and they are kept in their
 * own two columns - appended at the end, never inserted, so that every row
 * written before today still lines up with its heading.
 *
 * LOCALES is the way back for those earlier rows. It is the name side of
 * LANGUAGES in src/config.js and must stay equal to it; the test asserts that,
 * so the table cannot drift in silence.
 */
var LOCALES = {
  'English (United Kingdom)': 'en-GB', 'English (United States)': 'en-US',
  'French (France)': 'fr-FR', 'Italian (Italy)': 'it-IT',
  'Spanish (Spain)': 'es-ES', 'Spanish (Latin America)': 'es-419',
  'Portuguese (Brazil)': 'pt-BR', 'Portuguese (Portugal)': 'pt-PT',
  'German (Germany)': 'de-DE', 'Dutch (Netherlands)': 'nl-NL',
  'Flemish (Belgium)': 'nl-BE', 'Basque (Spain)': 'eu-ES',
  'Romanian (Romania)': 'ro-RO', 'Russian (Russia)': 'ru-RU',
  'Polish (Poland)': 'pl-PL', 'Czech (Czechia)': 'cs-CZ',
  'Ukrainian (Ukraine)': 'uk-UA', 'Turkish (Turkey)': 'tr-TR',
  'Slovak (Slovakia)': 'sk-SK', 'Bulgarian (Bulgaria)': 'bg-BG',
  'Croatian (Croatia)': 'hr-HR', 'Serbian (Serbia)': 'sr-RS',
  'Bosnian (Bosnia and Herzegovina)': 'bs-BA', 'Albanian (Albania)': 'sq-AL',
  'Hungarian (Hungary)': 'hu-HU', 'Greek (Greece)': 'el-GR',
  'Slovenian (Slovenia)': 'sl-SI', 'Estonian (Estonia)': 'et-EE',
  'Latvian (Latvia)': 'lv-LV', 'Lithuanian (Lithuania)': 'lt-LT',
  'Georgian (Georgia)': 'ka-GE', 'Danish (Denmark)': 'da-DK',
  'Swedish (Sweden)': 'sv-SE', 'Norwegian (Norway)': 'nb-NO',
  'Finnish (Finland)': 'fi-FI', 'Icelandic (Iceland)': 'is-IS',
  'Hindi (India)': 'hi-IN', 'Vietnamese (Vietnam)': 'vi-VN',
  'Thai (Thailand)': 'th-TH', 'Maltese (Malta)': 'mt-MT',
  'Persian (Iran)': 'fa-IR', 'Hebrew (Israel)': 'he-IL',
  'Swahili (Kenya)': 'sw-KE', 'Korean (South Korea)': 'ko-KR',
  'Japanese (Japan)': 'ja-JP', 'Chinese (Simplified)': 'zh-CN',
  'Chinese (Traditional)': 'zh-TW', 'Arabic (Saudi Arabia)': 'ar-SA'
};

/**
 * The locale code for one language, however the row happens to name it.
 *
 * A code goes through untouched, so a row that already carries codes is not
 * translated twice. A name is looked up. Anything else answers '' rather than
 * a plausible guess: an empty cell stops the project being created and asks a
 * human, which is the right outcome. A wrong locale would not.
 */
function localeFor(language) {
  var text = String(language || '').trim();
  if (!text) return '';
  if (LOCALES[text]) return LOCALES[text];
  return /^[a-z]{2,3}(-[A-Za-z0-9]{2,4})?$/.test(text) ? text : '';
}

/** Every locale in a cell that may hold one code or a list of them. */
function localesFrom(cell) {
  return String(cell || '').split(',')
    .map(function (x) { return localeFor(x); })
    .filter(function (x) { return x; });
}

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

function record(order, net, vat, gross, repriced, folderUrl, quote, extra) {
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
    b.company ? b.companyName : 'individual', b.vatNumber || '', b.taxCode || '',
    street, b.city || '', b.postalCode || '', b.countryName || b.country || '',
    order.project || '', order.solution, order.source, (order.targets || []).join(', '),
    (Number(order.words) || 0) * (order.targets || []).length, order.subject,
    order.layout ? order.layout.pages + ' pages' : 'no',
    order.eta ? new Date(order.eta) : '', net, vat, gross, order.currency,
    order.vatTreatment || '', repriced ? 'yes' : '', order.customer.reference || '',
    order.customer.note || '', folderUrl || '',
    // The page sends these; the fallback is only for a page older than the
    // columns, and answers '' rather than a guess when it cannot tell.
    order.sourceCode || localeFor(order.source),
    (order.targetCodes && order.targetCodes.length
      ? order.targetCodes.map(localeFor)
      : (order.targets || []).map(localeFor)).join(', '),
    // The Trados project is not created at checkout - the desk creates it from
    // the menu, and writes the id here when it does.
    '',
    quote || '',
    (extra || {}).language || 'en',
    (extra || {}).token || '',
    '',                                  // Decided: written when the customer answers
    (extra || {}).item || '',
    (extra || {}).trouble || ''
  ]);
}

/**
 * The order book, newest first, as rows in the order of HEADERS.
 *
 * Dates come back as ISO strings so the back office can format them in the
 * reader's own locale; everything else is handed over as the sheet holds it.
 */
/*
 * A row, read back for the desk.
 *
 * The three actions below all begin the same way: find the request, and
 * refuse the same way whether it is missing or the reference is nonsense.
 */
function deskRow(id) {
  if (!/^[A-Za-z0-9-]{4,40}$/.test(String(id || ''))) return null;
  var row = findOrderRow(String(id));
  if (!row) return null;
  return function (name) { return row[HEADERS.indexOf(name)]; };
}

/** The money of a row, as the documents want it. */
function deskMoney(at) {
  var net = Number(at('Net')) || 0;
  var vat = Number(at('VAT')) || 0;
  return { net: net, vat: vat, gross: Number(at('Total')) || round2(net + vat) };
}

/*
 * The quote, rebuilt.
 *
 * Nothing is stored and nothing needs to be: the PDF is drawn from the row,
 * so it can be produced today, after a refusal from Graph, or in a year. That
 * is what makes a request traceable even when the folder was never made - and
 * on 14 November 2026 that was every request, because Graph was refusing and
 * the only copy of the document was an attachment in somebody's mail.
 *
 * It is handed back as base64 because this is a web app answering JSON, not a
 * file server.
 */
function deskQuotePdf(req) {
  var at = deskRow((req || {}).id);
  if (!at) return { ok: false, error: 'not found' };

  var money = deskMoney(at);
  var quote = at('Quote number') || at('Reference');
  var language = mailLang(at('Language'));
  var decided = String(at('Status') || '') === 'Accepted';
  var pdf = confirmationPdf(orderFromRow(at), money.net, money.vat, money.gross,
    language, decided ? 'order' : 'quote', quote);

  return { ok: true, name: pdf.getName(), mime: 'application/pdf',
    data: Utilities.base64Encode(pdf.getBytes()) };
}

/*
 * The quote, sent again.
 *
 * Same document, same two links, same language: this is not a new quote and
 * it must not read as one, because the customer may already be holding the
 * first copy. It is for the case where the first never arrived - which is
 * common enough while the mail is leaving from a Gmail account rather than
 * from the desk's own, as it does until an administrator consents.
 *
 * A quote already answered is not sent again. Reopening a decision by email
 * is a conversation with the desk, not a button.
 */
function deskResend(req) {
  var at = deskRow((req || {}).id);
  if (!at) return { ok: false, error: 'not found' };

  var status = String(at('Status') || '');
  if (status === 'Accepted' || status === 'Declined') {
    return { ok: false, error: 'already decided', status: status };
  }

  var id = String(at('Reference'));
  var token = String(at('Decision token') || '');
  if (!token) return { ok: false, error: 'no decision token on this row' };

  var money = deskMoney(at);
  var quote = at('Quote number') || id;
  var language = mailLang(at('Language'));
  var order = orderFromRow(at);
  var pdf = confirmationPdf(order, money.net, money.vat, money.gross,
    language, 'quote', quote);

  send({
    to: at('Email'), name: SENDER_NAME, replyTo: DESK,
    subject: mail('quote_subject', language, { quote: quote, id: id }),
    htmlBody: quoteBody(order, money.net, money.vat, money.gross, language, quote, token),
    attachments: [pdf]
  });

  return { ok: true, sent: String(at('Email')) };
}

/*
 * The folder, tried again.
 *
 * For the backlog: every request whose folder Graph refused is still sitting
 * in the book with its `Trouble` filled in, and the day consent arrives they
 * all want making. This remakes the folder, files the quote PDF beside it,
 * and writes down where it ended up.
 *
 * It says plainly which drive it landed on, because a folder made on the
 * fallback is not the same answer as a folder made on OneDrive, and the desk
 * should not have to read a URL to tell them apart.
 */
function deskRetryFiles(req) {
  var at = deskRow((req || {}).id);
  if (!at) return { ok: false, error: 'not found' };

  var id = String(at('Reference'));
  var folder;
  try {
    folder = orderFolder(id, '');
  } catch (err) {
    return { ok: false, error: String(err) };
  }
  rememberFolder(id, folder);

  var filed = false;
  try {
    var money = deskMoney(at);
    var quote = at('Quote number') || id;
    var language = mailLang(at('Language'));
    var kind = String(at('Status') || '') === 'Accepted' ? 'order' : 'quote';
    var wanted = confirmationPdf(orderFromRow(at), money.net, money.vat, money.gross,
      language, kind, quote);
    var already = folder.list().some(function (child) {
      return !child.folder && child.name === wanted.getName();
    });
    if (!already) folder.put(wanted);
    filed = true;
  } catch (err) {
    noteTrouble(id, 'A second attempt made the folder but could not file the quote (' +
      err + ').');
  }

  /* The reason is cleared only when there is nothing left to say. Half a
     recovery that reads as a whole one is how a missing file goes unnoticed. */
  if (folder.isDrive) {
    noteTrouble(id, 'Tried again on ' + when(new Date().toISOString()) +
      ': Microsoft still refuses, so the folder is on Google Drive.');
  } else if (filed) {
    var where = rowIndexFor(id);
    if (where) where.sheet.getRange(where.row, HEADERS.indexOf('Trouble') + 1).setValue('');
  }

  return { ok: true, drive: !!folder.isDrive, filed: filed, url: folder.getUrl() };
}

function orders() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  /* The decision token is the one cell the back office has no business
     receiving: it is the key to accepting a quote on the customer's behalf, and
     a page that holds it puts it in a browser, a cache and a screenshot. The
     column stays in HEADERS so the heading and the rows line up; the value is
     replaced by a mark that says a token is there. */
  var secret = HEADERS.indexOf('Decision token');
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    out.push(rows[i].map(function (cell, column) {
      if (column === secret) return cell ? '(set)' : '';
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

/** The quote, which is the same document saying a different thing. */
function quotePdf(order, net, vat, gross, lang, quote) {
  return confirmationPdf(order, net, vat, gross, lang, 'quote', quote);
}

/**
 * The paperwork, on the shape of a Dema Solutions invoice.
 *
 * One document, two states. As a quote it carries the number, the figures and
 * the day it lapses; as an order confirmation it carries the same figures and
 * says the work is agreed. Neither is a fiscal invoice and both say so: the
 * invoice goes out through the SdI from accounting.
 *
 * `lang` is the customer's own language. The reverse-charge wording is not
 * translated - it is the legal formula, and a translated formula is a different
 * formula.
 */
function confirmationPdf(order, net, vat, gross, lang, kind, quote) {
  var navy = '#1b2a6b';
  var language = mailLang(lang);
  var isQuote = kind === 'quote';
  var docTitle = mail(isQuote ? 'pdf_doc_quote' : 'pdf_doc_order', language);
  var docName = isQuote ? mail('pdf_doc_quote', language) : mail('pdf_doc_order', language);
  var number = quote || order.id;
  var count = (order.targets || []).length;
  var block = recipientBlock(order).map(function (line, i) {
    return '<div style="' + (i === 0 ? 'font-weight:700;color:#111' : 'color:#6e768c') + '">' +
      escape(line) + '</div>';
  }).join('');

  var rows = '<tr style="background:#f0f2f7">' +
    '<td style="padding:10px 8px"><strong>' + escape(mail('pdf_services', language)) + ' ' +
    escape(order.id) + '</strong>' +
    '<div style="color:#5c7a6a">' + escape(order.source) + ' &gt; ' + escape((order.targets || []).join(', ')) +
    ' - ' + thousands((Number(order.words) || 0) * count) + ' ' + escape(mail('pdf_words', language)) +
    ' - ' + escape(order.solution) + '</div></td>' +
    '<td style="padding:10px 8px;text-align:right;color:#6e768c">' + count + ' ' +
    escape(mail(count === 1 ? 'pdf_language' : 'pdf_languages', language)) + '</td>' +
    '<td style="padding:10px 8px;text-align:right">' + money(net - layoutTotal(order), order.currency) +
    '</td></tr>';
  if (layoutTotal(order)) {
    rows += '<tr><td style="padding:10px 8px"><strong>' +
      escape(mail('pdf_dtp', language)) + '</strong>' +
      '<div style="color:#5c7a6a">' +
      escape(mail('pdf_dtp_note', language, { n: order.layout.pages })) + '</div></td>' +
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
    '<strong style="font-size:12pt">' + escape(docTitle) + '</strong><br>' +
    escape(mail('pdf_no', language)) + ' ' + escape(number) +
    ' ' + escape(mail('pdf_of', language)) + ' ' +
    Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd') +
    (isQuote ? '<br><span style="font-size:8pt">' +
      escape(mail('pdf_valid', language,
        { date: Utilities.formatDate(new Date(quoteExpiry(new Date())), 'Europe/Rome', 'yyyy-MM-dd') })) +
      '</span>' : '') + '</div>' +
    '<hr style="border:none;border-top:1px solid ' + navy + ';margin:16px 0">' +
    '<table style="width:100%;font-size:9pt"><tr style="vertical-align:top">' +
    '<td style="width:48%"><strong style="color:' + navy + '">' +
    escape(mail('pdf_vatid', language)) + '</strong> ' +
    escape((order.billing || {}).vatNumber || (order.billing || {}).countryName || '-') +
    '<br><strong style="color:' + navy + '">' + escape(mail('pdf_email', language)) +
    '</strong> ' + escape(order.customer.email) + '</td>' +
    '<td><div style="color:' + navy + ';font-weight:700;font-size:8pt">' +
    escape(mail('pdf_recipient', language)) + '</div>' + block + '</td>' +
    '</tr></table>' +
    '<hr style="border:none;border-top:1px solid ' + navy + ';margin:16px 0">' +
    '<div style="color:' + navy + ';font-weight:700;font-size:8pt">' +
    escape(mail('pdf_subject', language)) + '</div>' +
    '<div style="font-size:12pt;font-weight:700;margin-bottom:18px">' + escape(docName) + ' ' +
    escape(number) + (order.project ? ' - ' + escape(order.project) : '') + '</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:9.5pt">' +
    '<tr style="color:' + navy + ';font-size:7.5pt;font-weight:700;border-bottom:1px solid ' + navy + '">' +
    '<td style="padding:6px 8px">' + escape(mail('pdf_description', language)) + '</td>' +
    '<td style="padding:6px 8px;text-align:right">' + escape(mail('pdf_quantity', language)) + '</td>' +
    '<td style="padding:6px 8px;text-align:right">' + escape(mail('pdf_amount', language)) +
    '</td></tr>' + rows + '</table>' +
    '<table style="width:100%;margin-top:26px;font-size:9.5pt"><tr style="vertical-align:top">' +
    '<td style="width:48%"><div style="color:' + navy + ';font-weight:700;font-size:8pt">' +
    escape(mail('pdf_payment', language)) + '</div>' +
    escape(COMPANY.paymentMethod) + '<br><span style="color:#6e768c">IBAN: ' + COMPANY.iban +
    '<br>SWIFT: ' + COMPANY.swift + '</span>' +
    '<div style="color:' + navy + ';font-weight:700;font-size:8pt;margin-top:16px">' +
    escape(mail('pdf_delivery', language)) + '</div>' +
    escape(when(order.eta)) + '</td>' +
    '<td><div style="color:' + navy + ';font-weight:700;font-size:8pt;border-bottom:1px solid ' + navy +
    ';padding-bottom:4px">' + escape(mail('pdf_vatsummary', language)) + '</div>' +
    '<table style="width:100%;font-size:8.5pt;margin-top:8px"><tr style="color:' + navy + ';font-weight:700">' +
    '<td>' + escape(mail('pdf_totalamount', language)) + '</td>' +
    '<td style="text-align:right">' + escape(mail('pdf_taxes', language)) + '</td></tr>' +
    '<tr><td colspan="2" style="color:#6e768c">' + escape(order.vatTreatment || '') + '</td></tr>' +
    '<tr><td>' + money(net, order.currency) + '</td><td style="text-align:right">' +
    money(vat, order.currency) + '</td></tr></table>' +
    '<div style="border-top:1px solid #ccd2e1;margin-top:10px;padding-top:10px;color:' + navy +
    ';font-weight:700;font-size:11pt">' + escape(mail('pdf_total', language)) + ' ' +
    '<span style="float:right;color:#111;font-size:12pt">' + money(gross, order.currency) + '</span></div>' +
    '</td></tr></table>' +
    '<hr style="border:none;border-top:1px solid ' + navy + ';margin:40px 0 8px">' +
    '<table style="width:100%;font-size:8pt;color:#6e768c"><tr style="vertical-align:top">' +
    '<td><strong style="color:' + navy + '">' + escape(COMPANY.name) + '</strong><br>' +
    COMPANY.site + ' - ' + COMPANY.email + '<br>tel: ' + COMPANY.phone + '</td>' +
    '<td style="text-align:right">' + escape(docName) + ' ' + escape(number) + ' - 1 / 1<br>' +
    escape(mail(isQuote ? 'pdf_note_quote' : 'pdf_note_order', language)) + '<br>' +
    escape(COMPANY.terms) + '</td>' +
    '</tr></table></body></html>';

  var stem = (isQuote ? 'quote-' : 'order-confirmation-') +
    String(number).replace(/[^A-Za-z0-9-]/g, '-');
  return Utilities.newBlob(html, MimeType.HTML, stem + '.html')
    .getAs(MimeType.PDF)
    .setName(stem + '.pdf');
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

/**
 * The quote the customer receives: the figures, and the two links.
 *
 * The links are the only two things in the message that do anything, and
 * neither of them does it on being followed - see decisionPage.
 */
function quoteBody(order, net, vat, gross, lang, quote, token) {
  var rows = [
    [mail('row_quote', lang), quote || order.id],
    [mail('row_reference', lang), order.id],
    [mail('row_solution', lang), order.solution],
    [mail('row_from', lang), order.source],
    [mail('row_into', lang), (order.targets || []).join(', ')],
    [mail('row_words', lang), thousands((Number(order.words) || 0) * (order.targets || []).length)],
    [mail('row_delivery', lang), when(order.eta)],
    [mail('row_total', lang), money(gross, order.currency)]
  ];
  var accept = decisionUrl(order.id, token, 'yes', lang);
  var decline = decisionUrl(order.id, token, 'no', lang);

  return shell(mail('quote_title', lang),
    '<p>' + mail('quote_intro', lang) + '</p>' +
    table(rows) +
    '<p>' + escape(order.vatTreatment || '') + '</p>' +
    '<p>' + mail('quote_valid', lang, { date: when(quoteExpiry(new Date())) }) + '</p>' +
    '<p>' + mail('quote_ask', lang) + '</p>' +
    '<p style="margin:22px 0">' +
    '<a href="' + accept + '" style="background:#1b2a6b;color:#fff;text-decoration:none;' +
    'padding:12px 20px;border-radius:6px;font-weight:700">' + mail('btn_accept', lang) + '</a>' +
    '<a href="' + decline + '" style="color:#1b2a6b;text-decoration:underline;' +
    'padding:12px 20px">' + mail('btn_decline', lang) + '</a></p>' +
    '<p>' + mail('quote_files', lang) + '</p>' +
    '<p>' + escape(COMPANY.terms) + '</p>' +
    '<p>' + mail('any_question', lang) + '</p>');
}

/** What arrives once the quote has been accepted: the work is on. */
function confirmBody(order, net, vat, gross, lang, quote, payUrl) {
  return shell(mail('confirm_title', lang),
    '<p>' + mail('confirm_intro', lang) + '</p>' +
    table([
      [mail('row_quote', lang), quote || order.id],
      [mail('row_reference', lang), order.id],
      [mail('row_solution', lang), order.solution],
      [mail('row_into', lang), (order.targets || []).join(', ')],
      [mail('row_delivery', lang), when(order.eta)],
      [mail('row_total', lang), money(gross, order.currency)]
    ]) +
    '<p>' + escape(order.vatTreatment || '') + '</p>' +
    payParagraph(payUrl) +
    (payUrl ? '' : '<p>' + escape(COMPANY.terms) + '</p>') +
    '<p>' + mail('any_question', lang) + '</p>');
}

/** And what arrives when it is declined: nothing happens, and that is said. */
function declineBody(order, lang, quote) {
  return shell(mail('decline_title', lang),
    '<p>' + mail('decline_intro', lang) + '</p>' +
    table([[mail('row_quote', lang), quote || order.id]]) +
    '<p>' + mail('any_question', lang) + '</p>');
}

function deskBody(order, net, vat, gross, repriced, senderTrouble, lang) {
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
  rows.push(['Writing to the customer in', lang || 'en']);
  if (repriced) {
    rows.unshift(['Check this', 'the page sent a total below the list price; the list price was used']);
  }
  if (senderTrouble) {
    rows.unshift(['Check this', senderTrouble]);
  }
  return shell('New quote · ' + order.id,
    '<p>Quoted, not ordered: nothing is agreed until the customer answers the ' +
    'two links in their copy.</p>' + table(rows));
}

/**
 * The desk's word that a customer has answered.
 *
 * On an acceptance it says plainly what has to happen next, because nothing
 * downstream of this is automatic: the Trados project is made by a person, from
 * the sheet's own menu. A six-minute execution limit and a 50 MB fetch limit
 * are not something to discover halfway through a customer's first job.
 */
function deskDecisionBody(row, at, decision) {
  var rows = [
    ['Quote number', at('Quote number') || '-'],
    ['Reference', at('Reference')],
    ['Customer', [at('First name'), at('Last name')].filter(String).join(' ')],
    ['Company', at('Company') || '-'],
    ['Email', at('Email')],
    ['Solution', at('Solution')],
    ['From', at('From')],
    ['Into', at('Into')],
    ['Words', at('Words')],
    ['Total', at('Total') + ' ' + at('Currency')],
    ['Delivery promised', at('Delivery')],
    ['Files', at('Files') || '-']
  ];
  if (decision === 'accept') {
    return shell('Quote accepted · ' + at('Reference'),
      table(rows) +
      '<p><strong>Next:</strong> open the orders sheet, select this row and use ' +
      '<em>Quote Desk › Create the Trados project for the selected order</em>. ' +
      'It is not automatic, and it will refuse any row that is not Accepted.</p>');
  }
  return shell('Quote declined · ' + at('Reference'),
    table(rows) +
    '<p>Nothing to do. The customer\'s files are untouched in the folder above; ' +
    'they are not deleted by this, and nobody should delete them without asking.</p>');
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
  'ooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAK' +
  'KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAr5/wD2xfFfiFvB2hfA/wABagbXxb8X9VHhm1njw0ll' +
  'p20vqV6E/iWK3DA4IYeapByBX0BXy/oniHQvi7+3ldXOiaxY63o/wk8DyWRaBxIthr9/eFZwrjKlxbQeW4Byp3KeQQPVyiFq0sTJ' +
  'XjSi5eV18N/JycbrqtDKq9OXvp/n+B9C+CPBnh34d+ENH8C+EtPSx0bQrOKxs4F52xooALHqzHqzHlmJJySa84/aj/aZ8GfsvfDm' +
  'Txn4lX7fqd47WuiaPHIEl1C5xnGcHZEgwXkIIUEAAsyK3sVfid/wUv8Aivf/ABG/ah1vQFvTLo/giKLQ7CJWbYsoQSXTlc4DmZ3Q' +
  'sMZWKMH7or1OFMmXEOZ+zxDbgryl3eu1/NvXra5liq31eleO+yPIvjl+0r8Yv2iNdfWPiX4tuLq1EpltNIt2aLTrLrgQwAlQQGK7' +
  '23SEY3M1fZn/AASG8JeFPFlv8W7bxT4Y0nWYUXQQseoWUdwqhjqG7AdSBnaufoPSvzjr9Mf+CM33fi7/ALvh/wDnqNfrXGNGnguH' +
  'a1PDRUEuSyWlvfj2PKwcnPEJy13/ACPof46f8E5f2cvi/pc8nh3wtaeAPEAj222peH7ZYLdWAO0S2a7YZFyQW2hJDgfOK/IL42fB' +
  'bx58AfiFf/Df4h6cttqNmBLDPES1ve2zE+XcQOQN8bYPOAQVZWCsrKP6I6+JP+Cr3wk0zxb+z/B8U44IU1bwJqEDGc53PY3cqW8k' +
  'Pv8Avnt3Gc4CNjG418DwXxXi6GNhgMVNzpzfKru7i3tZ9m9LfNeffjMLCUHOKs0fj5X6kf8ABIv463Os+HvEX7PuuXbSPoCnXdCD' +
  'ZJW0kkC3UQ7BUmkjcDqTcSdhX5b19DfsAeNT4G/a3+H17JcyR22q30miTor7RL9rheGNW9QJnibHqor9O4sy6OZZRWptaxTkvWOv' +
  '46r5nm4Wo6daL+R+7FFFFfzYfRBRRRQAUUV4xr37Zf7L/hfXNR8Na/8AGbQbLVNJu5rG9tpDJvgnicpJG2ExlWUg/SujD4Wvim1Q' +
  'g5Ndk3+RMpRj8Tsez0V4lYftr/soalOLe3+PHhRHbobi7MC/99SBVH516X4O+I3w9+IdvNd+APHfh7xLBbkLNJo+pwXixk9AxiZg' +
  'p9jVVsDisMuatSlFecWvzQKcZbM6KiisHxx468I/DXwte+NfHWu22jaHp3l/ar24z5cXmSLGmcAnl3RenUiueEJVJKEFdvRJbtjb' +
  'SV2b1FeE/wDDc/7JP/RdPDv5y/8AxFH/AA3P+yT/ANF08O/nL/8AEV3/ANj5j/0Dz/8AAJf5Ee2p/wAy+892oryfwP8AtW/s7fEr' +
  'xTZeCvAvxX0XWdc1HzPstlbmTzJfLjaR8ZUDhEduvQGui+J3xp+FnwZtLG++KHjSw8OwanI8VpJd7sTOgBYDaDyAR+dYSwOKhVVC' +
  'VKSm9lyu79Fa/RlKcWua+h21FeN+G/2xP2Y/F+v6f4W8N/GXw/farq1zHaWVqjurTzOdqRruUDcxIAGeSQK9kqK+Fr4VqNeDi33T' +
  'X5hGUZfC7hRVPV9W0zQNJvdd1q+hstP023ku7u5mbbHBDGpZ3Y9lVQST6CvFf+G5/wBkn/ounh385f8A4iqw+CxOLTdCnKVt7Jv8' +
  'glOMfidj3aisbwh4w8NePvDVh4w8HaxBqujanGZbS8hzsmQMVJGQDjKkdO1bNYSjKEnGSs0UnfVBRXnHxL/aL+CHwd1e20D4m/En' +
  'SPD+pXlsLyC1unbzHgLMgkwoOFLI4BPUqfQ10Xw++JPgT4reHU8W/DrxRY6/pDzPbi6tH3KJU+8hBAIIyDggcEHoRW0sJiIUlXlB' +
  'qD2dnZ/PYlTi3yp6nS0UUVzlBRRXmfxW/aW+BHwQbyPih8TtG0W7Ko/2Au1xe7Gztf7NCrzbDg/Nsxwea2oYetipqnQi5SfRJt/c' +
  'hSkoq8nY9Mor5qg/4KO/sYXE6W8fxnRWkYIpk0DVEXJ9Wa2AUe5IAr2nwF8Wfhh8UraS6+HHxB8PeJUhVXmXTNRiuHgDdPMRWLRk' +
  '+jAGujE5XjsHHnxFGcF3lFr80TGrCekWmdZRRRXCWFFFFABRXwd+0/8A8FKvEP7Pfxu8Q/CSx+FGna1Boi2bLey6q8Ly+faxTnKC' +
  'NgMGUr17Zr6u/Z9+KVz8a/g14W+Kd5o8WlTeIrRrl7OKYyrCRI6YDkAn7ueg616+MyPHYDCU8bXhanUtyu6d7q60TutO5lCvCcnC' +
  'L1R6HRRRXkGoUUUUAFFFFABRX50fF/8A4Ks+Jvhj8VPF3w5t/g1pl/F4Z1q80pLp9YkjadYJWjDlREQpO3OMmvu74VeMpviN8L/B' +
  '/wAQriwSxl8T6Dp+svapIXWBrm3SUxhiBuCl8ZwM4r2MwyHH5XRp4jFQtGezunfS/RvoY068KsnGL1R1NFFFeObBRRRQAUUUUAFF' +
  'FFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAVn6/wCINC8K6LeeI/E2sWelaVp8Rnu7y8mWKGCMdWd2IAH1rh/jr+0H8L/2' +
  'dfCDeL/iXrotUkDrY2EAEl7qMqgExwRZG48qCxIRdw3MoOa/Gf8Aao/bL+J/7UesiLWpTonhKylMmneHbSYtCjdBLO+B582ONzAB' +
  'cnYq7m3fVcOcJ4viCfOvcpLeT/KPd/guvZ8uIxUKCtu+x7z+2R/wUu134jre/Df9n67vdD8LszQ3mvgNDfaog42wjhreBuueJWGA' +
  'fLG9G7P/AIIzSRC7+LkTMPNaPQmUd9oN9n9StfmrX27/AMEk/G0GgftG6r4SvL4Qx+KfDlxDbQn/AJb3cEkcyge4hW5P4Gv1LPsi' +
  'w2W8N4jCYGFkkm31fLJNtvrovRdLI8yhXlUxMZzZ+wNfgd+2v4ZvfCf7V/xQ0u/YGWfxBPqakHP7u8Auo/8Axyda/fGvz3/4Kk/s' +
  'o6v46021/aE+H+lyXmp6DZ/Y/EVlbx7pJrFCWjulUcs0W5g/U+WVPCxHP55wBmlLLs09nXdo1Fy37O6a+/b1aPQx9J1KV49D8qa/' +
  'TH/gjN934u/7vh/+eo1+Z1fpj/wRm+78Xf8Ad8P/AM9Rr9P46/5EFf8A7c/9LieZgf48fn+R+ltfKf8AwU78S6ToX7H3ijS9RuPL' +
  'uPEV/pem2C4z5k63kVyV9v3VtKf+A19J+L/GPhXwD4dvfF3jXxBY6Lo2nR+Zc3t7MIooxnAGT1JJACjJYkAAkgV+LH7d37XT/tP/' +
  'ABBtrTwx9ot/A3hgyQ6PFKux7yV8CW8kXqCwVQityqDorO4r8n4LyWvmeZ068U1TpyUm+l1ql6t207anq4ytGnTcerPmGu7+A/hc' +
  '+Nvjd4A8Ifb9SsV1jxNpllJd6bL5V3bRvcxq80L4OyRFJZWwcFQccVwlfUn/AATT8E3PjH9rnwtdraLPaeGra+1u83H7iJA0UT/U' +
  'TzwV+65viVg8vr4j+WEn87O34nh0o89SMfM/RfR/jJ8T/wBmPU7HwL+1Hctr/g24kW10b4o2luViQk7Y4NYiXP2eXJCicZRsqWJI' +
  'ldfp2zvLPUbSDUNPuobq1uo1mgnhcPHLGwyrqw4ZSCCCOCDUWq6TpevaZd6Jrmm2uo6dfwvb3VpdQrLDPE4wyOjAqykEggggg18x' +
  '3Xww+K/7Jd5J4h/Z8s77xr8Li0lxqnw5mnaS+0sMS0k2jSuSWHJc2rE7m3bSWkBT+c7YfNdrU633Qn+kJf8Akj/u7P6H3qXmvxX+' +
  'f5+p9T0VxPwj+Mnw9+OHhKLxl8Otej1CzLGG5gYeXdWNwPvQXEJ+aKRfQ9RhlLKQx7avKq0qlCbp1U1Jbp6NGqakroK/ns/ab/5O' +
  'S+LH/Y8a7/6XzV/QnX89n7Tf/JyXxY/7HjXf/S+av0vww/3uv/hX5nm5n8EfU+gdE/4JT/tG+IvCum+LdJ8T/D+W31Wwg1C3tzqV' +
  '2s2yWMOqnNrsDYYZ+bGe9eG/En4H/tDfsseI9P1Txf4f1zwjfLKW0zW9Pu/3bSfN/qbu3Yqr7QTs3BwpyVGa/dT4Of8AJIvA/wD2' +
  'Lemf+ksdP+K3wy8LfGP4ea58NvGNms+l65aPbudqs8D4zHNHkECSNwrqezKKzw3iJjqWKdPGxjOlezsrO17adH6Na+RUsvg43g7M' +
  '+M/+CfH7e+s/GDU4vgj8Z7tJ/FnkvJomtLGqf2okalngnVQAJ1QMyuAA6KwbDrmX2H/gpH/yZb8RP+4R/wCnazr8XvDut+IvhP8A' +
  'EbTvENjtg1zwhrMV1GCcqtzbTBtp9RuTB9Rmv2e/4KNzxXX7E3xAuYH3xzJo8iN6qdVsyD+VdOfZHh8q4hwOIwitCrUi7LZNSje3' +
  'ZO609baEUK8quHnGe6T/ACPyA+A3wO8WftEfEa0+GPgrUNJstUvLee5jm1SWWO3CxIXYFo45GzgcfL19K+n/APh0H+0p/wBDv8M/' +
  '/Blf/wDyFXhf7G/xx8Lfs7/HPTfif4x03Vb/AEyzsry2eHTI43nLSxFFIEjouATz8351+g//AA9+/Z0/6EL4j/8AgDY//JdfT8SY' +
  '7iPD4xQyqlzU+Va2T1u79fQ5sPDDyheq9Ti/2Sf+CcXxv+Av7QfhT4seMPFPga80jQvt32iHTb68kuW86xngXYslsinDyqTlhwD1' +
  'PBtf8Fj/APkRfht/2Fr/AP8ARMdfU/7Mv7WHw/8A2qdN17VPAWh+INNi8PTwW9yurwwRs7SqzKU8qWQEAIc5x2r5Y/4LH/8AIi/D' +
  'b/sLX/8A6Jjr4nLcfmGP4qw7zNWqwvG1rWXLJr8zsq06cMLL2ez/AOAfmBomtap4b1qw8RaHeyWepaXdRXtncxnDwzxuHjdfcMoI' +
  '+lf0RfCH4jaX8Xfhf4X+JmjiNbfxHpcF+YkkD+RKyjzYSw6tHIHQ+6Gv59NM8C6tqvw+1/4iWnz2PhzU9M0y7RUJZTex3bRyEjhV' +
  'BsypJ/ikQd6/T3/gkV8Xz4h+GHiT4NandlrrwlfDUdNR3X/jxuyS6IvXCTrI7H1uVFfUeI2AjjMAsXT1lRlaXpK36uL9Gzmy+o4V' +
  'OR9T0b/gpx8Xz8M/2Z7/AMN6deGHV/HlymhQiOQLItpjzLt8H7yGNfJbHT7QtfixX2Z/wVQ+L5+IH7RK+A7C6MmlfD+xXTwodWjN' +
  '/OFmuXUjkHHkRMD0aA18q+PfA+rfDzxCnhrXMC8bTNM1N02lTGt7YwXaRsDyGVLhVYf3lNetwRgI5blVOM9J1bzfppb7ly/NmWNq' +
  'e0qtrZaH7gfsG/8AJonwz/7BT/8ApRLXvnSvA/2Df+TRPhn/ANgp/wD0olqj+338Zj8Fv2ZvE2pWN00Os+JFHhvSWUsGWa5VhJIr' +
  'LyjJAs7q3Z0Qd6/GMXhKmPzqphaXxTqyS+cn+R7MZqnRUn0X6H5J/th/Ghvjz+0N4s8dWt2Z9HS6/szRMOWQafb5jidMgFRIQ0xX' +
  's0zV9T/8EiPjSNG8Y+JvgTq98EtvEMP9uaPG7gL9thULcIgxlnkgCP14W1NfOn7Df7PGn/tI/HCPwf4jilPhzTtKvNR1Z4pDG6p5' +
  'flQ7GA+/580LYPVUf0rgfCOv+L/2Z/j1Z6w8WzX/AIe+ImhvLeKYqszW8xjuIN4H3JFEkZIHKua/bsxwOEx+Bq8P0Pip04tLtvy/' +
  'O8dfJ+Z4tOc4TWIls3/w5/QxRWd4c8QaR4s8PaX4q0C8W70vWbKDULK4UECa3mQPG4z2KsD+NaNfztKLi3GW6PoNzwj9tf48337O' +
  '/wCz7rnjjQZo4/EN7JFo+htIgdUvJ8/vMEFSY4kmlAYFSYwCME1+NfwY+DHxU/av+KVx4W8LXseoa/exT6vqmqazevtjjDqJLieU' +
  'h5HJkkReFdizg4xkj9Vf+Cn/AMNtf+In7Ld3d+HreS4l8H6xbeI7iCNCzyW0cU0MxAHZEuDKx7LGxr8uf2WP2k/EP7LXxOPxD0PQ' +
  'rXWobuwk0vUdPuJWh8+2eSOQhJAG8tw8SEMVYYyMHNfsXA9KVPIq9fLkniG2te6S5V006rpdnkY1p14xqfCfTt//AMEd/i9HpyS6' +
  'X8WPB9xflfmguIbqGEH0Eio7Ee+wfSvRf2FP2M/jv+zd+0pe658Q9AsZNBl8O3ljDrGm3yT20s7TW7BAp2zLkI3LxqOK9F8Bf8FY' +
  '/wBmvxM0Fr4x07xT4OnaPdPNdWIvLSN/7qvbM8rfUwrX1p4F+IXgb4neH4vFXw98V6Z4g0mY7RdWFwsqK+ASj45RwGGUYBhnkCvn' +
  'c3z/AIloYaphc0p2hNNXcbfdKNlf7zopUMNKSlSeq8zoap6rrOkaDZvqOuarZ6daR/fnu51hjX6sxAFfPP7dP7VR/Ze+F0F5oEcE' +
  '/jHxPLJZaFFMm6OHYoM1069GWMPGAvd5EyCu6vyR8M+Cv2kP2x/H2oXekW2uePfESx+ffXl5eKsdrEzEqrSzMsUKZ3bIwVHBCLxg' +
  'cGQ8ITzbDPH4mqqNFfafW2+7SSvpdvfoXXxapS5Iq7P3w0XxDoHiS0+3+Hdc0/VLbOPOsrlJ48/7yEitCv5//iB8Hf2jf2SfE+la' +
  't4p0rWvBOqXAd9M1XTtRUiTHDiO4tnIDYIyhYNtYZXBGf1I/4J8/thah+0r4O1Dwt4/kg/4TrwqkbXU8apGuqWj/ACrdCNcBXVht' +
  'kCgIC0ZGN+xbzzg6WW4RZhg6yrUerXTW19G01fRtPR9AoYxVJ+zmrM/Pb/gpT/yeb49/656T/wCmy1r9Rf2Df+TRPhn/ANgp/wD0' +
  'olr8uv8AgpT/AMnm+Pf+uek/+my1r9EP2ff+Ubdl/wBiFrH/AKBc19HxPD2vDeXQ7+zX3wObDO2IqP1/M+saK/n6/ZPdz+078KQX' +
  'Yj/hMNJ7/wDT1HX9AtfHcT8Of6t14Ufa8/Mr35eW2tu7OzDYj6xFu1rBXi/7Zmv674X/AGX/AIh+IPDOtX+kapZaV5lte2Fy9vcQ' +
  'P5qDckiEMpwSMgjrX4YfDKSQ/Enwn87f8hyw7/8ATwlfVf8AwUD/AGdfjdc/Hf4k/GSD4f6i/gkPa3J1cSReT5S2kEbNjfuwHBXp' +
  '2r6iHA9HLc0oUq+KVn7yvFK7jKK5dZauV/w2ZzPGupSk4x8v+DsN/YA+O3xv8ZftceA/DfjD4yeOdd0i8/tT7RYal4hvLq2m2aZd' +
  'Ou+KSQq2HVWGRwVB6iv2Or+cHwF4D8YfEzxZY+CPAWhz6vrupeb9ksoGUPL5cTSvgsQOER25PQV+hn/BNn9mf49fB349ar4p+J3w' +
  '51PQdJn8L3dlHc3MkTI07XNqyphHJyVjc9P4a9DjrIMHOTxqrxpyhDSnZJys3qveW97fC9jPA15r3OVu73PSPif/AMEqfBvxO+I/' +
  'if4i3nxd1qxn8TatdarJbRaZE6QtPK0hQMXyQN2M19kfDjwbB8Ovh54X+H1rfSXsPhjRbLRo7mRAjTLbQJEHKjgFgmSB0zX4Xfto' +
  'O4/as+KIDsB/wkdz39xX7DfDH4g+HfhV+xp4I+I/i24aLSfD/wAO9HvrkqV3yBdPh2xoGIBkdtqKCRlmUZ5rweJsuzGGX4OVfEOq' +
  'p25Y8qVrxVldPXsb4apTdSfLG1utz2t3SNGkkdVRQSzMcAAdyaxtK8b+C9dv5dL0TxdouoXsORLbWuoRSypjruRWJH4ivww+N/7S' +
  'Hx0/a9+ICafe3Gp3FrqN4kGh+EdLd3toSWIiRYlx503zHMrDcSTjaoVVk+If7Df7Uvwq8I3Hjzxl8LLm20ewTzbye01C0vGtU7vI' +
  'kEruFAGS+CqjliK6qfh/SoxhDMMZGnVntHR/JXkr9tFv1ZDx7bbpwbS6n7zUV+VH/BPL9unxvY+PtH+BXxf8S3euaH4glWw0TUdQ' +
  'laa6sL1jiGBpWy8kUjERqGyUZowCqZA9L/4LGMy/D34c7WI/4nN50P8A0wWvBnwjicPnFPKcRO3PtJK6as3e112s1fT0tfdYuMqL' +
  'qxW3Q/Qysi08YeEr/VpdAsfFOkXOpwjMtlDfRPOn+9GG3D8RX4L/AAt8R/tF/EPwj/wzN8IDrl/pmqalPrN5pWk7ka6aSOGBjdSg' +
  'gC2URx8OVjDPlsnbir8Yf2YPj3+z9ZabrHxX+H91oNnqczQWl2t5bXURmUbthkt5HVHIBIViCwViM7Wx9FDw7owrPDV8bFVH8MbL' +
  'ma72ck/kr+pg8wbXNGDt1/qx/QTSEgDJPFfDH/BLj9pbxj8XfCHiL4Z/ETWbjV9V8IC3uNP1G7mMlzcWUxdTHIx5cxOgw7EkiVR/' +
  'Bz8l/wDBRL9qjxz8UPjH4l+FGk6/e2HgfwjfSaN/ZkDtCl9dwPtuJrgA/vcTIyoG+VVRSAGZifAwfBuMxWbVMqlJR9nrKW6s7WaX' +
  'Vu+1117G88ZCNJVe5+v1t498C3mptolp400GfUUJVrSPUYWmUjqCgbcPyrer8NbD/gnL+2NqWj2utWnwhPl3kaTRwy61p8M6oyhg' +
  'XSSdShwRlWwwPBANfcv/AAS88BfGr4Z6P8RPBvxd8P8AijRYbK707+yrTVVlFqoK3HnNaliYmBIj3NESCQuT0rXOOF8BgMJPE4TG' +
  'xquNrxVr6tLpJ9+3zFRxM6k1GcGr/wBdj7looor4o7AooooAKKKKACiiigAooooAKKKKAPzL/wCCqP7LXiq81P8A4aX8L32oatpk' +
  'VtDY69p80zynTET5Y54ASQtuS2HRcbHbfgiRyn5q1/StqGn2GrWFzpWqWUF5ZXkL29zbXEYkimidSro6tkMpBIIIwQSK/E79vD9j' +
  '69/Zn8ejXfCtncS/DzxJO7aTOS0g0+fBZrGRzk5UAmMscugPLMkhr9m4B4njiKccpxTtOK9x90vs+q6d15rXx8fhuV+1jt1Plqu6' +
  '+BnxOvfgx8X/AAl8ULLzm/4R7VIbq4jhIDz2pO24hBPA8yFpEz/tVwtFfpdajDEU5UqivGSafo9GebFuLuj+lTS9T0/WtNtNZ0i9' +
  'hvLG/gjurW4hcNHNE6hkdSOCpUgg+hq1XxN/wS2/aFi+JXwcf4Ra/qAfxF8PwsNuJJMyXOkuT5DDJyfKbMJAGFQQd2r7Zr+X82y6' +
  'plONqYOrvF/euj+a1PpqVRVYKa6nx/8AtAf8EyfgZ8ZNSuvFHhK5ufh/4gvGLzSaZbpNp88hIzI9oSoDED/lk8YJJZgxJNfGnxX8' +
  'FftFf8EyLqzi+HPxo057T4jiUSy2uiwNI40/bt8xLqOUR/8AH82PLbnnPQV+xdfmh/wWb/5o/wD9zB/7j6+x4QzrG5hjqWU4yftK' +
  'Er3jNKXwxclq03o0ra9DjxdGFODqwVpeR8BfEj4yfFX4wagup/E3x/rfiOWN2khS+u2eG3LY3eTDxHEDgZCKo4rjaKK/cKVKnQgq' +
  'dKKjFbJKy+5HituTuwr9Vf8AgkP8HZdD8B+KPjbq1lsm8TXK6RpLvHhvsdsSZ5EbukkxCEf3rWvzT+F3w48SfF34h6B8NfCNv5uq' +
  '+IL1LODIJSIHl5X2gkJGgeRiAcKjHtX9CHw28AeH/hX4B0D4deFoDFpfh6wisLfcqh5Ai4aR9oALu252OOWZj3r878R83WGwUcvg' +
  '/eqav/Cn+rt9zPQy6lzT9o9kdLXE/GP4veDvgd4B1D4g+Nbp1tLTbDbWsC77nULt8iK1t06vK7cAdAAzMQqsw6PxN4l0Hwb4e1Hx' +
  'X4o1WDTdI0m2e7vbudsJDEgyzH8B0HJ6DJr5w+D3hzW/2m/H9h+1D8T9FuLDwnpBZvhh4ZvVwY4m/wCY3dJ0NxKADCOVjTay7jsl' +
  'P5FgcLCaeJxOlKG/eT6RXm+r+yrvsn605Ne7Hd/1c3f2YPgn4h0LXfE/7RPxWsLex+I3xLKTXumWahLfRbBdvkWeFAEs4VUMsrDJ' +
  'YYHRmk+haKKwxmLqY2s61TySS2SSskvJLRFQgoKyCv57P2m/+Tkvix/2PGu/+l81f0J1/PZ+03/ycl8WP+x413/0vmr9F8MP97r/' +
  'AOFfmedmfwR9T94/g5/ySLwP/wBi3pn/AKSx1ueKfE2ieC/DWq+L/El8lnpOi2c1/e3DAkRQRIXdsDk4APA5PQV+ZXhL/gr1L4T8' +
  'GaL4Uh/Z9W4fRtMttPW5bxVtEhhiWPeUFnxnbnG7jPWvnb9pH9u743/tK6WfCviC4sNA8LGVZX0bR0dI7pkbdGbiR2Z5dpwduVj3' +
  'KrbNygjgw3AGbYvFv6zFU6ber5ot2v0Sb19bGksfSjD3XdniJj1z4keOjFp9kbjWPFGrbYLaLkyXNzN8qL9XcAV+0f8AwUXtYrH9' +
  'iLx7ZQ58u3j0aJM9dq6rZgfyr5b/AOCcP7DnidPFOm/tDfF/QptLsdLP2jwzpF7BtnurjGEvZY2GY40yWiBAZn2yDCqpk+qf+Ckf' +
  '/JlvxE/7hH/p2s69riPNsPjuIMBhMM01SqRu1tdyjp8kl87roY4elKGHnOXVP8j8pP2PvgZ4d/aK+OGm/C/xTq2o6bp95ZXdy9xp' +
  '5jEwaKIuoG9WXBI54r7/AP8Ahz58C/8Aopfjv/v5Z/8Axivzv/Zd+PX/AAzZ8XrD4q/8Ip/wkX2K0ubX7B9u+yb/ADoym7zPLkxj' +
  'OcbefavtX/h83/1bf/5eH/3DXtcUUOJ6mNTyhv2XKtpQWt3f4mn2McNLDKH73f5n19+y7+yb4K/ZU0zX9L8GeI9b1aPxDPBcTtqj' +
  'RFo2iV1AXy0Xg7znOelfL3/BY/8A5EX4bf8AYWv/AP0THVn4Yf8ABWj/AIWP8SvCXw8/4UD/AGd/wlGuWGi/bP8AhKvO+zfabhIf' +
  'N8v7Gu/bv3bdy5xjI61W/wCCx/8AyIvw2/7C1/8A+iY6+PynAZpg+JsPUzZWqTbd24u9otfZbOurUpTw0lS2X+Z49/wTq+FVv8bv' +
  'gn+0j8L5lj87XNM0OOyaRyqR3sZv5bV2I52rPHEx9QCO9eG/sa/HSP8AZm+P9l4z8Sx3kGkfZL3Sdcto4czNEyEqm08grcRQMe/y' +
  'EV9bf8EZP+awf9y//wC5Cvm//go38IR8KP2ofEFzZW/l6V4zVfE1nglsPcMwuVJPAP2lJm2jorp7V9pRxFLF57mGS4n4KkYtf+C4' +
  'qVvO1mv8JxuLhQp1o7r/ADZxPwE8D6x+1L+1HomjeJT9rk8W+IJtZ8QyjMYe33vdXhyv3CyiRVP951Heuu/4KQqqftofENVUBQNI' +
  'AAHAH9k2dfTf/BH34RbYvGnx01G35cr4Y0ptx6DZPdkrjnJ+yhWHpIPWvmX/AIKR/wDJ6XxE/wC4R/6abOujCZisVxTPCU/go0XH' +
  'yu5Qb/RfImdPlwqm92/0Z+pX7Bv/ACaJ8M/+wU//AKUS18A/8FYvjL/wmnxu034U6Zc7tO8BWX+kgDhtRulSSTkHDBYRbqOMqxlH' +
  'rX2z+yz490b4W/sB+FfiN4hJ/s/w54Yu9RmQOqtL5cszCJC3G92ARR3ZgO9fj3pGn+Mf2hPjRb2LTLceJfiD4i/ezeWfLFzdz5kl' +
  'KrnbGpdmOOFVT2FfM8JZdGed4zM62kKUp2fm27v5Rvf1R04uo1RhTW7sehfsw/th+N/2VLbxBH4E8H+GdTuPEj27Xdzq0dw8ipAH' +
  '8uNPLlQAZlcnjJJGTwMef/G/4t6r8dPibq/xT17QtK0jUtbMLXdvpaSJbmSOJY94V2YgsEBbnlsnvX7EW/8AwTY/Ywhgjhk+EEk7' +
  'ogVpZPEOqBpCByxC3IXJ68AD0AryX9rH/gnb8AtB+APi3xT8GPh42jeKfD1n/a8Eo1i+uBLBAd9xEUnlkUkwiQrgBi6oM4Jz6OC4' +
  'z4fqZj7enRnGrUtFzaXWyV/femi6aWM54PEKnytqy1t/SOj/AOCVvxnPxB+AU/w61O6MurfD28+xruZmZtOuC0lsxY/3WE8QUfdS' +
  'JPUV9pV+Hf8AwTu+Mp+D/wC034fjvrny9G8Y/wDFNagCCQDcMv2d+oA23Cwgsfuo0nrX7iV8LxzlX9m5vOcF7lT316v4l9936NHb' +
  'gqvtKST3WghAIwRkGvkn4vf8Exv2afifqFzruiafqfgbU7je7/2BKiWbynozW0isigf3YvKB+vNdL+3f+0X43/Zj+Fvh74g+BLHS' +
  'b26u/E8Gl3VtqcLyQy2z2l1Iw+R1ZW3QoQwPbkEZFfPfhf8A4LH+D7iHHjX4JazYSqAN2l6rFdq5xycSJEV57Zb61nkmVZ+qKzDK' +
  'b2d0+WSW3dNq/wBzKrVaF/Z1Txn4/wD/AASv+Ifwl8Fat8QPBPxB03xhpmg2c2o6jazWZ0+8jt4hudo1LyJLtQMxBdCQpCqxIWvK' +
  '/wBgT4z+KPhJ+0n4SsNIvbg6R4z1S08O6vp6yYiuUuZBFDIynjdFJIrqw+bAdQcOwPu/7R3/AAVVl+Jfw/134b/C/wCG9xo0HiKy' +
  'l0671bVbtJJltZkZJkjgjUqrMrECQyHAJwucMPEv+Ce/wR8Q/F39pHwxrFnZ3K6F4Hv7fxFqt+qfu4WgfzLaIseN8s0aqFzkoJGA' +
  'IQ1+m4apmMsixL4kS2dr2va2l+XS97cuzv8AI8ySp+3j9XPVf+CwGp38v7QPhPRpLqRrK18HQXUMJPypLLe3ayMB6sIYgf8AcFfS' +
  'v/BJHRNKsf2bdX1q1jia+1PxTdC7lCjfiOCBY4y3UqoLMAehkbHWuV/4Kz/AHWvF/hXQPjn4V0t7yXwpFLp2vLChaRdPdt8U+M/c' +
  'ikMgbAJAnDHCoxHy5+w1+3Mf2WBq3g/xf4evdb8Ga3dDUHSwaMXdjd7AjSxq5VZQ6JErIzrjYpBHIbxIYapnvB9PD5frODV49dG7' +
  'r1d1Lz9TdyVDFuVTZn6A/wDBTbQNJ1j9jvxdqOo2iTXGh3ml39g7dYZ2vobcuPfyriZfo5r4H/4JXavqGm/tZWNnZE+TquhajaXe' +
  'P+eQRZhn/tpDHWv+3D/wUGs/2k/C1n8Mvht4a1XRPCwuo7/UbnVGRLu+kQHy4TFE7okSsd5y7FmEZ+TYd3q3/BI34Da5bajr/wC0' +
  'Nr1jLbadPZPoOg+bHgXZaVGubhM87UMSRBhkEtMvVCKMNhKuQcI4inmK5ZTcuWL395JJet05W6LfqEpKvi4un0PnH/gpT/yeb49/' +
  '656T/wCmy1r9EP2ff+Ubdl/2IWsf+gXNfnf/AMFKf+TzfHv/AFz0n/02Wtfpn+xpoNr4p/Yf8E+GL5mW21fw5d2ExXqElknRiPwY' +
  '1lxNNU+HMtm9l7N/+SDwyviKi9fzPyG/ZO/5Od+FP/Y46T/6VR1/QPX87iQ+Nv2c/jXatrWjLa+J/h94hgu3s7pW8tri1nWRM4wW' +
  'ifYpDKcMjAg4INfor4p/4LB/DyTwBcv4N+GXiaLxpNaFYIdQNudNtrllxvMySGSVEJzjykL4xlM5HVx1keOzvEYetgIc8XG101bV' +
  '3T9GnuTga8KMZRqOzPzV+GP/ACUnwn/2HLD/ANKEr9wf26f+TSPib/2Bv/a0dfj3+yH8LNc+L37RXgjwvo+nyXMFrq1tququAdlv' +
  'p9tKkk7uwBC5UbFJ4LyIvVhX7L/tjeHtQ8Ufst/E3SNLiaW5Ph25uUjVdzOIQJmVQOSSIyABzkis+Oa9NZzgIX1i035Jyj/kx4KL' +
  '9jN/1sfk7/wTc/5PS+Hf/cX/APTTeV+41fz6/su/GPT/AIAfHjwr8W9W0W41ay0KS5FxaW8ipK8c9rLbsyFuCyibcFOA23GVzkfq' +
  't8CP+Cifw+/aH+L+mfCrwN4B8Q2a31pdXUuoatLBCYvJjL7ViiaTfnGMl1x156Vy+ImT47F41YyjTbpwp6vSys5N9ezX6FZfWhCH' +
  'I3q3/kfl9+2j/wAnWfFH/sY7n+Yr7u/a61jUNM/4JffDiysmYQ6to3hGzuwB1iFnHMAf+2kMf5V8I/to/wDJ1nxR/wCxjuf5iv1P' +
  '1r4N3Px6/wCCeHhj4aaYyLqt98PtAudLZ9oBvLe0t5okLNwodo/LLdlkY16meYinhMPlNet8MZQb8lyx1+W5lQi5yqxW+p8Nf8El' +
  'tB0nV/2ob/UNRs0muNE8JX1/YOw5gna4tbcuvv5VxMv0c1+weoWFjqthc6XqdpDdWd5C9vcQTIHjlidSrIyngqQSCD1Br+f/AOBH' +
  'xj8b/sq/Gu08b2miyjUdDmn03WdFvS9s08JJSe1l43RuGXIyp2SRqSrbSp+6fi1/wV28Iaj8PL3Tfg/4E8S2Xi3ULZrdLzV1tktt' +
  'Nd0IM0flySGZ0J+UMqAnBPAKHz+M+G8zzXNYYjCR5oSikndWja+/Za3v+ppg8TTpUnGbsz82biefwj4wkufDmpuJtF1IyWF7GRuD' +
  'Qy5jlUjjOVVhX6cf8Fjv+Se/Dn/sM3n/AKIWvib9i39n/Wv2gvjtoGiRaW8/hzRbuDVfEdy6EwxWUbhjCxyPmmK+UoGT8zNgqjEf' +
  'bP8AwWO/5J78Of8AsM3n/oha93OcTSnxJl+Gi7zhzuX/AG9HS/3N28/Mwoxaw1SXR2Oc/wCCMiqW+MDlRuA8PgHHIB/tD/AV6j/w' +
  'V3VW/Zn0AkZK+NrIj2P2K+H9a8v/AOCMn/NYP+5f/wDchXqP/BXb/k2bQv8AsdbL/wBIr2vlsd/yXS/xQ/8ASInVD/cfk/zPn/8A' +
  '4I7/APJXfHY/6luP/wBKkrjv29/2NPix4K+MPiv4p+E/CWqeIvBvim/n1xr3T4WuXsJ52aW5juEQbolWQuVcjZsZBu3ZUdh/wR3/' +
  'AOSu+O/+xbj/APSpK9313/gqz8PvAnxN8ZfDr4g/CzxDCvhXXtQ0aK+0a6guzdi2uJIhI0Uph8stsBwHfr1r1cdic0wfE+IrZZS9' +
  'paEOaN+lla2u9+1/QyhGlPDRVR21dj4g+Fv/AAUT/ao+FdpY6TbeOLfxJpGnxeTDp/iGzW6Xb2BnXZcnA4AMuAOMYr9QP2N/2wPD' +
  '37V/hHULpdGOheKPDzRR6xpgkMsW2Td5c8DkDdG+x/lPzIQQcgq7/nt+3N+2R8C/2lvD2m6d8PfhDe2OvQ3sdzceJdYtba3vlhRH' +
  'X7MvkPI0iMXBO9wo2DCkkMvpf/BHbwLr7+K/HvxMeOeLRIdOh0KN2UiO4unlWZgp6Fo0jTI6gTr/AHqXEmVYPFZJPM62FWHrLppr' +
  'qlrayd79Un123MPVnGsqalzRP1Fooor8dPXCiiigAooooAKKKKACiiigAooooAK5T4o/DLwh8YvAWsfDjx1p32zR9agMMyqQJIm6' +
  'pLGxB2yIwDK2DgqOCOK6uirp1J0ZqpTdpJ3TW6aE0mrM/nz/AGkP2fPGH7NfxOv/AId+K/8ASYV/0nStTSMpFqNkxISZQc7W4Kum' +
  'TtdWGWGGPltfvn+1j+zJ4X/ah+GM/hLU1gtNf0/fdeHtWZfmsroj7rEDJhkwFkXnICsBuRCPwo8b+CvFHw58W6r4G8a6NPpWt6Lc' +
  'ta3tpMPmjcdwRwysCGV1JVlZWUkEGv6F4R4lhn+F5ajtWh8S7/3l5Pr2flY8DF4Z0JXWzOt/Z4+N3iL9nn4t6H8UPD2+YafL5WoW' +
  'QkKLf2L8TQMenK8qSCFdUbBKiv348DeNvDPxI8H6R478HanHqGi65aJeWdwn8SMOjDqrqcqynlWBUgEEV/N9X3X/AME0f2wU+Ffi' +
  'dfgX8RdWaPwj4kugdHup2zHpWoucbCTykMxwD/CkmGwoeR68rj3hx5nh/r+GV6tNar+aP+a3XlfyNcDiPZy9nLZ/mfrnX5of8Fm/' +
  '+aP/APcwf+4+v0vr80P+Czf/ADR//uYP/cfX5zwL/wAj/D/9v/8ApEj0Md/Al8vzR+aFFFfV37AP7Idx+0f8Qf8AhKPF1hMPh94W' +
  'uEfUnIKrqV0MMlirdwQVaXbysZAypkRh+95jmFDK8NPF4h2jFff2S83sjwqdOVWShHc+tf8Aglp+y0/gbwjJ+0J4007Zrniq28nQ' +
  'IZUIa00tiCZ8Ho05ClTjiJVIbErCvvokAEkgAckmmxRRQRJBBGkccahERAAqqBgAAdAK+YfjL4n1v9pXx9f/ALLHws1u5sPDmmBf' +
  '+FneJ7Juba3bP/Emtn5BuZgCJDysablIc746/nbF4mvxLmFTF13yx3b6QgtF/klvKT7s+hhGOGpqEf8Ah2Z0n2r9t74i+UjZ/Z/8' +
  'D6h87D/V+OdYhb7o7SafA4/3JXH8eMxfWKqqKERQqqMAAYAFZfhbwt4e8E+HNN8I+E9Jg0zR9ItktLK0gGEhiQYAGeT7kkknJJJJ' +
  'NatefjsYsQ406S5acNIr82+8pbt+iWiSWkIcur3YUUUVwFhXyh4z/wCCZ/7NPjvxhrvjjXU8Vf2l4i1K61W88nVQkfnzytLJtXyz' +
  'hdznAzwK+r6K7cFmOLy6TlhKjg3vZ2uROnCppNXPj+y/4JWfsnWtws0+m+KLxFPMU2tMEb67FVvyNeu/DP8AY+/Zp+EGoR6x4D+E' +
  'Oi2upQyLNDfXnmX9zBIOjRS3LSNEfdCtex0V0YjPczxUXCtiJuL6czt917ExoUou6igrjfjB8KPCnxv+HWrfC/xuLw6LrXkfavsk' +
  '3lTfuZ4502tg4+eJc8dMiuyorzqVWdCpGrTdpRaaa6NapmjSkrM+OP8Ah1J+yp/c8Yf+Dhf/AI3R/wAOpP2VP7njD/wcL/8AG6+x' +
  '6K9n/WfOf+gmf/gTMfq1H+VHyj4L/wCCZ/7NPgPxjoXjnQU8Vf2n4d1O11ay87VQ8fn28qyx7l8sZXcgyM8ivU/2g/2Yvhl+0zpu' +
  'j6V8Sxqxg0OeW4tRp92IDvkVVbcSrZGFFet0Vy1M6zCtXhialaTnHZ31V+xao01FxS0Z43+zx+yj8Kf2Yv8AhIP+FYrrA/4SX7J9' +
  'u/tC8E//AB7+b5ezCrt/1756549KX9ob9lL4S/tOLoX/AAs221QyeHTcfYpdPvPs7gT+X5it8pDD90hHHGDjqa9jorP+08Z9a+u+' +
  '1l7X+a+u1t/TT0H7OHLyW07HG/CD4TeD/gf8PdL+GXgS3uItG0nzjB9pl82ZmlleV2d8DcSzt24GB0FeN/GD/gn3+z/8b/iLq3xQ' +
  '8br4kOta15H2r7JqQih/cwRwJtXYcfJEueeuTX0tRRQzTG4WvLFUaso1JXvJPV3d3d+b1CVKEoqLWiPHNZ/ZW+Gmt/AOy/ZtuLvX' +
  'ofBtkIl8uG9VbidI5jMiSSFDlRJtbgA5ReeueR+DX7AP7PnwL+IFh8TPBdnr0ut6XHMlo1/qPnRxGWNo3YKEXLbHdRn+8eM4I+ka' +
  'KuOcY+FKdGNaSjNtyV9G3vfvfqJ0oNptaoKbJHHNG0UqK6OCrKwyGB6gjuKdRXmmh8dt/wAEqf2VfPM8MXi+E796iPWcBOcgAmMn' +
  'jtzmvsC3iaC3igeeSdo0VDLJje5AxubaAMnqcAD2FSUV343M8ZmPL9bqufLtd3tff8iIU4U/gVjkPij8JPhz8avCz+Cvih4UtNf0' +
  'dpluFgnZ0aKZQQskckZV43AZhuRgcMwzgkH5U1z/AIJIfsz6pfTXmmeIvHujxytuS0ttStpIYh6L5tu8hH+85NfbVFa4HOswy2PJ' +
  'hK0orsnp92xM6NOprNXPivw1/wAEl/2YdE1GG/1jVfG/iGKM/PZX2qQxW8v+8beGOUfhIK+sPh/8OPAnwr8NW/g/4deFdO0DR7bl' +
  'bWyhCBnwAZHb70khCjLuSxwMk10lFTjs5x+ZpLF1pTS6N6fdsOFGnT+BWGSxRTxPDNGskcilXRhlWU8EEHqK+SviX/wS+/Zd+Iet' +
  'Sa7pun674MmmZ3mt/Dt5HFayOxzkQzRSrGB0CxbFA7V9cUVngczxmWTc8JUcG97Pf1Wz+Y504VFaaufHXgH/AIJW/sveDNZi1rWB' +
  '4o8XGAhks9a1CP7LvByGKW8URfpyrsykcFSK+vNO07T9H0+20nSLC3sbGyhS3tra2iWKKCJAFRERQAqgAAADAAAFWaKeOzTG5nJS' +
  'xlVzttd7ei2QoUoU9IKx81/GL/gn98APjj8RNU+J/jhfEh1rVxALn7HqQiixDCkKbV2HHyRrnnrmva/hf8OPDnwi8A6N8N/CX2r+' +
  'yNCgNvafapRJLsLs/wAzADJyx7V1NFKvmWMxVGOGrVHKEbWTeisrK3y0HGnCLcktWeM/H39kT4GftIpHc/EXwu66zbwm3t9b02b7' +
  'Nfwx5yF3gFZVHO1ZVdV3NgDJz8/R/wDBIL9nNZ97+O/iM8WeE+32IP4n7J/SvueiuvCcQ5pgKSo4evKMVsr6L0vt8iJ4elN80o6n' +
  'nXwW/Z8+EX7PuhzaF8KvB9vpKXhRr26Z2mu7xlBCmWZyXYDLELkIpZtqrk16LRRXl169XE1HVrScpPdt3b+bNYxUVZHyH8TP+CXX' +
  '7MXxD8Q3HiTT4PEPg+a6ZpJrTw/dwx2bSM2SwhmikEfXAWMogHRRXQfs+/8ABPn4L/s5ePIfiR4R1/xdqeswW01rF/at5bvAiyLt' +
  'chIoIznHHLEe1fTlFerPiLNamHeFnXk4NWabvp2vuZLD0lLmUdT5a+In/BOD9nL4n+ONb+IXihfFB1bX7x767+z6oscfmOedq+Wc' +
  'D2zX0b4N8K6X4F8IaH4I0Pzv7N8PabbaVZ+c++TyIIlij3NgZbagyccmtiiuPE5ljMZTjRxFRyjHZN3S6afIuNOEG3FWueEfHv8A' +
  'Yo/Z/wD2irt9b8b+F5rHxC6LG2u6NMLW9dVAAEhKtHMQqqoMqOVUAAgV4pYf8EiP2b7W9jubzxj8Qr2GNw/2eTULNFkAP3XKWwbB' +
  '6HaVPoRX3FRXZheIs1wdJUKGIkorZX29O3yIlh6U3zSirnIfC74R/Df4LeGE8HfC/wAI2OgaUjmVorcMzzSHgySyuTJK+ABudmOA' +
  'BnAAHK/tB/syfDP9pjStI0b4ljVjbaJcSXNqNPuxA291CncSrZGBXrNFefTx2JpYj61Co/ab819b977mjhFx5WtDxv8AZ4/ZR+FP' +
  '7MX/AAkH/CsV1gf8JL9k+3f2heCf/j383y9mFXb/AK989c8elbfx7+AHgD9o/wAG2vgX4jDUjplpqMeqR/YLkQSeekcka5Yqfl2z' +
  'PxjrivSaKuWY4ueJ+uSqN1dPevrorLX0EqcVHktoeF/s/fsbfBz9mnX9T8SfDUa4LvVrMWNx9vvhOnlhw4wAgwcqOa8/+LP/AATK' +
  '/Zv+LHjDU/HNxP4s8P6prV3cahqB0nU0MVzdTSGSSVluYpSpLsx2oVUZ4AGK+tKK6aee5lSxDxUK8vaNWbvq0tr9yXQpuPK4qx8T' +
  '+Hf+CSf7M2j6nBqGr674516GF9z2N3qUEVvMP7rmCBJcf7simvr3wZ4K8J/DvwzYeDfA/h+y0TRNMj8q1srOIJHGCSSfUszEszHL' +
  'MxLEkkmtuis8fnGPzNJYyrKaWyb0+7a/mOFGFL4FYKKKK800CiiigAooooAKKKKACiiigAooooAKKKKACvjj/goV+xlF8fvCR+JX' +
  'w/sAPiD4btjiGNOdas1yxtjjnzl5MTd8shHzKyfY9Fd+WZlXynFQxeGdpR+5rqn5MipTjVi4SP5o3R43aORCrKSGUjBBHY02v0j/' +
  'AOCmf7F72NxfftKfCzRlNrOxm8X6dapgxSE86iiDqrE/vschj5hBDSsv5uV/SGS5xh88wkcXQ67rrF9U/wBO61Pna1GVCfJI/Wz/' +
  'AIJw/trn4o6Ra/An4qawH8YaTb7dE1C4f59YtI1z5Tsfv3ESjk/ekRdxyyuzcF/wWb/5o/8A9zB/7j6/NrTNT1LRNStNZ0bULmx1' +
  'CwnjurW6tpWjmgmRgySI6kFWVgCCDkEAivvLRbT4tf8ABVLQfCmjane+H/DeofCQPD4i1yd5GbVE1HYIp4LSOMIsqjT5PMQyIhZw' +
  'V2g7F+QxfD1DIc6p57BqNBczmv5W4tJpdU20rLZvTTbrhiJV6LoPWXTzPl/9mT9m3xt+058SLfwT4WRrTTrfbca3rEkZaHTbXPLn' +
  'pvkbBWOPILN3VVd1/dn4Y/DXwh8IPAukfDrwLposdG0WAQwJwXkOcvLIwA3SOxZmbHLMelc9+z/8AfAP7OXw+tfAPgOzOxSJr+/m' +
  'A+0ahckANNKR34wFHCgAD3pftDfHFfg34bsbPw7pH/CRePfFdwdL8I+HY2/eajekDLsMjbbxAh5ZCVVVwCylga+E4k4gxHFWMjhs' +
  'MrUk/dW1+8pdtO+kV8zuw2HjhYc0t/60OZ/aJ+MXi6HXdM/Z4+BXl3HxO8XQNNJesu+DwvpOdsup3H+0M7YkP3nI6/Kknovwa+D/' +
  'AIQ+BvgGx8A+DreTyIC1xeXtw2+61K8fBmu7h+ryuRkk8ABVXCqoHN/s6/A24+EWhajrfjHXB4k+IvjG4XU/F3iBhzd3WMLBFkDb' +
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

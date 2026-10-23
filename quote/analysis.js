'use strict';

/**
 * File analysis — the counting a CAT tool does before it quotes.
 *
 * Opens the file, pulls out only the translatable text, and reports words,
 * characters, segments and exact repetitions. Office and InDesign files are ZIP
 * containers, so they are unzipped in place; bilingual formats (XLIFF, TMX) are
 * counted on their source side only, the way an analysis report does.
 *
 * No dependencies and no uploads: it runs on the bytes the browser already has.
 * Works in a browser and in Node 18+ (both provide DecompressionStream).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.FileAnalysis = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  /* --------------------------------------------------------------- counting */

  // Han, Hiragana, Katakana, Hangul: counted per character, as CAT tools do.
  const CJK = /[぀-ヿ㐀-䶿一-鿿豈-﫿가-힯]/g;
  const HAS_CONTENT = /[\p{L}\p{N}]/u;

  function countWords(text) {
    if (!text) return 0;
    const cjk = (text.match(CJK) || []).length;
    const tokens = text.replace(CJK, ' ').split(/\s+/);
    let words = 0;
    for (const token of tokens) if (token && HAS_CONTENT.test(token)) words += 1;
    return words + cjk;
  }

  const countCharacters = (text) => (text ? text.replace(/\s+/g, '').length : 0);

  /** Split into segments the way a CAT tool does: sentence by sentence. */
  function segment(text) {
    const out = [];
    for (const line of String(text || '').split(/[\r\n]+/)) {
      for (const piece of line.split(/(?<=[.!?…。！？])\s+/)) {
        const trimmed = piece.trim();
        if (trimmed && HAS_CONTENT.test(trimmed)) out.push(trimmed);
      }
    }
    return out;
  }

  /** Exact repetitions: segments identical to one already counted. */
  function repetitions(segments) {
    const seen = new Set();
    let repeated = 0;
    let repeatedWords = 0;
    for (const s of segments) {
      const key = s.toLowerCase().replace(/\s+/g, ' ');
      if (seen.has(key)) {
        repeated += 1;
        repeatedWords += countWords(s);
      } else {
        seen.add(key);
      }
    }
    return { repeated: repeated, repeatedWords: repeatedWords };
  }

  /* -------------------------------------------------------------------- zip */

  const dv = (bytes) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  async function inflateRaw(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  /** Read a ZIP container into { name: Uint8Array }, for the entries asked for. */
  async function readZip(bytes, wanted) {
    const view = dv(bytes);
    let eocd = -1;
    for (let i = bytes.length - 22; i >= 0 && i > bytes.length - 66000; i--) {
      if (view.getUint32(i, true) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error('not a zip container');

    const count = view.getUint16(eocd + 10, true);
    let offset = view.getUint32(eocd + 16, true);
    const files = {};

    for (let i = 0; i < count; i++) {
      if (view.getUint32(offset, true) !== 0x02014b50) break;
      const method = view.getUint16(offset + 10, true);
      const compressedSize = view.getUint32(offset + 20, true);
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const localOffset = view.getUint32(offset + 42, true);
      const name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
      offset += 46 + nameLength + extraLength + commentLength;

      if (!wanted(name)) continue;

      const localNameLength = view.getUint16(localOffset + 26, true);
      const localExtraLength = view.getUint16(localOffset + 28, true);
      const start = localOffset + 30 + localNameLength + localExtraLength;
      const raw = bytes.subarray(start, start + compressedSize);
      files[name] = method === 0 ? raw : await inflateRaw(raw);
    }
    return files;
  }

  const decode = (bytes) => new TextDecoder('utf-8').decode(bytes);

  /* ------------------------------------------------------------ xml helpers */

  const unescapeXml = (text) =>
    text
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (m, code) => String.fromCodePoint(Number(code)))
      .replace(/&#x([0-9a-f]+);/gi, (m, code) => String.fromCodePoint(parseInt(code, 16)))
      .replace(/&amp;/g, '&');

  const stripTags = (xml) => xml.replace(/<[^>]*>/g, '');

  /** Text inside every occurrence of one element, in document order. */
  function textOf(xml, tag) {
    const out = [];
    const pattern = new RegExp('<' + tag + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + tag + '>', 'g');
    let match;
    while ((match = pattern.exec(xml)) !== null) out.push(unescapeXml(stripTags(match[1])));
    return out;
  }

  /** Drop elements we never translate, then keep what is left. */
  function visibleText(xml, dropTags) {
    let text = xml;
    for (const tag of dropTags || []) {
      text = text.replace(new RegExp('<' + tag + '[\\s\\S]*?</' + tag + '>', 'gi'), ' ');
    }
    text = text.replace(/<!--[\s\S]*?-->/g, ' ');
    text = text.replace(/<(p|div|br|li|tr|h[1-6]|section)\b[^>]*>/gi, '\n');
    return unescapeXml(stripTags(text));
  }

  /* ---------------------------------------------------------------- formats */

  /** One line per block, so segmentation stays honest. */
  const lines = (parts) => parts.filter((p) => p && p.trim()).join('\n');

  async function fromDocx(bytes) {
    const parts = await readZip(bytes, (n) =>
      /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n)
    );
    const blocks = [];
    for (const name of Object.keys(parts).sort()) {
      const xml = decode(parts[name]);
      // One line per paragraph; <w:instrText> holds field codes, never text.
      for (const paragraph of xml.split('</w:p>')) {
        const clean = paragraph.replace(/<w:instrText[\s\S]*?<\/w:instrText>/g, '');
        const text = textOf(clean, 'w:t').join('');
        if (text.trim()) blocks.push(text);
      }
    }
    return lines(blocks);
  }

  async function fromXlsx(bytes) {
    const parts = await readZip(
      bytes,
      (n) => n === 'xl/sharedStrings.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(n)
    );
    const blocks = [];
    if (parts['xl/sharedStrings.xml']) {
      const xml = decode(parts['xl/sharedStrings.xml']);
      for (const item of xml.split('</si>')) {
        const text = textOf(item, 't').join('');
        if (text.trim()) blocks.push(text);
      }
    }
    for (const name of Object.keys(parts)) {
      if (!/worksheets/.test(name)) continue;
      const xml = decode(parts[name]);
      // Inline strings only: shared strings are already counted, formulas are not text.
      for (const cell of xml.split('</c>')) {
        if (!/t="inlineStr"/.test(cell)) continue;
        const text = textOf(cell, 't').join('');
        if (text.trim()) blocks.push(text);
      }
    }
    return lines(blocks);
  }

  async function fromPptx(bytes) {
    const parts = await readZip(bytes, (n) =>
      /^ppt\/(slides\/slide\d+|notesSlides\/notesSlide\d+)\.xml$/.test(n)
    );
    const blocks = [];
    for (const name of Object.keys(parts).sort()) {
      const xml = decode(parts[name]);
      for (const paragraph of xml.split('</a:p>')) {
        const text = textOf(paragraph, 'a:t').join('');
        if (text.trim()) blocks.push(text);
      }
    }
    return lines(blocks);
  }

  async function fromOpenDocument(bytes) {
    const parts = await readZip(bytes, (n) => n === 'content.xml');
    if (!parts['content.xml']) return '';
    const xml = decode(parts['content.xml']);
    const bodyAt = xml.indexOf('<office:body');
    const body = bodyAt === -1 ? xml : xml.slice(bodyAt);
    const blocks = [];
    for (const paragraph of body.split(/<\/text:(?:p|h)>/)) {
      const text = unescapeXml(stripTags(paragraph));
      if (text.trim()) blocks.push(text);
    }
    return lines(blocks);
  }

  async function fromIdml(bytes) {
    const parts = await readZip(bytes, (n) => /^Stories\/Story_.*\.xml$/.test(n));
    const blocks = [];
    for (const name of Object.keys(parts).sort()) {
      const text = textOf(decode(parts[name]), 'Content').join(' ');
      if (text.trim()) blocks.push(text);
    }
    return lines(blocks);
  }

  /** Bilingual formats are counted on the source side, like an analysis report. */
  function fromXliff(xml) {
    const units = [];
    for (const unit of xml.split(/<\/(?:trans-unit|unit|segment)>/)) {
      const source = textOf(unit, 'source')[0] || textOf(unit, 'seg-source')[0];
      if (source && source.trim()) units.push(source);
    }
    return lines(units);
  }

  function fromTmx(xml) {
    const units = [];
    for (const tu of xml.split('</tu>')) {
      const first = textOf(tu, 'seg')[0]; // the first variant is the source
      if (first && first.trim()) units.push(first);
    }
    return lines(units);
  }

  const fromSubtitles = (text) =>
    lines(
      text
        .split(/\r?\n/)
        .filter((line) => !/^\d+$/.test(line.trim()))
        .filter((line) => !/-->/.test(line))
        .filter((line) => !/^(WEBVTT|NOTE|STYLE)\b/.test(line.trim()))
    );

  function fromPo(text) {
    const out = [];
    const pattern = /^msgid(?:_plural)?\s+((?:"(?:[^"\\]|\\.)*"\s*)+)/gm;
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const value = match[1]
        .split('\n')
        .map((line) => line.trim().replace(/^"|"$/g, ''))
        .join('');
      if (value.trim()) out.push(value);
    }
    return lines(out);
  }

  function fromJson(text) {
    const out = [];
    const walk = (node) => {
      if (typeof node === 'string') {
        if (node.trim()) out.push(node);
      } else if (Array.isArray(node)) {
        node.forEach(walk);
      } else if (node && typeof node === 'object') {
        Object.keys(node).forEach((key) => walk(node[key]));
      }
    };
    try {
      walk(JSON.parse(text));
    } catch (err) {
      return text; // not valid JSON: count it as plain text rather than refuse it
    }
    return lines(out);
  }

  const fromKeyValue = (text) =>
    lines(
      text
        .split(/\r?\n/)
        .filter((line) => line.trim() && !/^\s*[#;!]/.test(line))
        .map((line) => {
          const at = line.search(/[=:]/);
          return at === -1 ? '' : line.slice(at + 1).trim().replace(/^"|";?$/g, '');
        })
    );

  const fromDelimited = (text, separator) =>
    lines(text.split(/\r?\n/).map((row) => row.split(separator).join(' ')));

  /* -------------------------------------------------------------- pages */

  /**
   * How many pages the file really has, for the layout fee. Office files cache
   * the figure themselves; a deck counts its slides; a PDF counts its page
   * objects. Anything else returns 0, meaning "we did not find out".
   */
  async function pageCount(ext, bytes) {
    const re = (source, flags) => new RegExp(source, flags);
    try {
      if (re('^(docx|docm|dotx)$').test(ext)) {
        const parts = await readZip(bytes, (n) => n === 'docProps/app.xml');
        const xml = parts['docProps/app.xml'] ? decode(parts['docProps/app.xml']) : '';
        const found = xml.match(re('<Pages>([0-9]+)</Pages>'));
        return found ? Number(found[1]) : 0;
      }
      if (re('^(pptx|pptm|potx)$').test(ext)) {
        const parts = await readZip(bytes, (n) => re('^ppt/slides/slide[0-9]+[.]xml$').test(n));
        return Object.keys(parts).length;
      }
      if (re('^(odt|ods|odp)$').test(ext)) {
        const parts = await readZip(bytes, (n) => n === 'meta.xml');
        const xml = parts['meta.xml'] ? decode(parts['meta.xml']) : '';
        const found = xml.match(re('meta:page-count="([0-9]+)"'));
        return found ? Number(found[1]) : 0;
      }
      if (ext === 'idml') {
        const parts = await readZip(bytes, (n) => re('^Spreads/Spread_.*[.]xml$').test(n));
        return Object.keys(parts).length;
      }
      if (ext === 'pdf') {
        // Page objects in the raw file: enough for an ordinary exported PDF.
        const raw = new TextDecoder('latin1').decode(bytes);
        const pages = raw.match(re('/Type[ ]*/Page[^s]', 'g'));
        return pages ? pages.length : 0;
      }
    } catch (err) {
      return 0;
    }
    return 0;
  }

  /* --------------------------------------------------------- the type table */

  /** Every type the counter opens, and how it reads it. */
  const TYPES = [
    { ext: ['docx', 'docm', 'dotx'], kind: 'Word document', zip: true, read: fromDocx },
    { ext: ['xlsx', 'xlsm', 'xltx'], kind: 'Excel workbook', zip: true, read: fromXlsx },
    { ext: ['pptx', 'pptm', 'potx'], kind: 'PowerPoint deck', zip: true, read: fromPptx },
    { ext: ['odt'], kind: 'OpenDocument text', zip: true, read: fromOpenDocument },
    { ext: ['ods'], kind: 'OpenDocument spreadsheet', zip: true, read: fromOpenDocument },
    { ext: ['odp'], kind: 'OpenDocument presentation', zip: true, read: fromOpenDocument },
    { ext: ['idml'], kind: 'InDesign (IDML)', zip: true, read: fromIdml },
    { ext: ['xliff', 'xlf', 'sdlxliff', 'mqxliff', 'mxliff'], kind: 'XLIFF (source side)', read: fromXliff },
    { ext: ['tmx'], kind: 'Translation memory (source side)', read: fromTmx },
    { ext: ['srt', 'vtt', 'sbv'], kind: 'Subtitles', read: fromSubtitles },
    { ext: ['po', 'pot'], kind: 'Gettext catalogue', read: fromPo },
    { ext: ['json'], kind: 'JSON', read: fromJson },
    { ext: ['properties', 'strings', 'ini', 'lang'], kind: 'Key/value resources', read: fromKeyValue },
    { ext: ['resx'], kind: '.NET resources', read: (t) => lines(textOf(t, 'value')) },
    { ext: ['html', 'htm', 'xhtml'], kind: 'HTML', read: (t) => visibleText(t, ['script', 'style', 'head']) },
    { ext: ['xml', 'dita', 'ditamap', 'svg', 'rss'], kind: 'XML', read: (t) => visibleText(t, ['script', 'style']) },
    { ext: ['md', 'markdown', 'txt', 'text', 'yaml', 'yml'], kind: 'Plain text', read: (t) => t },
    { ext: ['csv'], kind: 'CSV', read: (t) => fromDelimited(t, ',') },
    { ext: ['tsv'], kind: 'TSV', read: (t) => fromDelimited(t, '\t') }
  ];

  /**
   * Types a quote accepts but the counter cannot open. Each one says how to get
   * an exact count instead; none of them is ever given a guessed word count.
   */
  const ESTIMATED = {
    pdf: 'A PDF cannot be counted reliably. Send the file it was made from (.docx, .idml), or enter the word count.',
    doc: 'Legacy .doc cannot be opened here. Save it as .docx for an exact count.',
    xls: 'Legacy .xls cannot be opened here. Save it as .xlsx for an exact count.',
    ppt: 'Legacy .ppt cannot be opened here. Save it as .pptx for an exact count.',
    rtf: 'RTF cannot be opened here. Save it as .docx for an exact count.',
    indd: 'Native InDesign files cannot be read. Export an IDML for an exact count.',
    pages: 'Apple Pages files cannot be opened here. Export .docx for an exact count.',
    key: 'Apple Keynote files cannot be opened here. Export .pptx for an exact count.',
    numbers: 'Apple Numbers files cannot be opened here. Export .xlsx for an exact count.'
  };

  const extensionOf = (name) => {
    const dot = String(name).lastIndexOf('.');
    return dot === -1 ? '' : String(name).slice(dot + 1).toLowerCase();
  };

  const typeFor = (ext) => TYPES.find((t) => t.ext.indexOf(ext) !== -1) || null;

  /** Every extension the counter accepts, for the file picker. */
  function supportedExtensions() {
    const exact = [];
    for (const type of TYPES) for (const ext of type.ext) exact.push('.' + ext);
    return exact.concat(Object.keys(ESTIMATED).map((e) => '.' + e));
  }

  /**
   * The bytes of a file, however it arrived.
   * A browser File carries `bytes` as a METHOD (Blob.bytes()), while a test
   * fixture carries it as the array itself — check the type, not the truthiness.
   */
  async function bytesOf(file) {
    if (file.bytes instanceof Uint8Array) return file.bytes;
    if (typeof file.arrayBuffer === 'function') return new Uint8Array(await file.arrayBuffer());
    if (typeof file.bytes === 'function') return file.bytes();
    throw new Error('no readable bytes');
  }

  async function textContentOf(file) {
    if (typeof file.text === 'function') return file.text();
    return decode(await bytesOf(file));
  }

  /**
   * A file we cannot open is reported as uncounted — never as a guessed number.
   * The caller asks for the word count, or for the source file.
   */
  const unreadable = (base, kind, note) =>
    Object.assign(base, {
      kind: kind,
      supported: false,
      words: 0,
      characters: 0,
      segments: 0,
      repeatedSegments: 0,
      repeatedWords: 0,
      pages: 0,
      note: note
    });

  /**
   * Analyse one file.
   * @param {{name: string, size: number, bytes?: Uint8Array, arrayBuffer?: function, text?: function}} file
   */
  async function analyseFile(file) {
    const ext = extensionOf(file.name);
    const type = typeFor(ext);
    const base = { file: file.name, extension: ext, size: file.size || 0 };

    if (!type) {
      const known = await pageCount(ext, await bytesOf(file).catch(() => new Uint8Array()));
      if (known > 0) {
        return Object.assign(
          unreadable(base, ESTIMATED[ext] ? ext.toUpperCase() : 'Unknown type', ESTIMATED[ext] || 'We cannot open this file type. Enter the word count, or send us the source file.'),
          { pages: known }
        );
      }
      return unreadable(
        base,
        ESTIMATED[ext] ? ext.toUpperCase() : 'Unknown type',
        ESTIMATED[ext] || 'We cannot open this file type. Enter the word count, or send us the source file.'
      );
    }

    let text = '';
    try {
      text = type.zip ? await type.read(await bytesOf(file)) : type.read(await textContentOf(file));
    } catch (err) {
      return unreadable(base, type.kind, 'This file could not be opened (' + err.message + '). Enter the word count instead.');
    }

    const pages = await pageCount(ext, await bytesOf(file).catch(() => new Uint8Array()));
    const segments = segment(text);
    const reps = repetitions(segments);
    return Object.assign(base, {
      kind: type.kind,
      supported: true,
      words: countWords(text),
      characters: countCharacters(text),
      segments: segments.length,
      repeatedSegments: reps.repeated,
      repeatedWords: reps.repeatedWords,
      pages: pages,
      note: null
    });
  }

  /** Analyse a list of files and total them up. */
  async function analyse(files) {
    const results = [];
    for (const file of files) results.push(await analyseFile(file));
    const total = results.reduce(
      (sum, r) => ({
        words: sum.words + r.words,
        characters: sum.characters + r.characters,
        segments: sum.segments + r.segments,
        repeatedSegments: sum.repeatedSegments + r.repeatedSegments,
        repeatedWords: sum.repeatedWords + r.repeatedWords,
        pages: sum.pages + (r.pages || 0)
      }),
      { words: 0, characters: 0, segments: 0, repeatedSegments: 0, repeatedWords: 0, pages: 0 }
    );
    return {
      files: results,
      total: total,
      exact: results.every((r) => r.supported),
      uncountedFiles: results.filter((r) => !r.supported).map((r) => r.file)
    };
  }

  return {
    analyse: analyse,
    analyseFile: analyseFile,
    countWords: countWords,
    countCharacters: countCharacters,
    segment: segment,
    repetitions: repetitions,
    supportedExtensions: supportedExtensions,
    pageCount: pageCount,
    TYPES: TYPES,
    ESTIMATED: ESTIMATED
  };
});

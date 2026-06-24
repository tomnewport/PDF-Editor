// PDF document operations built on pdf-lib: create/insert/delete pages and
// read/write the document outline (bookmarks) used to name inserted sections.
import {
  PDFDocument,
  PDFName,
  PDFString,
  PDFHexString,
  PDFDict,
  PDFArray,
  PDFRawStream,
  PDFRef,
  StandardFonts,
  rgb,
  decodePDFRawStream,
} from 'pdf-lib';

/** Create an empty PDF (zero pages) and return its bytes. */
export async function createEmptyPdf() {
  const doc = await PDFDocument.create();
  // pdf-lib writes a valid (if unusual) zero-page document.
  return doc.save();
}

/** Load bytes into a pdf-lib document. Always pass a copy so callers keep theirs. */
async function load(bytes) {
  return PDFDocument.load(bytes, { updateMetadata: false, ignoreEncryption: true });
}

/** Number of pages in a PDF given its bytes. */
export async function getPageCount(bytes) {
  const doc = await load(bytes);
  return doc.getPageCount();
}

/**
 * Insert pages copied from `srcBytes` into `destBytes`.
 * @param {Uint8Array} destBytes current document
 * @param {Uint8Array} srcBytes  dropped document
 * @param {number[]} srcPageIndices 0-based indices to copy from src (in order)
 * @param {number} insertAt 0-based index in dest where pages are inserted
 * @param {Array<{title:string,pageIndex:number}>} outline current outline state
 * @param {string} [name] optional bookmark title for the first inserted page
 * @returns {Promise<{bytes:Uint8Array, outline:Array, insertedCount:number, insertAt:number}>}
 */
export async function insertPages(destBytes, srcBytes, srcPageIndices, insertAt, outline, name) {
  // A null/empty destination represents a brand-new document with no pages.
  const dest = destBytes && destBytes.length ? await load(destBytes) : await PDFDocument.create();
  const src = await load(srcBytes);

  const clamped = Math.max(0, Math.min(insertAt, dest.getPageCount()));
  const copied = await dest.copyPages(src, srcPageIndices);
  copied.forEach((page, i) => dest.insertPage(clamped + i, page));

  const count = copied.length;
  // Shift existing bookmarks that sit at/after the insertion point.
  let newOutline = (outline || []).map((e) =>
    e.pageIndex >= clamped ? { ...e, pageIndex: e.pageIndex + count } : { ...e }
  );
  if (name && name.trim()) {
    newOutline.push({ title: name.trim(), pageIndex: clamped });
  }
  newOutline.sort((a, b) => a.pageIndex - b.pageIndex);

  writeOutline(dest, newOutline);
  const bytes = await dest.save();
  return { bytes, outline: newOutline, insertedCount: count, insertAt: clamped };
}

/**
 * Delete pages by 0-based index.
 * @returns {Promise<{bytes:Uint8Array, outline:Array}>}
 */
export async function deletePages(destBytes, indices, outline) {
  const dest = await load(destBytes);
  const toDelete = [...new Set(indices)].filter(
    (i) => i >= 0 && i < dest.getPageCount()
  );
  const descending = [...toDelete].sort((a, b) => b - a);
  descending.forEach((i) => dest.removePage(i));

  const deleted = new Set(toDelete);
  const sortedAsc = [...toDelete].sort((a, b) => a - b);
  const newOutline = (outline || [])
    .filter((e) => !deleted.has(e.pageIndex))
    .map((e) => {
      const below = sortedAsc.filter((d) => d < e.pageIndex).length;
      return { ...e, pageIndex: e.pageIndex - below };
    });

  const remaining = dest.getPageCount();
  writeOutline(dest, newOutline);
  const bytes = await dest.save();
  return { bytes, outline: newOutline, remaining };
}

/** Re-apply the outline to existing bytes (used when only the outline changed). */
export async function applyOutline(destBytes, outline) {
  const dest = await load(destBytes);
  writeOutline(dest, outline);
  return dest.save();
}

/**
 * Permanently redact pages by replacing every affected page with a flattened
 * PNG rendering that already has the black redaction boxes burned in.
 */
export async function permanentlyRedactPages(destBytes, pageImages, outline) {
  const src = await load(destBytes);
  const out = await PDFDocument.create();
  const images = pageImages || [];
  const imagesByPage = new Map(images.map((image) => [image.pageIndex, image]));

  for (let i = 0; i < src.getPageCount(); i++) {
    const image = imagesByPage.get(i);
    if (image) {
      const page = out.addPage([image.width, image.height]);
      const png = await out.embedPng(image.pngBytes);
      page.drawImage(png, {
        x: 0,
        y: 0,
        width: image.width,
        height: image.height,
      });
    } else {
      const [copied] = await out.copyPages(src, [i]);
      out.addPage(copied);
    }
  }

  writeOutline(out, outline);
  return {
    bytes: await out.save(),
    redactedPageCount: images.length,
  };
}

/** Add a sticky-note comment annotation to a page. */
export async function addComment(destBytes, { pageIndex, rect, contents, author }) {
  const doc = await load(destBytes);
  const page = doc.getPage(pageIndex);
  const annots = ensurePageAnnots(doc, page);
  const annotation = createTextAnnotation(doc, {
    rect,
    contents,
    author,
    color: [1, 0.85, 0],
  });
  const ref = doc.context.register(annotation);
  annots.push(ref);
  return {
    bytes: await doc.save(),
    annotationId: normalizeRefId(ref.toString()),
  };
}

/** Add a reply annotation that points at an existing comment. */
export async function replyToComment(destBytes, { annotationId, contents, author }) {
  const doc = await load(destBytes);
  const target = findAnnotationById(doc, annotationId);
  if (!target?.ref) throw new Error('Cannot reply to this comment.');

  const annots = ensurePageAnnots(doc, target.page);
  const parentRect = rectFromPdfArray(lookupPdfArray(target.dict, 'Rect')) || {
    x: 36,
    y: 36,
    width: 22,
    height: 22,
  };
  const annotation = createTextAnnotation(doc, {
    rect: offsetReplyRect(parentRect, target.page),
    contents,
    author,
    color: [0.46, 0.68, 1],
    inReplyTo: target.ref,
    replyType: PDFName.of('R'),
  });
  const ref = doc.context.register(annotation);
  annots.push(ref);
  return {
    bytes: await doc.save(),
    annotationId: normalizeRefId(ref.toString()),
  };
}

/** Remove a single comment, or an entire thread when includeReplies is true. */
export async function removeComment(destBytes, { annotationId, includeReplies = false }) {
  const doc = await load(destBytes);
  const targetId = normalizeRefId(annotationId);
  if (!targetId) throw new Error('Comment not found.');

  let removed = false;
  for (const page of doc.getPages()) {
    const annots = lookupPageAnnots(page);
    if (!annots) continue;
    for (let i = annots.size() - 1; i >= 0; i--) {
      const raw = annots.get(i);
      const dict = raw instanceof PDFRef ? doc.context.lookup(raw) : raw;
      if (!(dict instanceof PDFDict)) continue;

      const id = normalizeRefId(raw?.toString?.());
      const parentId = normalizeRefId(dict.get(PDFName.of('IRT'))?.toString?.());
      if (id === targetId || (includeReplies && parentId === targetId)) {
        annots.remove(i);
        removed = true;
      }
    }
    if (annots.size() === 0) page.node.delete(PDFName.of('Annots'));
  }

  if (!removed) throw new Error('Comment not found.');
  return { bytes: await doc.save() };
}

/** Mark a top-level comment as resolved/reopened using the PDF review state. */
export async function setCommentResolved(destBytes, { annotationId, resolved }) {
  const doc = await load(destBytes);
  const target = findAnnotationById(doc, annotationId);
  if (!target?.dict) throw new Error('Comment not found.');

  const targetId = normalizeRefId(target.ref?.toString?.() || annotationId);
  removeReviewStateAnnotations(doc, targetId);

  if (resolved) {
    target.dict.set(PDFName.of('StateModel'), PDFName.of('Review'));
    target.dict.set(PDFName.of('State'), PDFName.of('Completed'));
  } else {
    target.dict.delete(PDFName.of('StateModel'));
    target.dict.delete(PDFName.of('State'));
  }
  target.dict.set(PDFName.of('M'), PDFString.of(pdfDate()));
  return {
    bytes: await doc.save(),
    annotationId: targetId,
  };
}

/** Best-effort reveal for suspected fake redactions found by the review pane. */
export async function unredactFinding(destBytes, finding) {
  const dest = await load(destBytes);
  const page = dest.getPage(finding.pageIndex);
  let removedAnnotation = false;
  let removedOverlay = false;

  if (finding.kind === 'overlay-redaction') {
    removedAnnotation = removeMatchingAnnotation(dest, page, finding);
    if (!removedAnnotation) {
      removedOverlay = removeOverlayRectangle(dest, page, finding.rect);
    }
  }

  if (!removedAnnotation && !removedOverlay) {
    await drawVisibleReveal(dest, page, finding);
  }

  const bytes = await dest.save();
  return {
    bytes,
    mode: removedAnnotation
      ? 'removed-annotation'
      : removedOverlay
        ? 'removed-overlay'
        : finding.kind === 'hidden-text'
          ? 'revealed-text'
          : 'revealed-overlay',
  };
}

// --- Outline (bookmarks) low-level helpers --------------------------------

/** Replace the document outline with a flat list of entries. */
function writeOutline(doc, entries) {
  const context = doc.context;
  const catalog = doc.catalog;
  catalog.delete(PDFName.of('Outlines'));

  const pageCount = doc.getPageCount();
  const valid = (entries || []).filter(
    (e) => e && e.title && e.pageIndex >= 0 && e.pageIndex < pageCount
  );
  if (valid.length === 0) return;

  const pages = doc.getPages();
  const outlinesRef = context.nextRef();
  const itemRefs = valid.map(() => context.nextRef());

  valid.forEach((entry, i) => {
    const page = pages[entry.pageIndex];
    const dict = context.obj({
      Title: PDFHexString.fromText(entry.title),
      Parent: outlinesRef,
      Dest: [page.ref, PDFName.of('XYZ'), null, null, null],
    });
    if (i > 0) dict.set(PDFName.of('Prev'), itemRefs[i - 1]);
    if (i < valid.length - 1) dict.set(PDFName.of('Next'), itemRefs[i + 1]);
    context.assign(itemRefs[i], dict);
  });

  const outlines = context.obj({
    Type: 'Outlines',
    First: itemRefs[0],
    Last: itemRefs[itemRefs.length - 1],
    Count: valid.length,
  });
  context.assign(outlinesRef, outlines);
  catalog.set(PDFName.of('Outlines'), outlinesRef);
}

function removeOverlayRectangle(doc, page, rect) {
  const contents = page.node.get(PDFName.of('Contents'));
  if (!contents) return false;

  const streams = contentStreamEntries(doc, contents);
  for (const entry of streams) {
    const decoded = decodeStreamBytes(entry.stream);
    if (!decoded || !isAscii(decoded)) continue;

    const text = bytesToString(decoded);
    const stripped = stripRectangleBlock(text, rect);
    if (!stripped.removed) continue;

    const replacement = doc.context.flateStream(stringToBytes(stripped.text));
    if (entry.ref) doc.context.assign(entry.ref, replacement);
    else if (entry.parentArray) entry.parentArray.set(entry.index, replacement);
    else page.node.set(PDFName.of('Contents'), replacement);
    return true;
  }

  return false;
}

function contentStreamEntries(doc, contents) {
  const entries = [];
  const add = (obj, parentArray = null, index = -1) => {
    if (obj instanceof PDFRef) {
      const looked = doc.context.lookup(obj);
      if (looked instanceof PDFArray) {
        for (let i = 0; i < looked.size(); i++) add(looked.get(i), looked, i);
        return;
      }
      if (looked instanceof PDFRawStream) {
        entries.push({ ref: obj, stream: looked, parentArray, index });
      }
      return;
    }
    if (obj instanceof PDFArray) {
      for (let i = 0; i < obj.size(); i++) add(obj.get(i), obj, i);
      return;
    }
    if (obj instanceof PDFRawStream) entries.push({ ref: null, stream: obj, parentArray, index });
  };

  if (contents instanceof PDFArray) {
    for (let i = 0; i < contents.size(); i++) add(contents.get(i), contents, i);
  } else {
    add(contents);
  }
  return entries;
}

function removeMatchingAnnotation(doc, page, finding) {
  const annots = lookupPageAnnots(page);
  if (!annots) return false;

  const targetId = normalizeRefId(finding.annotationId);
  const targetRect = finding.annotationRect || finding.rect;
  let removed = false;

  for (let i = annots.size() - 1; i >= 0; i--) {
    const raw = annots.get(i);
    const annotation = raw instanceof PDFRef ? doc.context.lookup(raw) : raw;
    if (!(annotation instanceof PDFDict)) continue;

    const subtype = annotation.get(PDFName.of('Subtype'))?.toString();
    if (subtype === '/Link') continue;

    const refMatches = targetId && normalizeRefId(raw?.toString?.()) === targetId;
    const rect = rectFromPdfArray(lookupPdfArray(annotation, 'Rect'));
    const rectMatches =
      subtype === '/Square' &&
      rect &&
      targetRect &&
      overlapRatio(rect, targetRect) > 0.65;

    if (!refMatches && !rectMatches) continue;
    annots.remove(i);
    removed = true;
  }

  if (removed && annots.size() === 0) page.node.delete(PDFName.of('Annots'));
  return removed;
}

function createTextAnnotation(
  doc,
  { rect, contents, author, color, inReplyTo = null, replyType = null }
) {
  const normalized = normalizeRectObject(rect);
  const dict = doc.context.obj({
    Type: PDFName.of('Annot'),
    Subtype: PDFName.of('Text'),
    Rect: doc.context.obj([
      normalized.x,
      normalized.y,
      normalized.x + normalized.width,
      normalized.y + normalized.height,
    ]),
    Contents: PDFHexString.fromText(String(contents || '')),
    T: PDFHexString.fromText(String(author || 'Me')),
    M: PDFString.of(pdfDate()),
    NM: PDFHexString.fromText(`pdf-workbench-${Date.now()}-${Math.random().toString(36).slice(2)}`),
    Name: PDFName.of('Comment'),
    C: doc.context.obj(color || [1, 0.85, 0]),
    F: 4,
    Open: false,
  });
  if (inReplyTo) dict.set(PDFName.of('IRT'), inReplyTo);
  if (replyType) dict.set(PDFName.of('RT'), replyType);
  return dict;
}

function ensurePageAnnots(doc, page) {
  const existing = lookupPageAnnots(page);
  if (existing) return existing;
  const annots = doc.context.obj([]);
  page.node.set(PDFName.of('Annots'), annots);
  return annots;
}

function findAnnotationById(doc, annotationId) {
  const targetId = normalizeRefId(annotationId);
  if (!targetId) return null;

  const pages = doc.getPages();
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
    const page = pages[pageIndex];
    const annots = lookupPageAnnots(page);
    if (!annots) continue;

    for (let index = 0; index < annots.size(); index++) {
      const raw = annots.get(index);
      const id = normalizeRefId(raw?.toString?.());
      if (id !== targetId) continue;
      const dict = raw instanceof PDFRef ? doc.context.lookup(raw) : raw;
      if (!(dict instanceof PDFDict)) continue;
      return {
        page,
        pageIndex,
        annots,
        index,
        ref: raw instanceof PDFRef ? raw : null,
        dict,
      };
    }
  }
  return null;
}

function removeReviewStateAnnotations(doc, parentId) {
  if (!parentId) return;
  for (const page of doc.getPages()) {
    const annots = lookupPageAnnots(page);
    if (!annots) continue;
    for (let i = annots.size() - 1; i >= 0; i--) {
      const raw = annots.get(i);
      const dict = raw instanceof PDFRef ? doc.context.lookup(raw) : raw;
      if (!(dict instanceof PDFDict)) continue;

      const replyTo = normalizeRefId(dict.get(PDFName.of('IRT'))?.toString?.());
      const stateModel = dict.get(PDFName.of('StateModel'))?.toString?.();
      if (replyTo === parentId && stateModel === '/Review') annots.remove(i);
    }
    if (annots.size() === 0) page.node.delete(PDFName.of('Annots'));
  }
}

function normalizeRectObject(rect) {
  const x = Number(rect?.x);
  const y = Number(rect?.y);
  const width = Number(rect?.width);
  const height = Number(rect?.height);
  return {
    x: Number.isFinite(x) ? x : 36,
    y: Number.isFinite(y) ? y : 36,
    width: Number.isFinite(width) && width > 0 ? width : 22,
    height: Number.isFinite(height) && height > 0 ? height : 22,
  };
}

function offsetReplyRect(rect, page) {
  const pageSize = page.getSize();
  const size = 22;
  const x = Math.max(12, Math.min(pageSize.width - size - 12, rect.x + 26));
  const y = Math.max(12, Math.min(pageSize.height - size - 12, rect.y - 26));
  return { x, y, width: size, height: size };
}

function pdfDate(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');
  return [
    'D:',
    date.getUTCFullYear(),
    pad(date.getUTCMonth() + 1),
    pad(date.getUTCDate()),
    pad(date.getUTCHours()),
    pad(date.getUTCMinutes()),
    pad(date.getUTCSeconds()),
    'Z',
  ].join('');
}

function lookupPageAnnots(page) {
  try {
    return page.node.lookup(PDFName.of('Annots'), PDFArray);
  } catch {
    return null;
  }
}

function lookupPdfArray(dict, name) {
  try {
    return dict.lookup(PDFName.of(name), PDFArray);
  } catch {
    return null;
  }
}

function rectFromPdfArray(arr) {
  if (!arr || arr.size() < 4) return null;
  const values = [];
  for (let i = 0; i < 4; i++) {
    const value = arr.lookup(i);
    const number = value?.asNumber?.();
    if (!Number.isFinite(number)) return null;
    values.push(number);
  }
  const [x1, y1, x2, y2] = values;
  const x = Math.min(x1, x2);
  const y = Math.min(y1, y2);
  return { x, y, width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) };
}

function normalizeRefId(value) {
  if (!value) return null;
  const match = String(value).match(/(\d+)\s*(?:\d+\s*)?R/i);
  return match ? `${match[1]}R` : null;
}

function overlapRatio(a, b) {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.width, b.x + b.width);
  const y2 = Math.min(a.y + a.height, b.y + b.height);
  if (x2 <= x1 || y2 <= y1) return 0;
  const overlap = (x2 - x1) * (y2 - y1);
  const smaller = Math.max(1, Math.min(a.width * a.height, b.width * b.height));
  return overlap / smaller;
}

function decodeStreamBytes(stream) {
  try {
    return decodePDFRawStream(stream).decode();
  } catch {
    try {
      return stream.getContents();
    } catch {
      return null;
    }
  }
}

function stripRectangleBlock(text, rect) {
  const blockPattern = /(^|\n)\s*q\b[\s\S]*?\bQ\s*(?=\n|$)/g;
  let removed = false;
  const next = text.replace(blockPattern, (block) => {
    if (!blockLooksLikeOverlay(block, rect)) return block;
    removed = true;
    return '\n';
  });
  return { text: next, removed };
}

function blockLooksLikeOverlay(block, rect) {
  if (/\bBT\b/.test(block)) return false;
  if (!/(^|\s)(?:0\s+0\s+0\s+rg|0\s+g)(\s|$)/.test(block)) return false;
  if (!/(^|\s)(?:f|f\*|B|B\*|b|b\*)(\s|$)/.test(block)) return false;
  if (!/\b(?:re|m|l|h)\b/.test(block)) return false;

  const numbers = [...block.matchAll(/[+-]?(?:\d*\.\d+|\d+)/g)].map((m) =>
    Number(m[0])
  );
  return (
    hasCloseNumber(numbers, rect.x) &&
    hasCloseNumber(numbers, rect.y) &&
    hasCloseNumber(numbers, rect.width) &&
    hasCloseNumber(numbers, rect.height)
  );
}

function hasCloseNumber(numbers, target) {
  const tolerance = Math.max(0.25, Math.abs(target) * 0.002);
  return numbers.some((n) => Math.abs(n - target) <= tolerance);
}

async function drawVisibleReveal(doc, page, finding) {
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const rect = finding.rect || finding.textRect;
  const textRect = finding.textRect || rect;
  if (!rect || !textRect || !finding.sampleText) return;

  if (finding.kind === 'overlay-redaction') {
    page.drawRectangle({
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: rect.height,
      color: rgb(1, 1, 1),
      borderWidth: 0,
    });
  }

  const size = Math.max(6, Math.min(24, (textRect.height || rect.height) * 0.82));
  page.drawText(finding.sampleText, {
    x: textRect.x,
    y: textRect.y,
    size,
    font,
    color: rgb(0, 0, 0),
    maxWidth: Math.max(20, rect.width),
  });
}

function isAscii(bytes) {
  return bytes.every((byte) => byte < 128);
}

function bytesToString(bytes) {
  let out = '';
  const chunkSize = 8192;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    out += String.fromCharCode(...bytes.slice(i, i + chunkSize));
  }
  return out;
}

function stringToBytes(text) {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i) & 0xff;
  return bytes;
}

/** Best-effort read of an existing outline into a flat list. Never throws. */
export async function readOutline(destBytes) {
  try {
    const doc = await load(destBytes);
    const catalog = doc.catalog;
    const outlines = catalog.lookup(PDFName.of('Outlines'), PDFDict);
    if (!outlines) return [];

    const refToIndex = new Map();
    doc.getPages().forEach((p, i) => refToIndex.set(p.ref.toString(), i));

    const entries = [];
    walkOutline(doc, outlines.get(PDFName.of('First')), refToIndex, entries, 0);
    entries.sort((a, b) => a.pageIndex - b.pageIndex);
    return entries;
  } catch {
    return [];
  }
}

function walkOutline(doc, ref, refToIndex, out, depth) {
  let cur = ref;
  let guard = 0;
  while (cur && guard++ < 5000 && depth < 32) {
    const item = doc.context.lookup(cur, PDFDict);
    if (!item) break;
    const title = decodeTitle(item.get(PDFName.of('Title')));
    const pageIndex = destToIndex(doc, item.get(PDFName.of('Dest')), item, refToIndex);
    if (title != null && pageIndex != null) out.push({ title, pageIndex });
    // Recurse into children (flattened).
    const first = item.get(PDFName.of('First'));
    if (first) walkOutline(doc, first, refToIndex, out, depth + 1);
    cur = item.get(PDFName.of('Next'));
  }
}

function decodeTitle(obj) {
  if (!obj) return null;
  try {
    if (obj instanceof PDFHexString || obj instanceof PDFString) return obj.decodeText();
  } catch {
    /* ignore */
  }
  return null;
}

function destToIndex(doc, destObj, item, refToIndex) {
  try {
    let dest = destObj;
    // Some bookmarks store an action (/A) with the destination inside.
    if (!dest) {
      const action = doc.context.lookupMaybe?.(item.get(PDFName.of('A')), PDFDict);
      if (action) dest = action.get(PDFName.of('D'));
    }
    if (!dest) return null;
    if (typeof dest.lookup !== 'function' && !(dest instanceof PDFArray)) {
      dest = doc.context.lookup(dest);
    }
    if (dest instanceof PDFArray) {
      const pageRef = dest.get(0);
      if (pageRef && refToIndex.has(pageRef.toString())) {
        return refToIndex.get(pageRef.toString());
      }
    }
  } catch {
    /* ignore */
  }
  return null;
}

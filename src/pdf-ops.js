// PDF document operations built on pdf-lib: create/insert/delete pages and
// read/write the document outline (bookmarks) used to name inserted sections.
import { PDFDocument, PDFName, PDFString, PDFHexString, PDFDict, PDFArray } from 'pdf-lib';

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

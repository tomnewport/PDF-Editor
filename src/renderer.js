// Renderer: UI, thumbnail/preview rendering (pdf.js), selection, drag-and-drop,
// context menus, and the insert dialog. PDF mutations live in ./pdf-ops.
import * as pdfjsLib from 'pdfjs-dist';
import {
  createEmptyPdf,
  insertPages,
  deletePages,
  readOutline,
} from './pdf-ops.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  './pdf.worker.min.mjs',
  import.meta.url
).href;

const THUMB_WIDTH = 138; // CSS px (before devicePixelRatio scaling)

// --- Application state ----------------------------------------------------
const state = {
  bytes: null, // Uint8Array: canonical current document
  pdfjsDoc: null, // pdf.js document proxy for rendering
  numPages: 0,
  currentPage: 0, // 0-based
  selection: new Set(), // 0-based indices
  anchor: null, // anchor index for shift-range selection
  outline: [], // [{title, pageIndex}]
  filePath: null,
  fileName: 'Untitled.pdf',
  dirty: false,
  renderId: 0, // bumped on each reload to cancel stale renders
};

// --- DOM references -------------------------------------------------------
const $ = (id) => document.getElementById(id);
const thumbList = $('thumb-list');
const dropIndicator = $('drop-indicator');
const sidebarEmpty = $('sidebar-empty');
const pageCanvas = $('page-canvas');
const pageStage = $('page-stage');
const viewerEmpty = $('viewer-empty');
const pageIndicator = $('page-indicator');
const viewer = $('viewer');
const sidebar = $('sidebar');
const dragOverlay = $('drag-overlay');
const contextMenu = $('context-menu');
const statusEl = $('status');

// =========================================================================
// Document loading & rendering
// =========================================================================

async function loadBytes(bytes, { path = null, name = null, freshOpen = false } = {}) {
  // Tear down previous pdf.js document.
  if (state.pdfjsDoc) {
    try {
      await state.pdfjsDoc.destroy();
    } catch {
      /* ignore */
    }
    state.pdfjsDoc = null;
  }

  // A null/empty `bytes` represents an empty document (zero pages). True
  // zero-page PDFs cannot be round-tripped (pdf-lib and pdf.js both report a
  // phantom page), so emptiness is tracked as the absence of bytes instead.
  state.bytes = bytes && bytes.length ? bytes : null;
  const renderId = ++state.renderId;

  if (path !== null) state.filePath = path;
  if (name !== null) state.fileName = name;

  if (freshOpen) {
    state.outline = state.bytes ? await readOutline(state.bytes) : [];
    state.selection.clear();
    state.anchor = null;
    state.currentPage = 0;
  }

  // Build the pdf.js document for rendering (only when we have real pages).
  if (!state.bytes) {
    state.numPages = 0;
  } else {
    try {
      state.pdfjsDoc = await pdfjsLib.getDocument({ data: state.bytes.slice() }).promise;
      if (renderId !== state.renderId) return; // superseded
      state.numPages = state.pdfjsDoc.numPages;
    } catch {
      state.numPages = 0;
    }
  }

  // Clamp view/selection to the new page count.
  if (state.numPages === 0) {
    state.currentPage = 0;
    state.selection.clear();
  } else {
    state.currentPage = Math.min(state.currentPage, state.numPages - 1);
    state.selection = new Set(
      [...state.selection].filter((i) => i < state.numPages)
    );
    if (state.selection.size === 0) state.selection.add(state.currentPage);
  }

  renderSidebar();
  renderMainView();
  updateChrome();
}

function renderSidebar() {
  // Remove existing thumbs but keep the drop indicator element.
  [...thumbList.querySelectorAll('.thumb')].forEach((el) => el.remove());

  if (state.numPages === 0) {
    sidebarEmpty.classList.remove('hidden');
    return;
  }
  sidebarEmpty.classList.add('hidden');

  const observer = new IntersectionObserver(
    (entries, obs) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          const el = entry.target;
          obs.unobserve(el);
          renderThumb(Number(el.dataset.index), el);
        }
      });
    },
    { root: sidebar, rootMargin: '200px' }
  );

  for (let i = 0; i < state.numPages; i++) {
    const thumb = document.createElement('div');
    thumb.className = 'thumb';
    thumb.dataset.index = String(i);

    const placeholder = document.createElement('div');
    placeholder.className = 'thumb-placeholder';
    thumb.appendChild(placeholder);

    const label = document.createElement('div');
    label.className = 'thumb-label';
    label.textContent = `Page ${i + 1}`;
    thumb.appendChild(label);

    thumb.addEventListener('click', (e) => onThumbClick(i, e));
    thumb.addEventListener('contextmenu', (e) => onThumbContextMenu(i, e));

    thumbList.appendChild(thumb);
    observer.observe(thumb);
  }

  applySelectionStyles();
}

async function renderThumb(index, el) {
  const renderId = state.renderId;
  if (!state.pdfjsDoc) return;
  try {
    const page = await state.pdfjsDoc.getPage(index + 1);
    if (renderId !== state.renderId) return;
    const dpr = window.devicePixelRatio || 1;
    const base = page.getViewport({ scale: 1 });
    const scale = (THUMB_WIDTH / base.width) * dpr;
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;
    if (renderId !== state.renderId) return;

    const placeholder = el.querySelector('.thumb-placeholder');
    if (placeholder) placeholder.replaceWith(canvas);
    else {
      const existing = el.querySelector('canvas');
      if (existing) existing.replaceWith(canvas);
    }
  } catch {
    /* ignore render errors for a single thumb */
  }
}

async function renderMainView() {
  if (state.numPages === 0 || !state.pdfjsDoc) {
    pageStage.classList.add('hidden');
    viewerEmpty.classList.remove('hidden');
    pageIndicator.classList.add('hidden');
    return;
  }
  pageStage.classList.remove('hidden');
  viewerEmpty.classList.add('hidden');
  pageIndicator.classList.remove('hidden');

  const renderId = state.renderId;
  const pageNum = state.currentPage + 1;
  try {
    const page = await state.pdfjsDoc.getPage(pageNum);
    if (renderId !== state.renderId) return;

    const dpr = window.devicePixelRatio || 1;
    const available = viewer.clientWidth - 64;
    const base = page.getViewport({ scale: 1 });
    const fitScale = Math.min(available / base.width, 1.6);
    const viewport = page.getViewport({ scale: fitScale * dpr });

    pageCanvas.width = Math.ceil(viewport.width);
    pageCanvas.height = Math.ceil(viewport.height);
    pageCanvas.style.width = `${Math.ceil(viewport.width / dpr)}px`;
    pageCanvas.style.height = `${Math.ceil(viewport.height / dpr)}px`;
    const ctx = pageCanvas.getContext('2d');
    ctx.clearRect(0, 0, pageCanvas.width, pageCanvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
  } catch {
    /* ignore */
  }
  pageIndicator.textContent = `Page ${pageNum} of ${state.numPages}`;
}

function updateChrome() {
  const dirtyMark = state.dirty ? ' •' : '';
  const title = `${state.fileName}${dirtyMark} — PDF Editor`;
  window.api.setTitle(title);
  const pages = `${state.numPages} page${state.numPages === 1 ? '' : 's'}`;
  const sel = state.selection.size > 1 ? `, ${state.selection.size} selected` : '';
  statusEl.textContent = `${state.fileName}${dirtyMark} · ${pages}${sel}`;
}

// =========================================================================
// Selection
// =========================================================================

function onThumbClick(index, e) {
  if (e.shiftKey && state.anchor !== null) {
    const [a, b] = [state.anchor, index].sort((x, y) => x - y);
    if (!(e.ctrlKey || e.metaKey)) state.selection.clear();
    for (let i = a; i <= b; i++) state.selection.add(i);
  } else if (e.ctrlKey || e.metaKey || e.altKey) {
    if (state.selection.has(index)) state.selection.delete(index);
    else state.selection.add(index);
    state.anchor = index;
  } else {
    state.selection.clear();
    state.selection.add(index);
    state.anchor = index;
  }

  if (state.selection.size === 0) state.selection.add(index);
  state.currentPage = index;
  applySelectionStyles();
  renderMainView();
  updateChrome();
}

function applySelectionStyles() {
  thumbList.querySelectorAll('.thumb').forEach((el) => {
    const i = Number(el.dataset.index);
    el.classList.toggle('selected', state.selection.has(i));
    el.classList.toggle('current', i === state.currentPage);
  });
  const cur = thumbList.querySelector('.thumb.current');
  if (cur) cur.scrollIntoView({ block: 'nearest' });
}

function selectAll() {
  if (state.numPages === 0) return;
  state.selection = new Set(
    Array.from({ length: state.numPages }, (_, i) => i)
  );
  applySelectionStyles();
  updateChrome();
}

// =========================================================================
// Context menu
// =========================================================================

function onThumbContextMenu(index, e) {
  e.preventDefault();
  if (!state.selection.has(index)) {
    state.selection.clear();
    state.selection.add(index);
    state.anchor = index;
    state.currentPage = index;
    applySelectionStyles();
    renderMainView();
    updateChrome();
  }
  showContextMenu(e.clientX, e.clientY);
}

function showContextMenu(x, y) {
  if (state.numPages === 0) return;
  contextMenu.classList.remove('hidden');
  const rect = contextMenu.getBoundingClientRect();
  const px = Math.min(x, window.innerWidth - rect.width - 4);
  const py = Math.min(y, window.innerHeight - rect.height - 4);
  contextMenu.style.left = `${px}px`;
  contextMenu.style.top = `${py}px`;
}

function hideContextMenu() {
  contextMenu.classList.add('hidden');
}

contextMenu.addEventListener('click', async (e) => {
  const action = e.target.dataset.action;
  hideContextMenu();
  if (action === 'delete') await deleteSelectedPages();
  else if (action === 'select-all') selectAll();
});

document.addEventListener('click', (e) => {
  if (!contextMenu.contains(e.target)) hideContextMenu();
});
window.addEventListener('blur', hideContextMenu);

// Right-click on the main viewer.
viewer.addEventListener('contextmenu', (e) => {
  if (state.numPages === 0) return;
  e.preventDefault();
  if (!state.selection.has(state.currentPage)) {
    state.selection.clear();
    state.selection.add(state.currentPage);
    state.anchor = state.currentPage;
    applySelectionStyles();
    updateChrome();
  }
  showContextMenu(e.clientX, e.clientY);
});

// =========================================================================
// Mutations
// =========================================================================

async function deleteSelectedPages() {
  if (state.selection.size === 0 || state.numPages === 0) return;
  const indices = [...state.selection];
  const minIndex = Math.min(...indices);
  const { bytes, outline, remaining } = await deletePages(
    state.bytes,
    indices,
    state.outline
  );
  state.outline = outline;
  state.dirty = true;
  // Keep the view near where pages were removed.
  state.currentPage = minIndex;
  state.selection = new Set();
  // Removing the last page returns the document to the empty state.
  await loadBytes(remaining === 0 ? null : bytes);
  if (state.numPages > 0) {
    state.currentPage = Math.min(minIndex, state.numPages - 1);
    state.selection.add(state.currentPage);
    state.anchor = state.currentPage;
    applySelectionStyles();
    renderMainView();
    updateChrome();
  }
}

async function doInsert({ srcBytes, pageIndices, insertAt, name }) {
  const { bytes, outline, insertedCount, insertAt: at } = await insertPages(
    state.bytes,
    srcBytes,
    pageIndices,
    insertAt,
    state.outline,
    name
  );
  state.outline = outline;
  state.dirty = true;
  state.currentPage = at;
  await loadBytes(bytes);
  // Select the newly inserted pages.
  if (state.numPages > 0) {
    state.selection = new Set();
    for (let i = 0; i < insertedCount; i++) {
      if (at + i < state.numPages) state.selection.add(at + i);
    }
    state.currentPage = at;
    state.anchor = at;
    applySelectionStyles();
    renderMainView();
    updateChrome();
  }
}

// =========================================================================
// File operations
// =========================================================================

async function newDocument() {
  if (!(await confirmDiscardIfDirty())) return;
  state.filePath = null;
  state.fileName = 'Untitled.pdf';
  state.outline = [];
  state.dirty = false;
  await loadBytes(null, { freshOpen: true });
}

async function openDocument() {
  if (!(await confirmDiscardIfDirty())) return;
  const result = await window.api.openFile();
  if (!result) return;
  state.dirty = false;
  await loadBytes(result.data, {
    path: result.path,
    name: result.name,
    freshOpen: true,
  });
}

async function saveDocument(forceDialog = false) {
  let targetPath = state.filePath;
  if (forceDialog || !targetPath) {
    targetPath = await window.api.saveFileDialog(state.fileName);
    if (!targetPath) return;
  }
  // An empty document is serialized to a valid zero-page PDF only on save.
  const outBytes = state.bytes || (await createEmptyPdf());
  const res = await window.api.writeFile(targetPath, outBytes);
  state.filePath = targetPath;
  if (res && res.name) state.fileName = res.name;
  state.dirty = false;
  updateChrome();
}

async function confirmDiscardIfDirty() {
  if (!state.dirty) return true;
  return window.confirm('You have unsaved changes. Discard them?');
}

// =========================================================================
// Drag & drop of external PDFs
// =========================================================================

let dragDepth = 0;

function dragHasFiles(e) {
  if (!e.dataTransfer) return false;
  return [...e.dataTransfer.types].includes('Files');
}

window.addEventListener('dragenter', (e) => {
  if (!dragHasFiles(e)) return;
  e.preventDefault();
  dragDepth++;
  dragOverlay.classList.remove('hidden');
});

window.addEventListener('dragover', (e) => {
  if (!dragHasFiles(e)) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';

  // When hovering over the sidebar, show an insertion indicator at the
  // nearest gap between thumbnails.
  if (state.numPages > 0 && sidebar.contains(e.target)) {
    const idx = computeSidebarInsertIndex(e.clientY);
    state.sidebarDropIndex = idx;
    showDropIndicator(idx);
  } else {
    state.sidebarDropIndex = null;
    hideDropIndicator();
  }
});

window.addEventListener('dragleave', (e) => {
  if (!dragHasFiles(e)) return;
  dragDepth--;
  if (dragDepth <= 0) {
    dragDepth = 0;
    dragOverlay.classList.add('hidden');
    hideDropIndicator();
    state.sidebarDropIndex = null;
  }
});

window.addEventListener('drop', async (e) => {
  if (!dragHasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  dragOverlay.classList.add('hidden');
  hideDropIndicator();

  const droppedOnSidebar = state.sidebarDropIndex;
  state.sidebarDropIndex = null;

  const files = [...e.dataTransfer.files].filter(
    (f) => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf')
  );
  if (files.length === 0) return;

  // Process files one at a time, each through the insert dialog.
  for (const file of files) {
    const srcBytes = new Uint8Array(await file.arrayBuffer());
    let srcCount;
    try {
      const probe = await pdfjsLib.getDocument({ data: srcBytes.slice() }).promise;
      srcCount = probe.numPages;
      await probe.destroy();
    } catch {
      window.alert(`Could not read "${file.name}" as a PDF.`);
      continue;
    }
    const choice = await openInsertDialog({
      fileName: file.name,
      srcCount,
      dropIndex: droppedOnSidebar,
    });
    if (!choice) continue; // cancelled
    await doInsert({
      srcBytes,
      pageIndices: choice.pageIndices,
      insertAt: choice.insertAt,
      name: choice.name,
    });
  }
});

function computeSidebarInsertIndex(clientY) {
  const thumbs = [...thumbList.querySelectorAll('.thumb')];
  for (let i = 0; i < thumbs.length; i++) {
    const r = thumbs[i].getBoundingClientRect();
    if (clientY < r.top + r.height / 2) return i;
  }
  return thumbs.length;
}

function showDropIndicator(index) {
  const thumbs = [...thumbList.querySelectorAll('.thumb')];
  if (thumbs.length === 0) return;
  let top;
  if (index >= thumbs.length) {
    const last = thumbs[thumbs.length - 1];
    top = last.offsetTop + last.offsetHeight + 2;
  } else {
    top = thumbs[index].offsetTop - 6;
  }
  dropIndicator.style.top = `${top}px`;
  dropIndicator.classList.remove('hidden');
}

function hideDropIndicator() {
  dropIndicator.classList.add('hidden');
}

// =========================================================================
// Insert dialog
// =========================================================================

let dialogResolve = null;

function openInsertDialog({ fileName, srcCount, dropIndex }) {
  return new Promise((resolve) => {
    dialogResolve = resolve;

    $('dlg-title').textContent = 'Add pages';
    $('dlg-subtitle').textContent = `From "${fileName}" — ${srcCount} page${
      srcCount === 1 ? '' : 's'
    }`;
    $('dlg-all-label').textContent = `Entire document (${srcCount} page${
      srcCount === 1 ? '' : 's'
    })`;

    // Reset "which pages".
    document.querySelector('input[name="pages"][value="all"]').checked = true;
    const rangeInput = $('dlg-range');
    rangeInput.value = '';
    rangeInput.disabled = true;
    $('dlg-range-error').classList.add('hidden');
    $('dlg-name').value = '';

    const posFieldset = $('dlg-position-fieldset');
    const dropWrap = $('dlg-pos-drop-wrap');

    if (state.numPages === 0) {
      // Nothing to position relative to.
      posFieldset.classList.add('hidden');
    } else {
      posFieldset.classList.remove('hidden');
      const cur = state.currentPage + 1;
      $('dlg-pos-before-label').textContent = `Before current page (page ${cur})`;
      $('dlg-pos-after-label').textContent = `After current page (page ${cur})`;

      if (dropIndex != null) {
        const before = dropIndex; // page number before the gap (1-based = dropIndex)
        const after = dropIndex + 1;
        const desc =
          dropIndex === 0
            ? 'At drop position (before page 1)'
            : dropIndex >= state.numPages
            ? `At drop position (after page ${state.numPages})`
            : `At drop position (between pages ${before} and ${after})`;
        $('dlg-pos-drop-label').textContent = desc;
        dropWrap.classList.remove('hidden');
        document.querySelector('input[name="position"][value="drop"]').checked = true;
        state.dialogDropIndex = dropIndex;
      } else {
        dropWrap.classList.add('hidden');
        document.querySelector('input[name="position"][value="after"]').checked = true;
        state.dialogDropIndex = null;
      }
    }

    state.dialogSrcCount = srcCount;
    $('dialog-overlay').classList.remove('hidden');
    if (state.numPages > 0) rangeInput.focus();
  });
}

function closeDialog(result) {
  $('dialog-overlay').classList.add('hidden');
  const resolve = dialogResolve;
  dialogResolve = null;
  if (resolve) resolve(result);
}

// Enable range input only when "Selected pages" is chosen.
document.querySelectorAll('input[name="pages"]').forEach((radio) => {
  radio.addEventListener('change', () => {
    const isRange =
      document.querySelector('input[name="pages"]:checked').value === 'range';
    const input = $('dlg-range');
    input.disabled = !isRange;
    if (isRange) input.focus();
  });
});

$('dlg-cancel').addEventListener('click', () => closeDialog(null));

$('dlg-confirm').addEventListener('click', () => {
  const srcCount = state.dialogSrcCount;
  const pagesMode = document.querySelector('input[name="pages"]:checked').value;

  let pageIndices;
  if (pagesMode === 'all') {
    pageIndices = Array.from({ length: srcCount }, (_, i) => i);
  } else {
    pageIndices = parsePageRanges($('dlg-range').value, srcCount);
    if (pageIndices.length === 0) {
      const err = $('dlg-range-error');
      err.textContent = `Enter valid page numbers between 1 and ${srcCount}.`;
      err.classList.remove('hidden');
      return;
    }
  }

  let insertAt;
  if (state.numPages === 0) {
    insertAt = 0;
  } else {
    const pos = document.querySelector('input[name="position"]:checked').value;
    if (pos === 'drop') insertAt = state.dialogDropIndex ?? state.currentPage + 1;
    else if (pos === 'before') insertAt = state.currentPage;
    else insertAt = state.currentPage + 1;
  }

  const name = $('dlg-name').value.trim();
  closeDialog({ pageIndices, insertAt, name });
});

$('dialog-overlay').addEventListener('click', (e) => {
  if (e.target === $('dialog-overlay')) closeDialog(null);
});

function parsePageRanges(text, max) {
  const result = [];
  const seen = new Set();
  const add = (n) => {
    if (n >= 1 && n <= max && !seen.has(n)) {
      seen.add(n);
      result.push(n - 1); // 0-based
    }
  };
  for (const rawToken of text.split(',')) {
    const token = rawToken.trim();
    if (!token) continue;
    const rangeMatch = token.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      let a = parseInt(rangeMatch[1], 10);
      let b = parseInt(rangeMatch[2], 10);
      if (a > b) [a, b] = [b, a];
      for (let n = a; n <= b; n++) add(n);
    } else if (/^\d+$/.test(token)) {
      add(parseInt(token, 10));
    }
  }
  return result;
}

// =========================================================================
// Keyboard & menu wiring
// =========================================================================

document.addEventListener('keydown', (e) => {
  // Don't hijack typing inside dialog inputs.
  const inField =
    e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA';
  const dialogOpen = !$('dialog-overlay').classList.contains('hidden');

  if (dialogOpen) {
    if (e.key === 'Escape') closeDialog(null);
    else if (e.key === 'Enter' && !inField) $('dlg-confirm').click();
    return;
  }

  if (inField) return;

  if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    deleteSelectedPages();
  } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
    e.preventDefault();
    selectAll();
  } else if (e.key === 'ArrowDown' || e.key === 'ArrowRight') {
    if (state.currentPage < state.numPages - 1) navigateTo(state.currentPage + 1);
  } else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') {
    if (state.currentPage > 0) navigateTo(state.currentPage - 1);
  } else if (e.key === 'Escape') {
    hideContextMenu();
  }
});

function navigateTo(index) {
  state.currentPage = index;
  state.selection.clear();
  state.selection.add(index);
  state.anchor = index;
  applySelectionStyles();
  renderMainView();
  updateChrome();
}

// Toolbar buttons.
$('btn-new').addEventListener('click', newDocument);
$('btn-open').addEventListener('click', openDocument);
$('btn-save').addEventListener('click', () => saveDocument(false));
$('btn-save-as').addEventListener('click', () => saveDocument(true));

// Application menu actions from main process.
window.api.onMenuAction((action) => {
  switch (action) {
    case 'new':
      newDocument();
      break;
    case 'open':
      openDocument();
      break;
    case 'save':
      saveDocument(false);
      break;
    case 'save-as':
      saveDocument(true);
      break;
    case 'select-all':
      selectAll();
      break;
    case 'delete':
      deleteSelectedPages();
      break;
  }
});

// Re-render the main page on resize so it keeps fitting the viewer.
let resizeTimer = null;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => renderMainView(), 120);
});

// =========================================================================
// Startup
// =========================================================================
(async function init() {
  await loadBytes(null, { freshOpen: true });
})();

// --- Debug-only end-to-end test harness -----------------------------------
if (window.api.debug) {
  const snapshot = () => ({
    numPages: state.numPages,
    currentPage: state.currentPage,
    selection: [...state.selection].sort((a, b) => a - b),
    outline: state.outline.map((e) => ({ title: e.title, pageIndex: e.pageIndex })),
    thumbCanvases: thumbList.querySelectorAll('.thumb canvas').length,
    mainRendered: pageCanvas.width > 0 && !pageStage.classList.contains('hidden'),
  });
  const b64ToBytes = (b64) =>
    Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  window.__test = {
    snapshot,
    async insert(b64, pageIndices, insertAt, name) {
      await doInsert({ srcBytes: b64ToBytes(b64), pageIndices, insertAt, name });
      await new Promise((r) => setTimeout(r, 300)); // let thumbs/main render
      return snapshot();
    },
    async del(indices) {
      state.selection = new Set(indices);
      await deleteSelectedPages();
      await new Promise((r) => setTimeout(r, 200));
      return snapshot();
    },
  };
}

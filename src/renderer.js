// Renderer: UI, thumbnail/preview rendering (pdf.js), selection, drag-and-drop,
// context menus, and the insert dialog. PDF mutations live in ./pdf-ops.
import * as pdfjsLib from 'pdfjs-dist';
import {
  createEmptyPdf,
  insertPages,
  deletePages,
  readOutline,
  unredactFinding,
  permanentlyRedactPages,
  addComment as addPdfComment,
  replyToComment as replyToPdfComment,
  removeComment as removePdfComment,
  setCommentResolved as setPdfCommentResolved,
} from './pdf-ops.js';
import { analyzeSuspiciousRedactions } from './redaction-review.js';
import { extractTableFromPdf } from './table-extract.js';
import { createXlsxWorkbook } from './xlsx-export.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  './pdf.worker.min.mjs',
  import.meta.url
).href;

const THUMB_WIDTH = 138; // CSS px (before devicePixelRatio scaling)
const DIALOG_THUMB_WIDTH = 140; // CSS px (before devicePixelRatio scaling)
const REDACTION_RENDER_SCALE = 2;
const REVIEW_PREVIEW_TARGET_WIDTH = 260;
const REVIEW_PREVIEW_MAX_SCALE = 4;
const REVIEW_PREVIEW_TIMEOUT_MS = 7000;
const REVIEW_PREVIEW_CONCURRENCY = 2;
const APP_NAME = 'PDF Workbench';
const PDFJS_DOCUMENT_OPTIONS = {
  standardFontDataUrl: new URL('./standard_fonts/', import.meta.url).href,
};
const PDFJS_IMAGE_RESOURCES_PATH = new URL('./images/', import.meta.url).href;
const PDFJS_ANNOTATION_TYPE_NAMES = {
  1: 'TEXT',
  2: 'LINK',
  3: 'FREETEXT',
  4: 'LINE',
  5: 'SQUARE',
  6: 'CIRCLE',
  7: 'POLYGON',
  8: 'POLYLINE',
  9: 'HIGHLIGHT',
  10: 'UNDERLINE',
  11: 'SQUIGGLY',
  12: 'STRIKEOUT',
  13: 'STAMP',
  14: 'CARET',
  15: 'INK',
  16: 'POPUP',
  17: 'FILEATTACHMENT',
  20: 'WIDGET',
  26: 'REDACT',
};
const COMMENT_ANNOTATION_TYPES = new Set([
  'TEXT',
  'FREETEXT',
  'HIGHLIGHT',
  'UNDERLINE',
  'SQUIGGLY',
  'STRIKEOUT',
  'STAMP',
  'INK',
  'CARET',
  'FILEATTACHMENT',
]);
const NON_COMMENT_ANNOTATION_TYPES = new Set(['LINK', 'POPUP', 'WIDGET']);
const COMMENT_TYPE_LABELS = {
  TEXT: 'Sticky note',
  FREETEXT: 'Free text',
  HIGHLIGHT: 'Highlight',
  UNDERLINE: 'Underline',
  SQUIGGLY: 'Squiggly underline',
  STRIKEOUT: 'Strikeout',
  STAMP: 'Stamp',
  INK: 'Ink',
  CARET: 'Caret',
  FILEATTACHMENT: 'File attachment',
  SQUARE: 'Shape note',
  CIRCLE: 'Shape note',
  LINE: 'Line note',
  POLYGON: 'Shape note',
  POLYLINE: 'Shape note',
};

// --- Application state ----------------------------------------------------
let nextDocumentId = 1;

function createDocumentState({ filePath = null, fileName = 'Untitled.pdf' } = {}) {
  return {
    id: nextDocumentId++,
    bytes: null, // Uint8Array: canonical current document
    pdfjsDoc: null, // pdf.js document proxy for rendering
    numPages: 0,
    currentPage: 0, // 0-based
    selection: new Set(), // 0-based indices
    anchor: null, // anchor index for shift-range selection
    outline: [], // [{title, pageIndex}]
    filePath,
    fileName,
    dirty: false,
    renderId: 0, // bumped on each reload to cancel stale renders
    sidebarDropIndex: null,
    previewDropIndex: null,
    dialogDropIndex: null,
    dialogSrcCount: 0,
    fieldObjectsPromise: null,
    hasJSActionsPromise: null,
    reviewFindings: [],
    reviewComments: [],
    ignoredReviewFindingIds: new Set(),
    focusedReviewFindingId: null,
    focusedReviewCommentId: null,
    reviewScanId: 0,
    reviewStatus: 'idle',
    reviewPreviewCache: new Map(),
    reviewPreviewSession: 0,
    activeTool: 'select',
    commentDialogOpen: false,
    pendingRedactions: [],
    pageSizes: [],
    pageSizesPromise: null,
    pageSizesRenderId: null,
  };
}

const workspace = {
  documents: [],
  activeId: null,
};

let state = null;

// --- DOM references -------------------------------------------------------
const $ = (id) => document.getElementById(id);
const thumbList = $('thumb-list');
const dropIndicator = $('drop-indicator');
const sidebarEmpty = $('sidebar-empty');
const pageStack = $('page-stack');
const pageStage = $('page-stage');
const viewerEmpty = $('viewer-empty');
const pageIndicator = $('page-indicator');
const viewer = $('viewer');
const sidebar = $('sidebar');
const dragOverlay = $('drag-overlay');
const contextMenu = $('context-menu');
const statusEl = $('status');
const tabbar = $('tabbar');
const viewerDropIndicator = $('viewer-drop-indicator');
const sidebarScrollIndicator = $('sidebar-scroll-indicator');
const toolbarNew = $('toolbar-new');
const toolbarOpen = $('toolbar-open');
const toolbarSave = $('toolbar-save');
const toolbarTableExport = $('toolbar-table-export');
const toolbarComment = $('toolbar-comment');
const toolbarRedact = $('toolbar-redact');
const reviewSummary = $('review-summary');
const reviewList = $('review-list');
const reviewIgnored = $('review-ignored');
const reviewIgnoredCount = $('review-ignored-count');
const reviewIgnoredList = $('review-ignored-list');
const reviewAddComment = $('review-add-comment');
const reviewRescan = $('review-rescan');
const redactionTools = $('redaction-tools');
const redactionApply = $('redaction-apply');
const redactionCount = $('redaction-count');
const commentDialogOverlay = $('comment-dialog-overlay');
const commentDialogTitle = $('comment-dlg-title');
const commentAuthor = $('comment-author');
const commentText = $('comment-text');
const commentError = $('comment-error');
const commentConfirm = $('comment-confirm');
const commentCancel = $('comment-cancel');
const tableDialogOverlay = $('table-dialog-overlay');
const tableSummary = $('table-summary');
const tableGrid = $('table-grid');
const tableError = $('table-error');
const tableFirstRowHeader = $('table-first-row-header');
const tableAddRow = $('table-add-row');
const tableRemoveRow = $('table-remove-row');
const tableAddColumn = $('table-add-column');
const tableRemoveColumn = $('table-remove-column');
const tableCancel = $('table-cancel');
const tableExport = $('table-export');

let mainPageObserver = null;
let viewerScrollTimer = null;
let suppressViewerScrollSync = false;
let pendingMainScrollTimers = [];
const mainPageRenderPromises = new Map();
let redactionDraft = null;
let commentDialogResolve = null;
let commentPlacementInProgress = false;
let contextCommentTarget = null;
let reviewPreviewObserver = null;
let reviewPreviewActive = 0;
const reviewPreviewQueue = [];
let tableDialogResolve = null;
let tableDraftRows = [];
let selectedTableCell = { row: 0, col: 0 };

const annotationLinkService = {
  externalLinkEnabled: false,
  isInPresentationMode: false,
  pagesCount: 0,
  page: 1,
  rotation: 0,
  addLinkAttributes(link, url) {
    link.href = url;
  },
  getDestinationHash() {
    return '';
  },
  getAnchorUrl(hash) {
    return hash;
  },
  goToDestination() {},
  goToPage() {},
  executeNamedAction() {},
  executeSetOCGState() {},
  cachePageRef() {},
};

const annotationDownloadManager = {
  openOrDownloadData() {},
  downloadData() {},
};

function markDocumentDirty(docState) {
  docState.dirty = true;
  if (docState === state) {
    updateChrome();
    renderTabs();
  }
}

function wireAnnotationStorage(docState) {
  const storage = docState.pdfjsDoc?.annotationStorage;
  if (!storage) return;
  storage.onSetModified = () => markDocumentDirty(docState);
}

function toUint8Array(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  return new Uint8Array(data);
}

async function getCurrentDocumentBytes() {
  if (!state.bytes) return createEmptyPdf();
  const storage = state.pdfjsDoc?.annotationStorage;
  if (state.pdfjsDoc && storage && storage.size > 0) {
    try {
      return toUint8Array(await state.pdfjsDoc.saveDocument());
    } catch {
      return state.bytes;
    }
  }
  return state.bytes;
}

// =========================================================================
// Document loading & rendering
// =========================================================================

async function loadBytes(bytes, { path = null, name = null, freshOpen = false } = {}) {
  const docState = state;
  // Tear down previous pdf.js document.
  if (docState.pdfjsDoc) {
    try {
      await docState.pdfjsDoc.destroy();
    } catch {
      /* ignore */
    }
    docState.pdfjsDoc = null;
  }

  // A null/empty `bytes` represents an empty document (zero pages). True
  // zero-page PDFs cannot be round-tripped (pdf-lib and pdf.js both report a
  // phantom page), so emptiness is tracked as the absence of bytes instead.
  docState.bytes = bytes && bytes.length ? bytes : null;
  const renderId = ++docState.renderId;
  docState.reviewFindings = [];
  docState.reviewComments = [];
  docState.focusedReviewFindingId = null;
  docState.focusedReviewCommentId = null;
  docState.reviewStatus = 'idle';

  if (path !== null) docState.filePath = path;
  if (name !== null) docState.fileName = name;

  if (freshOpen) {
    docState.outline = docState.bytes ? await readOutline(docState.bytes) : [];
    docState.selection.clear();
    docState.anchor = null;
    docState.currentPage = 0;
    docState.activeTool = 'select';
    docState.pendingRedactions = [];
  }

  // Build the pdf.js document for rendering (only when we have real pages).
  if (!docState.bytes) {
    docState.numPages = 0;
    docState.pageSizes = [];
    docState.pageSizesPromise = null;
    docState.pageSizesRenderId = null;
    docState.fieldObjectsPromise = null;
    docState.hasJSActionsPromise = null;
    docState.reviewFindings = [];
    docState.reviewComments = [];
    docState.focusedReviewFindingId = null;
    docState.focusedReviewCommentId = null;
    docState.reviewStatus = 'idle';
    docState.reviewPreviewCache = new Map();
    docState.reviewPreviewSession++;
  } else {
    try {
      docState.pdfjsDoc = await pdfjsLib.getDocument({
        data: docState.bytes.slice(),
        ...PDFJS_DOCUMENT_OPTIONS,
      }).promise;
      if (docState !== state || renderId !== docState.renderId) return; // superseded
      docState.numPages = docState.pdfjsDoc.numPages;
      docState.pageSizes = [];
      docState.pageSizesPromise = null;
      docState.pageSizesRenderId = null;
      docState.reviewPreviewCache = new Map();
      docState.reviewPreviewSession++;
      docState.fieldObjectsPromise = docState.pdfjsDoc.getFieldObjects().catch(() => null);
      docState.hasJSActionsPromise = docState.pdfjsDoc.hasJSActions().catch(() => false);
      wireAnnotationStorage(docState);
    } catch {
      docState.numPages = 0;
      docState.pageSizes = [];
      docState.pageSizesPromise = null;
      docState.pageSizesRenderId = null;
      docState.fieldObjectsPromise = null;
      docState.hasJSActionsPromise = null;
      docState.reviewFindings = [];
      docState.reviewComments = [];
      docState.focusedReviewFindingId = null;
      docState.focusedReviewCommentId = null;
      docState.reviewStatus = 'error';
      docState.reviewPreviewCache = new Map();
      docState.reviewPreviewSession++;
    }
  }

  // Clamp view/selection to the new page count.
  if (docState.numPages === 0) {
    docState.currentPage = 0;
    docState.selection.clear();
  } else {
    docState.currentPage = Math.min(docState.currentPage, docState.numPages - 1);
    docState.selection = new Set(
      [...docState.selection].filter((i) => i < docState.numPages)
    );
    if (docState.selection.size === 0) docState.selection.add(docState.currentPage);
  }

  if (docState !== state) return;
  renderSidebar();
  renderMainView();
  updateChrome();
  renderTabs();
  renderReviewPane();
  renderToolState();
  if (docState.bytes && docState.pdfjsDoc) runReviewScan(docState, renderId);
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
  const docState = state;
  const renderId = docState.renderId;
  if (!docState.pdfjsDoc) return;
  try {
    const page = await docState.pdfjsDoc.getPage(index + 1);
    if (docState !== state || renderId !== docState.renderId) return;
    const dpr = window.devicePixelRatio || 1;
    const base = page.getViewport({ scale: 1 });
    const scale = (THUMB_WIDTH / base.width) * dpr;
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const ctx = canvas.getContext('2d');
    await page.render({
      canvasContext: ctx,
      viewport,
      annotationMode: pdfjsLib.AnnotationMode.ENABLE_STORAGE,
    }).promise;
    if (docState !== state || renderId !== docState.renderId) return;

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
  const docState = state;
  if (docState.numPages === 0 || !docState.pdfjsDoc) {
    resetMainPageRendering();
    pageStage.classList.add('hidden');
    viewerEmpty.classList.remove('hidden');
    pageIndicator.classList.add('hidden');
    return;
  }
  pageStage.classList.remove('hidden');
  viewerEmpty.classList.add('hidden');
  pageIndicator.classList.remove('hidden');

  const renderId = docState.renderId;
  const pageSizes = await ensurePageSizes(docState, renderId);
  if (docState !== state || renderId !== docState.renderId) return;
  resetMainPageRendering();

  for (let i = 0; i < docState.numPages; i++) {
    const metrics = pageDisplayMetrics(pageSizes[i]);
    const pageView = document.createElement('div');
    pageView.className = 'page-view';
    pageView.dataset.index = String(i);

    const paper = document.createElement('div');
    paper.className = 'page-paper';
    paper.style.width = `${metrics.cssWidth}px`;
    paper.style.minHeight = `${metrics.cssHeight}px`;
    const placeholder = document.createElement('div');
    placeholder.className = 'page-placeholder';
    placeholder.textContent = `Page ${i + 1}`;
    placeholder.style.width = `${metrics.cssWidth}px`;
    placeholder.style.height = `${metrics.cssHeight}px`;
    paper.appendChild(placeholder);
    pageView.appendChild(paper);

    const label = document.createElement('div');
    label.className = 'page-view-label';
    label.textContent = `Page ${i + 1}`;
    pageView.appendChild(label);

    pageView.addEventListener('click', (e) => selectPageFromEvent(i, e, { scrollMain: false }));
    pageStack.appendChild(pageView);
  }

  mainPageObserver = new IntersectionObserver(
    (entries, obs) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        const el = entry.target;
        obs.unobserve(el);
        renderMainPage(Number(el.dataset.index), el, docState, renderId);
      });
    },
    { root: viewer, rootMargin: '800px' }
  );

  pageStack.querySelectorAll('.page-view').forEach((el) => {
    mainPageObserver.observe(el);
  });

  applySelectionStyles({ scrollThumb: false });
  updatePageIndicator();
  requestAnimationFrame(() => scrollMainPageIntoView(docState.currentPage, 'auto'));
}

async function ensurePageSizes(docState, renderId = docState.renderId) {
  const hasAllSizes =
    docState.pageSizes.length === docState.numPages &&
    Array.from({ length: docState.numPages }, (_, i) => Boolean(docState.pageSizes[i])).every(
      Boolean
    );
  if (hasAllSizes) {
    return docState.pageSizes;
  }
  if (docState.pageSizesRenderId !== renderId) {
    docState.pageSizesPromise = null;
    docState.pageSizesRenderId = renderId;
  }
  if (!docState.pageSizesPromise) {
    docState.pageSizesPromise = loadPageSizes(docState, renderId);
  }
  return docState.pageSizesPromise;
}

async function loadPageSizes(docState, renderId) {
  const sizes = new Array(docState.numPages);
  let nextIndex = 0;
  const workerCount = Math.min(8, docState.numPages);
  const workers = Array.from({ length: workerCount }, async () => {
    while (nextIndex < docState.numPages) {
      const index = nextIndex++;
      const page = await docState.pdfjsDoc.getPage(index + 1);
      if (renderId !== docState.renderId) return;
      const viewport = page.getViewport({ scale: 1 });
      sizes[index] = {
        width: viewport.width,
        height: viewport.height,
      };
    }
  });

  await Promise.all(workers);
  if (renderId === docState.renderId) {
    docState.pageSizes = sizes;
  }
  return sizes;
}

function pageDisplayMetrics(pageSize) {
  const naturalWidth = Math.max(1, pageSize?.width || 612);
  const naturalHeight = Math.max(1, pageSize?.height || 792);
  const available = Math.max(320, viewer.clientWidth - 96);
  const fitScale = Math.min(available / naturalWidth, 1.35);
  return {
    fitScale,
    cssWidth: Math.ceil(naturalWidth * fitScale),
    cssHeight: Math.ceil(naturalHeight * fitScale),
  };
}

function resetMainPageRendering() {
  clearPendingMainScrollTimers();
  if (mainPageObserver) {
    mainPageObserver.disconnect();
    mainPageObserver = null;
  }
  pageStack.replaceChildren();
}

async function renderMainPage(index, el, docState, renderId) {
  if (!docState.pdfjsDoc) return;
  try {
    const page = await docState.pdfjsDoc.getPage(index + 1);
    if (docState !== state || renderId !== docState.renderId) return;

    const dpr = window.devicePixelRatio || 1;
    const base = page.getViewport({ scale: 1 });
    const { fitScale, cssWidth, cssHeight } = pageDisplayMetrics(base);
    const viewport = page.getViewport({ scale: fitScale * dpr });
    const annotationViewport = page.getViewport({ scale: fitScale });

    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    canvas.style.width = `${cssWidth}px`;
    canvas.style.height = `${cssHeight}px`;
    const ctx = canvas.getContext('2d');
    await page.render({
      canvasContext: ctx,
      viewport,
      annotationMode: pdfjsLib.AnnotationMode.ENABLE_STORAGE,
    }).promise;
    if (docState !== state || renderId !== docState.renderId) return;

    const paper = el.querySelector('.page-paper');
    paper.style.width = `${cssWidth}px`;
    paper.style.minHeight = `${cssHeight}px`;
    paper.replaceChildren(canvas);
    await renderMainAnnotationLayer(page, paper, annotationViewport, docState, renderId, index, el);
    renderReviewHighlights(paper, index, annotationViewport);
    paper.__pdfViewport = annotationViewport;
    renderRedactionLayer(paper, index, annotationViewport);
    renderCommentLayer(paper, index, annotationViewport);
	    el.classList.add('rendered');
  } catch {
    /* ignore */
  }
}

async function ensureMainPageRendered(index) {
  const docState = state;
  const renderId = docState.renderId;
  const key = `${docState.id}:${renderId}:${index}`;
  if (mainPageRenderPromises.has(key)) return mainPageRenderPromises.get(key);

  const promise = (async () => {
    const el = await waitForMainPageView(index);
    if (!el || docState !== state || renderId !== docState.renderId) return false;
    if (el.classList.contains('rendered')) return true;

    mainPageObserver?.unobserve(el);
    await renderMainPage(index, el, docState, renderId);
    return el.classList.contains('rendered');
  })();

  mainPageRenderPromises.set(key, promise);
  try {
    return await promise;
  } finally {
    mainPageRenderPromises.delete(key);
  }
}

async function waitForMainPageView(index) {
  for (const delay of [0, 16, 50, 100, 200]) {
    const el = pageStack.querySelector(`.page-view[data-index="${index}"]`);
    if (el) return el;
    await sleep(delay);
  }
  return null;
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

async function renderMainAnnotationLayer(page, paper, viewport, docState, renderId, pageIndex, pageView) {
  const annotations = await page.getAnnotations({ intent: 'display' });
  if (
    annotations.length === 0 ||
    docState !== state ||
    renderId !== docState.renderId
  ) {
    return;
  }

  const [fieldObjects, hasJSActions] = await Promise.all([
    docState.fieldObjectsPromise || Promise.resolve(null),
    docState.hasJSActionsPromise || Promise.resolve(false),
  ]);
  if (docState !== state || renderId !== docState.renderId) return;

  const annotationDiv = document.createElement('div');
  annotationDiv.className = 'annotationLayer';
  annotationDiv.style.width = `${Math.ceil(viewport.width)}px`;
  annotationDiv.style.height = `${Math.ceil(viewport.height)}px`;
  annotationDiv.style.setProperty('--scale-factor', String(viewport.scale));
  paper.appendChild(annotationDiv);

  const annotationLayer = new pdfjsLib.AnnotationLayer({
    div: annotationDiv,
    page,
    viewport: viewport.clone({ dontFlip: true }),
    accessibilityManager: null,
    annotationCanvasMap: new Map(),
    annotationEditorUIManager: null,
    structTreeLayer: null,
  });

  await annotationLayer.render({
    annotations,
    viewport: viewport.clone({ dontFlip: true }),
    div: annotationDiv,
    page,
    linkService: annotationLinkService,
    downloadManager: annotationDownloadManager,
    annotationStorage: docState.pdfjsDoc.annotationStorage,
    imageResourcesPath: PDFJS_IMAGE_RESOURCES_PATH,
    renderForms: true,
    enableScripting: false,
    hasJSActions,
    fieldObjects,
  });
  prepareAnnotationControls(annotationDiv, docState, pageIndex, pageView, renderId);
}

function prepareAnnotationControls(annotationDiv, docState, pageIndex, pageView, renderId) {
  const controls = annotationDiv.querySelectorAll('input, textarea, select');
  let rerenderTimer = null;

  const scheduleRerender = () => {
    if (rerenderTimer) window.clearTimeout(rerenderTimer);
    rerenderTimer = window.setTimeout(() => {
      rerenderTimer = null;
      if (
        docState !== state ||
        renderId !== docState.renderId ||
        !pageView.isConnected ||
        annotationDiv.contains(document.activeElement) ||
        !annotationDiv.querySelector('[data-pending-render="true"]')
      ) {
        return;
      }
      renderMainPage(pageIndex, pageView, docState, renderId);
    }, 80);
  };

  controls.forEach((control) => {
    const container = control.closest('section');
    if (control.classList.contains('comb')) {
      const maxLength = Number(control.maxLength || control.getAttribute('maxlength'));
      if (Number.isFinite(maxLength) && maxLength > 0) {
        control.style.setProperty('--comb-columns', String(maxLength));
      }
    }

    const markPending = (target) => {
      target.dataset.pendingRender = 'true';
      const targetContainer = target.closest('section');
      if (targetContainer) targetContainer.dataset.pendingRender = 'true';
    };
    const markEdited = () => {
      if (control.type === 'radio' && control.name) {
        controls.forEach((candidate) => {
          if (candidate.type === 'radio' && candidate.name === control.name) {
            markPending(candidate);
          }
        });
        return;
      }
      markPending(control);
    };
    control.addEventListener('input', markEdited);
    control.addEventListener('change', () => {
      markEdited();
      scheduleRerender();
    });
    control.addEventListener('blur', scheduleRerender);
    control.addEventListener('resetform', () => {
      markEdited();
      scheduleRerender();
    });
  });
}

function updatePageIndicator() {
  if (state.numPages === 0) return;
  pageIndicator.textContent = `Page ${state.currentPage + 1} of ${state.numPages}`;
}

function scrollMainPageIntoView(index, behavior = 'smooth', { stabilize = false } = {}) {
  clearPendingMainScrollTimers();
  const el = pageStack.querySelector(`.page-view[data-index="${index}"]`);
  if (!el) return;
  suppressViewerScrollSync = true;
  el.scrollIntoView({ block: 'start', inline: 'nearest', behavior });
  const releaseDelay = stabilize ? 900 : behavior === 'smooth' ? 320 : 80;
  if (stabilize) {
    for (const delay of [40, 120, 260, 520]) {
      pendingMainScrollTimers.push(
        window.setTimeout(() => {
          const target = pageStack.querySelector(`.page-view[data-index="${index}"]`);
          if (target) target.scrollIntoView({ block: 'start', inline: 'nearest', behavior: 'auto' });
        }, delay)
      );
    }
  }
  pendingMainScrollTimers.push(window.setTimeout(() => {
    suppressViewerScrollSync = false;
  }, releaseDelay));
}

function clearPendingMainScrollTimers() {
  pendingMainScrollTimers.forEach((timer) => window.clearTimeout(timer));
  pendingMainScrollTimers = [];
}

function updateCurrentPageFromViewerScroll() {
  if (suppressViewerScrollSync || state.numPages === 0) return;
  const pages = [...pageStack.querySelectorAll('.page-view')];
  if (pages.length === 0) return;

  const viewerRect = viewer.getBoundingClientRect();
  const centerY = viewerRect.top + viewerRect.height * 0.38;
  let best = null;
  let bestDistance = Infinity;
  for (const page of pages) {
    const rect = page.getBoundingClientRect();
    const pageCenter = rect.top + rect.height / 2;
    const distance = Math.abs(pageCenter - centerY);
    if (distance < bestDistance) {
      bestDistance = distance;
      best = Number(page.dataset.index);
    }
  }

  if (best == null || best === state.currentPage) return;
  state.currentPage = best;
  applySelectionStyles({ scrollThumb: false });
  updatePageIndicator();
}

function updateChrome() {
  const dirtyMark = state.dirty ? ' •' : '';
  const title = `${state.fileName}${dirtyMark} — ${APP_NAME}`;
  window.api.setTitle(title);
  const pages = `${state.numPages} page${state.numPages === 1 ? '' : 's'}`;
  const sel = state.selection.size > 1 ? `, ${state.selection.size} selected` : '';
  statusEl.textContent =
    state.activeTool === 'comment'
      ? 'Click a page to place a comment'
      : `${state.fileName}${dirtyMark} · ${pages}${sel}`;
}

// =========================================================================
// Review pane
// =========================================================================

async function collectPdfComments(pdfjsDoc) {
  const entries = [];
  for (let pageIndex = 0; pageIndex < pdfjsDoc.numPages; pageIndex++) {
    let annotations = [];
    try {
      const page = await pdfjsDoc.getPage(pageIndex + 1);
      annotations = await page.getAnnotations({ intent: 'display' });
    } catch {
      continue;
    }

    annotations.forEach((annotation, annotationIndex) => {
      if (!isPdfCommentAnnotation(annotation)) return;
      const names = annotationTypeNames(annotation);
      const type = names.find((name) => !NON_COMMENT_ANNOTATION_TYPES.has(name)) || 'ANNOTATION';
      const contents = annotationText(annotation.contentsObj) || annotationText(annotation.contents);
      const title = annotationText(annotation.titleObj) || annotationText(annotation.title);
      const rect = rectFromAnnotation(annotation.rect);
      const stateName = pdfNameText(annotation.state);
      const stateModel = pdfNameText(annotation.stateModel);
      const parentId = normalizeAnnotationId(annotation.inReplyTo);
      const replyType = pdfNameText(annotation.replyType);
      entries.push({
        id: `comment:${pageIndex}:${annotation.id || annotation.ref || annotationIndex}`,
        annotationId: normalizeAnnotationId(annotation.id || annotation.ref) || null,
        parentId,
        replyType,
        stateName,
        stateModel,
        pageIndex,
        rect,
        type,
        label: COMMENT_TYPE_LABELS[type] || prettyAnnotationType(type),
        title,
        contents,
        modified: formatPdfDate(annotation.modificationDate),
        color: annotationColor(annotation.color),
        replies: [],
        resolved: stateModel === 'Review' && stateName === 'Completed',
      });
    });
  }

  const byAnnotationId = new Map(
    entries.filter((entry) => entry.annotationId).map((entry) => [entry.annotationId, entry])
  );
  const stateByParent = new Map();
  const repliesByParent = new Map();
  const topLevel = [];

  for (const entry of entries) {
    const isReviewState =
      entry.parentId &&
      entry.replyType === 'Group' &&
      entry.stateModel === 'Review' &&
      entry.stateName;
    if (isReviewState) {
      if (!stateByParent.has(entry.parentId)) stateByParent.set(entry.parentId, []);
      stateByParent.get(entry.parentId).push(entry);
      continue;
    }
    if (entry.parentId && byAnnotationId.has(entry.parentId)) {
      if (!repliesByParent.has(entry.parentId)) repliesByParent.set(entry.parentId, []);
      repliesByParent.get(entry.parentId).push(entry);
      continue;
    }
    topLevel.push(entry);
  }

  for (const comment of topLevel) {
    const replies = repliesByParent.get(comment.annotationId) || [];
    replies.sort(compareCommentEntries);
    comment.replies = replies;

    const states = stateByParent.get(comment.annotationId) || [];
    states.sort(compareCommentEntries);
    const latestState = states.at(-1);
    if (latestState) comment.resolved = latestState.stateName === 'Completed';
  }

  topLevel.sort(compareCommentEntries);
  return topLevel;
}

function isPdfCommentAnnotation(annotation) {
  const names = annotationTypeNames(annotation);
  if (names.some((name) => NON_COMMENT_ANNOTATION_TYPES.has(name))) return false;
  const contents = annotationText(annotation.contentsObj) || annotationText(annotation.contents);
  const title = annotationText(annotation.titleObj) || annotationText(annotation.title);
  if (contents || title || annotation.hasPopup) return true;
  return names.some((name) => COMMENT_ANNOTATION_TYPES.has(name));
}

function annotationTypeNames(annotation) {
  const names = [];
  const fromEnum = annotationTypeName(annotation.annotationType);
  if (fromEnum) names.push(fromEnum);
  const subtype = normalizeAnnotationType(annotation.subtype || annotation.name || '');
  if (subtype) names.push(subtype);
  return [...new Set(names)];
}

function annotationTypeName(type) {
  if (!Number.isFinite(type)) return '';
  return PDFJS_ANNOTATION_TYPE_NAMES[type] || '';
}

function normalizeAnnotationType(value) {
  return String(value || '').replace(/[^a-z0-9]/gi, '').toUpperCase();
}

function annotationText(value) {
  if (value == null) return '';
  if (typeof value === 'string') return cleanAnnotationText(value);
  if (typeof value === 'object') {
    if (typeof value.str === 'string') return cleanAnnotationText(value.str);
    if (typeof value.value === 'string') return cleanAnnotationText(value.value);
  }
  return cleanAnnotationText(String(value));
}

function pdfNameText(value) {
  if (!value) return '';
  if (typeof value === 'string') return value.replace(/^\//, '');
  if (typeof value === 'object' && typeof value.name === 'string') return value.name;
  return String(value).replace(/^\//, '');
}

function normalizeAnnotationId(value) {
  if (!value) return null;
  const match = String(value).match(/(\d+)\s*(?:\d+\s*)?R/i);
  return match ? `${match[1]}R` : String(value);
}

function compareCommentEntries(a, b) {
  if (a.pageIndex !== b.pageIndex) return a.pageIndex - b.pageIndex;
  return String(a.modified || '').localeCompare(String(b.modified || ''));
}

function cleanAnnotationText(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

function rectFromAnnotation(rect) {
  if (!Array.isArray(rect) || rect.length < 4) return null;
  const values = rect.slice(0, 4).map(Number);
  if (values.some((value) => !Number.isFinite(value))) return null;
  const [x1, y1, x2, y2] = values;
  return {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1),
  };
}

function annotationColor(color) {
  if (!color || typeof color.length !== 'number' || color.length < 3) return null;
  const channels = Array.from(color).slice(0, 3);
  if (channels.some((value) => !Number.isFinite(value))) return null;
  const normalized = channels.map((value) =>
    Math.max(0, Math.min(255, value <= 1 ? Math.round(value * 255) : Math.round(value)))
  );
  return `rgb(${normalized.join(', ')})`;
}

function prettyAnnotationType(type) {
  const text = String(type || 'Annotation').toLowerCase().replace(/_/g, ' ');
  return text.replace(/\b\w/g, (ch) => ch.toUpperCase());
}

function formatPdfDate(value) {
  if (!value) return '';
  const text = String(value);
  const match = text.match(/^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?/);
  if (!match) return text;
  const [, year, month = '01', day = '01', hour, minute] = match;
  const date = `${year}-${month}-${day}`;
  return hour ? `${date} ${hour}:${minute || '00'}` : date;
}

async function runReviewScan(docState = state, renderId = docState.renderId) {
  if (!docState.pdfjsDoc) {
    docState.reviewFindings = [];
    docState.reviewComments = [];
    docState.reviewStatus = 'idle';
    if (docState === state) renderReviewPane();
    return;
  }

  const scanId = ++docState.reviewScanId;
  docState.reviewStatus = 'scanning';
  if (docState === state) renderReviewPane();

  try {
    const [findings, comments] = await Promise.all([
      analyzeSuspiciousRedactions(docState.pdfjsDoc),
      collectPdfComments(docState.pdfjsDoc),
    ]);
    if (scanId !== docState.reviewScanId || renderId !== docState.renderId) return;
    docState.reviewFindings = findings;
    docState.reviewComments = comments;
    docState.reviewPreviewCache = new Map();
    docState.reviewPreviewSession++;
    docState.reviewStatus = 'ready';
    if (docState === state) {
      renderReviewPane();
      renderAllReviewHighlights();
    }
  } catch {
    if (scanId !== docState.reviewScanId || renderId !== docState.renderId) return;
    docState.reviewFindings = [];
    docState.reviewComments = [];
    docState.reviewStatus = 'error';
    if (docState === state) renderReviewPane();
  }
}

function renderReviewPane() {
  if (!state) return;
  resetReviewPreviewObserver();
  clearReviewPreviewQueue();
  const { active, ignored } = partitionReviewFindings();
  const comments = state.reviewComments || [];

  reviewAddComment.disabled = !state.numPages || state.reviewStatus === 'scanning';
  reviewAddComment.classList.toggle('active', state.activeTool === 'comment');

  if (state.reviewStatus === 'scanning') {
    reviewSummary.textContent = 'Scanning redactions and comments';
  } else if (state.reviewStatus === 'error') {
    reviewSummary.textContent = 'Review scan failed for this PDF';
  } else if (active.length === 0 && ignored.length === 0 && comments.length === 0) {
    reviewSummary.textContent = state.numPages ? 'No redactions or comments found' : 'No document pages';
  } else {
    reviewSummary.textContent = reviewSummaryText(active.length, comments.length, ignored.length);
  }

  reviewList.replaceChildren();
  if (state.reviewStatus === 'scanning') {
    reviewList.appendChild(reviewEmpty('Scanning PDF content...'));
  } else if (active.length === 0 && comments.length === 0) {
    reviewList.appendChild(
      reviewEmpty(
        state.numPages
          ? 'Nothing active. Ignored redaction findings stay collapsed below.'
          : 'Open or build a PDF to review it.'
      )
    );
  } else {
    if (active.length > 0) {
      reviewList.appendChild(reviewSectionTitle('Suspected redactions', active.length));
      active.forEach((finding) => reviewList.appendChild(reviewCard(finding)));
    }
    if (comments.length > 0) {
      reviewList.appendChild(reviewSectionTitle('Comments', comments.length));
      comments.forEach((comment) => reviewList.appendChild(commentCard(comment)));
    }
  }

  reviewIgnoredCount.textContent = String(ignored.length);
  reviewIgnoredList.replaceChildren();
  ignored.forEach((finding) =>
    reviewIgnoredList.appendChild(reviewCard(finding, { ignored: true }))
  );
  if (ignored.length === 0) reviewIgnored.open = false;
  observeReviewPreviews();
}

function reviewSummaryText(activeCount, commentCount, ignoredCount) {
  const parts = [];
  if (activeCount > 0) parts.push(`${activeCount} redaction ${activeCount === 1 ? 'issue' : 'issues'}`);
  if (commentCount > 0) {
    const replyCount = (state.reviewComments || []).reduce(
      (total, comment) => total + (comment.replies?.length || 0),
      0
    );
    parts.push(
      `${commentCount} ${commentCount === 1 ? 'comment' : 'comments'}${
        replyCount ? `, ${replyCount} ${replyCount === 1 ? 'reply' : 'replies'}` : ''
      }`
    );
  }
  if (ignoredCount > 0) parts.push(`${ignoredCount} ignored`);
  return parts.join(', ');
}

function partitionReviewFindings() {
  const active = [];
  const ignored = [];
  for (const finding of state.reviewFindings) {
    if (state.ignoredReviewFindingIds.has(finding.id)) ignored.push(finding);
    else active.push(finding);
  }
  return { active, ignored };
}

function reviewEmpty(text) {
  const el = document.createElement('div');
  el.className = 'review-empty';
  el.textContent = text;
  return el;
}

function reviewSectionTitle(label, count) {
  const section = document.createElement('div');
  section.className = 'review-section-title';
  const title = document.createElement('span');
  title.textContent = label;
  section.appendChild(title);
  const badge = document.createElement('span');
  badge.textContent = String(count);
  section.appendChild(badge);
  return section;
}

function reviewCard(finding, { ignored = false } = {}) {
  const card = document.createElement('article');
  card.className = 'review-card';
  card.classList.toggle('focused', state.focusedReviewFindingId === finding.id);
  card.dataset.id = finding.id;

  const title = document.createElement('h3');
  title.textContent = finding.title;
  card.appendChild(title);

  const meta = document.createElement('div');
  meta.className = 'review-meta';
  const badge = document.createElement('span');
  badge.className = `review-badge ${finding.severity === 'high' ? 'high' : ''}`;
  badge.textContent = finding.severity;
  meta.appendChild(badge);
  const page = document.createElement('span');
  page.textContent = `Page ${finding.pageIndex + 1}`;
  meta.appendChild(page);
  card.appendChild(meta);

  card.appendChild(reviewPreview(finding));

  const reason = document.createElement('p');
  reason.className = 'review-reason';
  reason.textContent = finding.reason;
  card.appendChild(reason);

  const actions = document.createElement('div');
  actions.className = 'review-actions';
  actions.appendChild(reviewButton('Show', 'show'));
  if (ignored) actions.appendChild(reviewButton('Unignore', 'unignore'));
  else {
    if (finding.kind === 'overlay-redaction' || finding.kind === 'hidden-text') {
      actions.appendChild(reviewButton('Redact Properly', 'redact-properly', 'dangerish'));
    }
    if (finding.repairable) {
      actions.appendChild(reviewButton(finding.repairLabel || 'Unredact', 'unredact', 'dangerish'));
    }
    actions.appendChild(reviewButton('Ignore', 'ignore'));
  }
  card.appendChild(actions);

  return card;
}

function commentCard(comment) {
  const card = document.createElement('article');
  card.className = 'review-card comment-card';
  card.classList.toggle('resolved', Boolean(comment.resolved));
  card.classList.toggle('focused', state.focusedReviewCommentId === comment.id);
  card.dataset.id = comment.id;

  const title = document.createElement('h3');
  title.textContent = comment.label || 'Comment';
  card.appendChild(title);

  const meta = document.createElement('div');
  meta.className = 'review-meta';
  const badge = document.createElement('span');
  badge.className = 'review-badge comment';
  badge.textContent = comment.type || 'COMMENT';
  if (comment.color) badge.style.setProperty('--comment-color', comment.color);
  meta.appendChild(badge);

  const page = document.createElement('span');
  page.textContent = `Page ${comment.pageIndex + 1}`;
  meta.appendChild(page);
  if (comment.title) {
    const author = document.createElement('span');
    author.textContent = comment.title;
    meta.appendChild(author);
  }
  if (comment.modified) {
    const modified = document.createElement('span');
    modified.textContent = comment.modified;
    meta.appendChild(modified);
  }
  if (comment.resolved) {
    const resolved = document.createElement('span');
    resolved.className = 'review-status resolved';
    resolved.textContent = 'Resolved';
    meta.appendChild(resolved);
  }
  card.appendChild(meta);

  const body = document.createElement('p');
  body.className = 'review-comment-content';
  body.textContent = comment.contents || 'No comment text attached to this annotation.';
  card.appendChild(body);

  if (comment.replies?.length) {
    const replies = document.createElement('div');
    replies.className = 'comment-replies';
    for (const reply of comment.replies) {
      replies.appendChild(commentReply(reply));
    }
    card.appendChild(replies);
  }

  const actions = document.createElement('div');
  actions.className = 'review-actions';
  actions.appendChild(reviewButton('Show', 'show-comment'));
  actions.appendChild(reviewButton('Reply', 'reply-comment'));
  actions.appendChild(
    reviewButton(comment.resolved ? 'Reopen' : 'Resolve', 'toggle-resolved')
  );
  actions.appendChild(reviewButton('Remove', 'remove-comment', 'dangerish'));
  card.appendChild(actions);

  return card;
}

function commentReply(reply) {
  const row = document.createElement('article');
  row.className = 'comment-reply';
  row.classList.toggle('focused', state.focusedReviewCommentId === reply.id);
  row.dataset.id = reply.id;

  const meta = document.createElement('div');
  meta.className = 'review-meta';
  const badge = document.createElement('span');
  badge.className = 'review-badge comment reply';
  badge.textContent = 'REPLY';
  if (reply.color) badge.style.setProperty('--comment-color', reply.color);
  meta.appendChild(badge);
  if (reply.title) {
    const author = document.createElement('span');
    author.textContent = reply.title;
    meta.appendChild(author);
  }
  if (reply.modified) {
    const modified = document.createElement('span');
    modified.textContent = reply.modified;
    meta.appendChild(modified);
  }
  row.appendChild(meta);

  const body = document.createElement('p');
  body.className = 'review-comment-content';
  body.textContent = reply.contents || 'No reply text.';
  row.appendChild(body);

  const actions = document.createElement('div');
  actions.className = 'review-actions';
  actions.appendChild(reviewButton('Show', 'show-comment'));
  actions.appendChild(reviewButton('Remove', 'remove-reply', 'dangerish'));
  row.appendChild(actions);
  return row;
}

function reviewPreview(finding) {
  const preview = document.createElement('div');
  preview.className = 'review-preview';
  preview.dataset.findingId = finding.id;
  preview.appendChild(reviewPreviewPanel('As-is', 'redacted'));
  preview.appendChild(reviewPreviewPanel('Revealed', 'revealed'));

  const docState = state;
  const cached = docState.reviewPreviewCache.get(finding.id);
  if (cached) {
    applyReviewPreviewImages(preview, cached);
  } else {
    preview.dataset.needsPreview = 'true';
  }

  return preview;
}

function reviewPreviewPanel(label, side) {
  const pane = document.createElement('div');
  pane.className = 'review-preview-pane';
  pane.dataset.side = side;

  const caption = document.createElement('div');
  caption.className = 'review-preview-label';
  caption.textContent = label;
  pane.appendChild(caption);

  const frame = document.createElement('div');
  frame.className = 'review-preview-frame loading';
  frame.textContent = 'Queued';
  pane.appendChild(frame);

  return pane;
}

function applyReviewPreviewImages(preview, result) {
  setReviewPreviewFrame(
    preview.querySelector('[data-side="redacted"] .review-preview-frame'),
    result.redactedUrl,
    result.redactedError || 'As-is preview unavailable',
    'As-is PDF crop'
  );
  setReviewPreviewFrame(
    preview.querySelector('[data-side="revealed"] .review-preview-frame'),
    result.revealedUrl,
    result.revealedError || 'Hidden content not recoverable',
    'Revealed PDF crop'
  );
}

function setReviewPreviewFrame(frame, imageUrl, fallback, altText) {
  if (!frame) return;
  frame.replaceChildren();
  frame.classList.remove('loading', 'empty');
  if (imageUrl) {
    const img = document.createElement('img');
    img.src = imageUrl;
    img.alt = altText;
    frame.appendChild(img);
  } else {
    frame.classList.add('empty');
    frame.textContent = fallback;
  }
}

function resetReviewPreviewObserver() {
  if (reviewPreviewObserver) {
    reviewPreviewObserver.disconnect();
    reviewPreviewObserver = null;
  }
}

function clearReviewPreviewQueue() {
  reviewPreviewQueue.length = 0;
}

function observeReviewPreviews() {
  if (reviewPreviewObserver) reviewPreviewObserver.disconnect();
  reviewPreviewObserver = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        reviewPreviewObserver.unobserve(entry.target);
        scheduleReviewPreview(entry.target);
      }
    },
    { root: reviewList, rootMargin: '160px' }
  );

  document.querySelectorAll('.review-preview[data-needs-preview="true"]').forEach((preview) => {
    reviewPreviewObserver.observe(preview);
  });
}

function scheduleReviewPreview(preview) {
  if (!preview.isConnected || preview.dataset.loadingPreview === 'true') return;
  const finding = findReviewFinding(preview.dataset.findingId);
  if (!finding) return;

  preview.dataset.loadingPreview = 'true';
  delete preview.dataset.needsPreview;
  setReviewPreviewLoading(preview, 'Rendering...');

  reviewPreviewQueue.push({
    preview,
    findingId: finding.id,
    docState: state,
    renderId: state.renderId,
    session: state.reviewPreviewSession,
  });
  drainReviewPreviewQueue();
}

function setReviewPreviewLoading(preview, text) {
  preview.querySelectorAll('.review-preview-frame').forEach((frame) => {
    frame.replaceChildren();
    frame.classList.add('loading');
    frame.classList.remove('empty');
    frame.textContent = text;
  });
}

function drainReviewPreviewQueue() {
  while (
    reviewPreviewActive < REVIEW_PREVIEW_CONCURRENCY &&
    reviewPreviewQueue.length > 0
  ) {
    const job = reviewPreviewQueue.shift();
    reviewPreviewActive++;
    runReviewPreviewJob(job).finally(() => {
      reviewPreviewActive--;
      drainReviewPreviewQueue();
    });
  }
}

async function runReviewPreviewJob(job) {
  const { preview, docState, renderId, session, findingId } = job;
  if (
    !preview.isConnected ||
    state !== docState ||
    renderId !== docState.renderId ||
    session !== docState.reviewPreviewSession
  ) {
    return;
  }

  const cached = docState.reviewPreviewCache.get(findingId);
  if (cached) {
    applyReviewPreviewImages(preview, cached);
    return;
  }

  const finding = findReviewFinding(findingId);
  if (!finding) return;

  const result = await withTimeout(
    renderReviewPreviewImages(finding, docState, renderId, preview),
    REVIEW_PREVIEW_TIMEOUT_MS,
    { redactedError: 'Preview timed out', revealedError: 'Preview timed out' }
  );
  if (!result) return;
  if (
    state !== docState ||
    renderId !== docState.renderId ||
    session !== docState.reviewPreviewSession
  ) {
    return;
  }
  docState.reviewPreviewCache.set(findingId, result);
  if (preview.isConnected) applyReviewPreviewImages(preview, result);
}

function withTimeout(promise, ms, fallback) {
  return new Promise((resolve) => {
    const timer = window.setTimeout(() => resolve(fallback), ms);
    promise
      .then((value) => resolve(value))
      .catch(() => resolve(fallback))
      .finally(() => window.clearTimeout(timer));
  });
}

async function renderReviewPreviewImages(finding, docState, renderId, preview) {
  if (!docState.pdfjsDoc) return null;

  const result = {
    redactedUrl: null,
    revealedUrl: null,
  };
  result.redactedUrl = await renderFindingPreviewFromDoc(docState.pdfjsDoc, finding);
  if (preview?.isConnected) {
    setReviewPreviewFrame(
      preview.querySelector('[data-side="redacted"] .review-preview-frame'),
      result.redactedUrl,
      'As-is preview unavailable',
      'As-is PDF crop'
    );
  }
  if (docState !== state || renderId !== docState.renderId) return null;

  if (finding.repairable && docState.bytes) {
    const repaired = await unredactFinding(docState.bytes, finding);
    if (docState !== state || renderId !== docState.renderId) return null;
    const repairedDoc = await pdfjsLib.getDocument({
      data: toUint8Array(repaired.bytes).slice(),
      ...PDFJS_DOCUMENT_OPTIONS,
    }).promise;
    try {
      result.revealedUrl = await renderFindingPreviewFromDoc(repairedDoc, finding);
    } finally {
      repairedDoc.destroy().catch(() => {});
    }
  }

  return result;
}

async function renderFindingPreviewFromDoc(pdfjsDoc, finding) {
  const page = await pdfjsDoc.getPage(finding.pageIndex + 1);
  const pageViewport = page.getViewport({ scale: 1 });
  const crop = findingPreviewCrop(finding, pageViewport);
  const scale = Math.min(
    REVIEW_PREVIEW_MAX_SCALE,
    Math.max(1.5, REVIEW_PREVIEW_TARGET_WIDTH / Math.max(1, crop.width))
  );
  const viewport = page.getViewport({ scale });
  const cropBox = viewportRect(viewport, crop);
  const output = document.createElement('canvas');
  output.width = Math.max(1, Math.ceil(cropBox.width));
  output.height = Math.max(1, Math.ceil(cropBox.height));
  const outCtx = output.getContext('2d');
  await page.render({
    canvasContext: outCtx,
    viewport,
    transform: [1, 0, 0, 1, -cropBox.left, -cropBox.top],
    annotationMode: pdfjsLib.AnnotationMode.ENABLE_STORAGE,
  }).promise;

  return output.toDataURL('image/png');
}

function findingPreviewCrop(finding, pageViewport) {
  const basis = unionBBoxes([finding.rect, finding.textRect].filter(Boolean)) || finding.rect;
  const padX = Math.max(18, basis.width * 1.4);
  const padY = Math.max(14, basis.height * 3);
  const page = {
    x: pageViewport.viewBox?.[0] || 0,
    y: pageViewport.viewBox?.[1] || 0,
    width: pageViewport.viewBox
      ? pageViewport.viewBox[2] - pageViewport.viewBox[0]
      : pageViewport.width,
    height: pageViewport.viewBox
      ? pageViewport.viewBox[3] - pageViewport.viewBox[1]
      : pageViewport.height,
  };
  const x = Math.max(page.x, basis.x - padX);
  const y = Math.max(page.y, basis.y - padY);
  const right = Math.min(page.x + page.width, basis.x + basis.width + padX);
  const top = Math.min(page.y + page.height, basis.y + basis.height + padY);
  return {
    x,
    y,
    width: Math.max(1, right - x),
    height: Math.max(1, top - y),
  };
}

function unionBBoxes(boxes) {
  if (!boxes.length) return null;
  const minX = Math.min(...boxes.map((box) => box.x));
  const minY = Math.min(...boxes.map((box) => box.y));
  const maxX = Math.max(...boxes.map((box) => box.x + box.width));
  const maxY = Math.max(...boxes.map((box) => box.y + box.height));
  return {
    x: minX,
    y: minY,
    width: maxX - minX,
    height: maxY - minY,
  };
}

function reviewButton(label, action, extraClass = '') {
  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = label;
  button.dataset.action = action;
  if (extraClass) button.classList.add(extraClass);
  return button;
}

function renderReviewHighlights(paper, pageIndex, viewport) {
  paper.querySelector('.reviewLayer')?.remove();
  const findings = state.reviewFindings.filter(
    (finding) =>
      finding.pageIndex === pageIndex && !state.ignoredReviewFindingIds.has(finding.id)
  );
  const comments = flatReviewComments().filter(
    (comment) => comment.pageIndex === pageIndex && comment.rect
  );
  if (findings.length === 0 && comments.length === 0) return;

  const layer = document.createElement('div');
  layer.className = 'reviewLayer';
  layer.style.width = `${Math.ceil(viewport.width)}px`;
  layer.style.height = `${Math.ceil(viewport.height)}px`;
  for (const finding of findings) {
    appendReviewHighlight(layer, viewport, finding.rect, {
      className: 'redaction-highlight',
      focused: state.focusedReviewFindingId === finding.id,
    });
  }
  for (const comment of comments) {
    appendReviewHighlight(layer, viewport, comment.rect, {
      className: 'comment-highlight',
      focused: state.focusedReviewCommentId === comment.id,
      color: comment.color,
    });
  }
  paper.appendChild(layer);
}

function appendReviewHighlight(layer, viewport, sourceRect, { className, focused, color } = {}) {
  const rect = viewportRect(viewport, sourceRect);
  const width = Math.max(12, rect.width);
  const height = Math.max(12, rect.height);
  const el = document.createElement('div');
  el.className = `review-highlight ${className || ''}`.trim();
  el.classList.toggle('focused', Boolean(focused));
  if (color) el.style.setProperty('--comment-color', color);
  el.style.left = `${rect.left}px`;
  el.style.top = `${rect.top}px`;
  el.style.width = `${width}px`;
  el.style.height = `${height}px`;
  layer.appendChild(el);
}

function renderAllReviewHighlights() {
  pageStack.querySelectorAll('.page-view.rendered').forEach((pageView) => {
    const pageIndex = Number(pageView.dataset.index);
    const paper = pageView.querySelector('.page-paper');
    const canvas = paper?.querySelector('canvas');
    if (!paper || !canvas || !state.pdfjsDoc) return;
    const scale =
      Number(paper.querySelector('.annotationLayer')?.style.getPropertyValue('--scale-factor')) ||
      (canvas.clientWidth / canvas.width) * (window.devicePixelRatio || 1);
    state.pdfjsDoc.getPage(pageIndex + 1).then((page) => {
      if (!paper.isConnected || state.numPages <= pageIndex) return;
      renderReviewHighlights(paper, pageIndex, page.getViewport({ scale }));
    });
  });
}

function viewportRect(viewport, rect) {
  const [x1, y1, x2, y2] = viewport.convertToViewportRectangle([
    rect.x,
    rect.y,
    rect.x + rect.width,
    rect.y + rect.height,
  ]);
  const left = Math.min(x1, x2);
  const top = Math.min(y1, y2);
  return {
    left,
    top,
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1),
  };
}

function findReviewFinding(id) {
  return state.reviewFindings.find((finding) => finding.id === id) || null;
}

function findReviewComment(id) {
  return flatReviewComments().find((comment) => comment.id === id) || null;
}

function findReviewCommentByAnnotationId(annotationId) {
  const normalized = normalizeAnnotationId(annotationId);
  return flatReviewComments().find((comment) => comment.annotationId === normalized) || null;
}

function flatReviewComments() {
  const comments = [];
  for (const comment of state.reviewComments || []) {
    comments.push(comment);
    comments.push(...(comment.replies || []));
  }
  return comments;
}

async function focusReviewFinding(finding) {
  const docState = state;
  const renderId = docState.renderId;
  state.focusedReviewFindingId = finding.id;
  state.focusedReviewCommentId = null;
  state.currentPage = finding.pageIndex;
  state.selection = new Set([finding.pageIndex]);
  state.anchor = finding.pageIndex;
  applySelectionStyles();
  scrollMainPageIntoView(finding.pageIndex, 'auto', { stabilize: true });
  updatePageIndicator();
  renderReviewPane();
  await ensureMainPageRendered(finding.pageIndex);
  if (
    docState !== state ||
    renderId !== state.renderId ||
    state.focusedReviewFindingId !== finding.id
  ) {
    return;
  }
  scrollMainPageIntoView(finding.pageIndex, 'auto', { stabilize: true });
  renderAllReviewHighlights();
}

async function focusReviewComment(comment) {
  const docState = state;
  const renderId = docState.renderId;
  state.focusedReviewCommentId = comment.id;
  state.focusedReviewFindingId = null;
  state.currentPage = comment.pageIndex;
  state.selection = new Set([comment.pageIndex]);
  state.anchor = comment.pageIndex;
  applySelectionStyles();
  scrollMainPageIntoView(comment.pageIndex, 'auto', { stabilize: true });
  updatePageIndicator();
  renderReviewPane();
  await ensureMainPageRendered(comment.pageIndex);
  if (
    docState !== state ||
    renderId !== state.renderId ||
    state.focusedReviewCommentId !== comment.id
  ) {
    return;
  }
  scrollMainPageIntoView(comment.pageIndex, 'auto', { stabilize: true });
  renderAllReviewHighlights();
}

async function mutatePdfComment(mutator, { focusAnnotationId = null } = {}) {
  const docState = state;
  const sourceBytes = await getCurrentDocumentBytes();
  const result = await mutator(sourceBytes);
  if (docState !== state || !result?.bytes) return null;

  docState.bytes = result.bytes;
  docState.dirty = true;
  await loadBytes(result.bytes);
  docState.dirty = true;
  updateChrome();
  renderTabs();

  const targetId = focusAnnotationId || result.annotationId;
  if (targetId) {
    await waitForReviewReady();
    const comment = findReviewCommentByAnnotationId(targetId);
    if (comment) {
      state.focusedReviewCommentId = comment.id;
      renderReviewPane();
      renderAllReviewHighlights();
    }
  }
  return result;
}

async function waitForReviewReady() {
  for (let i = 0; i < 50 && state.reviewStatus === 'scanning'; i++) {
    await sleep(50);
  }
}

async function replyToReviewComment(comment) {
  const draft = await openCommentDialog({
    title: 'Reply to comment',
    actionLabel: 'Reply',
  });
  if (!draft) return;
  try {
    await mutatePdfComment((bytes) =>
      replyToPdfComment(bytes, {
        annotationId: comment.annotationId,
        contents: draft.text,
        author: draft.author,
      })
    );
  } catch {
    window.alert('Could not add this reply.');
  }
}

async function toggleReviewCommentResolved(comment) {
  try {
    await mutatePdfComment(
      (bytes) =>
        setPdfCommentResolved(bytes, {
          annotationId: comment.annotationId,
          resolved: !comment.resolved,
        }),
      { focusAnnotationId: comment.annotationId }
    );
  } catch {
    window.alert('Could not update this comment.');
  }
}

async function removeReviewComment(comment, { includeReplies = false } = {}) {
  const message = includeReplies
    ? 'Remove this comment and all replies?'
    : 'Remove this reply?';
  if (!window.confirm(message)) return;
  try {
    await mutatePdfComment((bytes) =>
      removePdfComment(bytes, {
        annotationId: comment.annotationId,
        includeReplies,
      })
    );
  } catch {
    window.alert('Could not remove this comment.');
  }
}

async function repairReviewFinding(finding) {
  if (!finding?.repairable) return;
  state.bytes = await getCurrentDocumentBytes();
  const result = await unredactFinding(state.bytes, finding);
  state.bytes = result.bytes;
  state.dirty = true;
  await loadBytes(result.bytes);
  state.dirty = true;
  updateChrome();
  renderTabs();
}

// =========================================================================
// Editing tools: permanent redaction
// =========================================================================

function setActiveTool(tool) {
  if (!state) return;
  cancelRedactionDraft();
  state.activeTool = tool;
  renderToolState();
}

function renderToolState() {
  if (!state) return;
  const activeTool = state.activeTool || 'select';
  document.body.classList.toggle('redaction-mode', activeTool === 'redact');
  document.body.classList.toggle('comment-mode', activeTool === 'comment');
  reviewAddComment.classList.toggle('active', activeTool === 'comment');
  toolbarComment.classList.toggle('active', activeTool === 'comment');
  toolbarRedact.classList.toggle('active', activeTool === 'redact');
  toolbarComment.disabled = state.numPages === 0;
  toolbarRedact.disabled = state.numPages === 0;
  toolbarTableExport.disabled = state.numPages === 0 || !state.pdfjsDoc;
  toolbarSave.disabled = !state;

  const count = state.pendingRedactions.length;
  redactionTools?.classList.toggle('hidden', count === 0);
  if (redactionApply) {
    redactionApply.disabled = !state.bytes || count === 0;
    redactionApply.textContent = count > 0 ? 'Apply redactions' : 'Apply redactions';
  }
  if (redactionCount) {
    redactionCount.textContent =
      count > 0 ? `${count} pending ${count === 1 ? 'box' : 'boxes'}` : '';
  }
  renderAllRedactionLayers();
  renderAllCommentLayers();
  updateChrome();
}

function renderRedactionLayer(paper, pageIndex, viewport) {
  paper.querySelector('.redactionLayer')?.remove();

  const layer = document.createElement('div');
  layer.className = 'redactionLayer';
  layer.classList.toggle('active', state.activeTool === 'redact');
  layer.dataset.pageIndex = String(pageIndex);
  layer.style.width = `${Math.ceil(viewport.width)}px`;
  layer.style.height = `${Math.ceil(viewport.height)}px`;
  layer.__pdfViewport = viewport;
  layer.addEventListener('pointerdown', onRedactionPointerDown);
  layer.addEventListener('click', swallowRedactionLayerClick);

  for (const redaction of state.pendingRedactions.filter((item) => item.pageIndex === pageIndex)) {
    const box = document.createElement('div');
    box.className = 'redaction-box queued';
    drawRedactionBox(box, viewport, redaction.rect);
    layer.appendChild(box);
  }

  paper.appendChild(layer);
}

function renderAllRedactionLayers() {
  pageStack.querySelectorAll('.page-view.rendered').forEach((pageView) => {
    const paper = pageView.querySelector('.page-paper');
    const viewport = paper?.__pdfViewport;
    if (!paper || !viewport) return;
    renderRedactionLayer(paper, Number(pageView.dataset.index), viewport);
  });
}

function renderCommentLayer(paper, pageIndex, viewport) {
  paper.querySelector('.commentLayer')?.remove();

  const layer = document.createElement('div');
  layer.className = 'commentLayer';
  layer.classList.toggle('active', state.activeTool === 'comment');
  layer.dataset.pageIndex = String(pageIndex);
  layer.style.width = `${Math.ceil(viewport.width)}px`;
  layer.style.height = `${Math.ceil(viewport.height)}px`;
  layer.__pdfViewport = viewport;
  layer.addEventListener('pointerdown', onCommentPointerDown);
  paper.appendChild(layer);
}

function renderAllCommentLayers() {
  pageStack.querySelectorAll('.page-view.rendered').forEach((pageView) => {
    const paper = pageView.querySelector('.page-paper');
    const viewport = paper?.__pdfViewport;
    if (!paper || !viewport) return;
    renderCommentLayer(paper, Number(pageView.dataset.index), viewport);
  });
}

async function onCommentPointerDown(e) {
  if (
    !state ||
    state.activeTool !== 'comment' ||
    state.commentDialogOpen ||
    commentPlacementInProgress ||
    e.button !== 0
  ) {
    return;
  }
  e.preventDefault();
  e.stopPropagation();

  const layer = e.currentTarget;
  const pageIndex = Number(layer.dataset.pageIndex);
  const viewport = layer.__pdfViewport;
  const point = pdfPointFromEvent(layer, viewport, e);
  const rect = commentRectFromPoint(point);

  await addCommentAt(pageIndex, rect, {
    title: `Add comment on page ${pageIndex + 1}`,
    selectAfter: true,
  });
}

async function addCommentAt(pageIndex, rect, { title, selectAfter = false } = {}) {
  commentPlacementInProgress = true;
  try {
    const draft = await openCommentDialog({
      title: title || `Add comment on page ${pageIndex + 1}`,
      actionLabel: 'Add',
    });
    if (!draft) return;
    await mutatePdfComment((bytes) =>
      addPdfComment(bytes, {
        pageIndex,
        rect,
        contents: draft.text,
        author: draft.author,
      })
    );
    if (selectAfter) setActiveTool('select');
  } catch {
    window.alert('Could not add this comment.');
  } finally {
    commentPlacementInProgress = false;
  }
}

function commentRectFromPoint(point) {
  const size = 22;
  return {
    x: Math.max(0, point.x - size / 2),
    y: Math.max(0, point.y - size / 2),
    width: size,
    height: size,
  };
}

function startCommentTool() {
  if (!state?.numPages) return;
  setActiveTool('comment');
  renderReviewPane();
}

function toggleCommentTool() {
  if (!state?.numPages) return;
  if (state.activeTool === 'comment') setActiveTool('select');
  else startCommentTool();
}

function swallowRedactionLayerClick(e) {
  if (state?.activeTool !== 'redact') return;
  e.preventDefault();
  e.stopPropagation();
}

function onRedactionPointerDown(e) {
  if (!state || state.activeTool !== 'redact' || e.button !== 0) return;
  e.preventDefault();
  e.stopPropagation();

  const layer = e.currentTarget;
  const pageIndex = Number(layer.dataset.pageIndex);
  const viewport = layer.__pdfViewport;
  const start = pdfPointFromEvent(layer, viewport, e);
  const box = document.createElement('div');
  box.className = 'redaction-box draft';
  layer.appendChild(box);

  redactionDraft = { layer, viewport, pageIndex, start, current: start, box };
  layer.setPointerCapture(e.pointerId);
  layer.addEventListener('pointermove', onRedactionPointerMove);
  layer.addEventListener('pointerup', onRedactionPointerUp, { once: true });
  layer.addEventListener('pointercancel', onRedactionPointerCancel, { once: true });
}

function onRedactionPointerMove(e) {
  if (!redactionDraft) return;
  e.preventDefault();
  redactionDraft.current = pdfPointFromEvent(
    redactionDraft.layer,
    redactionDraft.viewport,
    e
  );
  drawRedactionBox(
    redactionDraft.box,
    redactionDraft.viewport,
    pdfRectFromPoints(redactionDraft.start, redactionDraft.current)
  );
}

function onRedactionPointerUp(e) {
  if (!redactionDraft) return;
  e.preventDefault();
  const draft = redactionDraft;
  const rect = pdfRectFromPoints(draft.start, draft.current);
  finishRedactionDraft(e.pointerId);

  if (rect.width < 2 || rect.height < 2) return;
  state.pendingRedactions.push({
    id: makeRedactionId(draft.pageIndex, rect),
    pageIndex: draft.pageIndex,
    rect,
  });
  renderToolState();
}

function onRedactionPointerCancel(e) {
  finishRedactionDraft(e.pointerId);
}

function finishRedactionDraft(pointerId) {
  if (!redactionDraft) return;
  const { layer, box } = redactionDraft;
  layer.removeEventListener('pointermove', onRedactionPointerMove);
  try {
    layer.releasePointerCapture(pointerId);
  } catch {
    /* ignore */
  }
  box.remove();
  redactionDraft = null;
}

function cancelRedactionDraft() {
  if (!redactionDraft) return;
  redactionDraft.box.remove();
  redactionDraft.layer.removeEventListener('pointermove', onRedactionPointerMove);
  redactionDraft = null;
}

function pdfPointFromEvent(layer, viewport, e) {
  return pdfPointFromClient(layer, viewport, e.clientX, e.clientY);
}

function pdfPointFromClient(layer, viewport, clientX, clientY) {
  const bounds = layer.getBoundingClientRect();
  const x = Math.max(0, Math.min(bounds.width, clientX - bounds.left));
  const y = Math.max(0, Math.min(bounds.height, clientY - bounds.top));
  const [pdfX, pdfY] = viewport.convertToPdfPoint(x, y);
  return { x: pdfX, y: pdfY };
}

function pdfRectFromPoints(a, b) {
  const x = Math.min(a.x, b.x);
  const y = Math.min(a.y, b.y);
  return {
    x,
    y,
    width: Math.abs(a.x - b.x),
    height: Math.abs(a.y - b.y),
  };
}

function drawRedactionBox(el, viewport, rect) {
  const box = viewportRect(viewport, rect);
  el.style.left = `${box.left}px`;
  el.style.top = `${box.top}px`;
  el.style.width = `${box.width}px`;
  el.style.height = `${box.height}px`;
}

function makeRedactionId(pageIndex, rect) {
  return `redact:${pageIndex}:${Date.now()}:${Math.round(rect.x)}:${Math.round(rect.y)}`;
}

async function applyQueuedRedactions() {
  if (!state.pendingRedactions.length) return;
  if (redactionApply) {
    redactionApply.disabled = true;
    redactionApply.textContent = 'Applying...';
  }
  try {
    await applyPermanentRedactions(state.pendingRedactions, { clearQueued: true });
  } finally {
    if (redactionApply) redactionApply.textContent = 'Apply redactions';
    renderToolState();
  }
}

async function applyPermanentRedactions(redactions, { clearQueued = false } = {}) {
  const items = (redactions || []).filter((item) => item?.rect && item.pageIndex >= 0);
  if (!items.length || !state.pdfjsDoc) return;

  const docState = state;
  const sourceBytes = await getCurrentDocumentBytes();
  const pageImages = await renderRedactedPageImages(docState, items);
  if (docState !== state) return;

  const result = await permanentlyRedactPages(sourceBytes, pageImages, docState.outline);
  docState.activeTool = 'select';
  if (clearQueued) docState.pendingRedactions = [];
  docState.reviewFindings = [];
  docState.reviewComments = [];
  docState.focusedReviewFindingId = null;
  docState.focusedReviewCommentId = null;
  docState.reviewStatus = 'idle';
  docState.dirty = true;
  await loadBytes(result.bytes);
  docState.dirty = true;
  updateChrome();
  renderTabs();
}

async function renderRedactedPageImages(docState, redactions) {
  const grouped = new Map();
  for (const redaction of redactions) {
    if (!grouped.has(redaction.pageIndex)) grouped.set(redaction.pageIndex, []);
    grouped.get(redaction.pageIndex).push(redaction.rect);
  }

  const pageImages = [];
  for (const [pageIndex, rects] of grouped) {
    const page = await docState.pdfjsDoc.getPage(pageIndex + 1);
    if (docState !== state) return [];

    const baseViewport = page.getViewport({ scale: 1 });
    const renderViewport = page.getViewport({ scale: REDACTION_RENDER_SCALE });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(renderViewport.width);
    canvas.height = Math.ceil(renderViewport.height);

    const ctx = canvas.getContext('2d');
    await page.render({
      canvasContext: ctx,
      viewport: renderViewport,
      annotationMode: pdfjsLib.AnnotationMode.ENABLE_STORAGE,
    }).promise;
    if (docState !== state) return [];

    ctx.fillStyle = '#000';
    for (const rect of rects) {
      const box = viewportRect(renderViewport, rect);
      ctx.fillRect(box.left, box.top, box.width, box.height);
    }

    pageImages.push({
      pageIndex,
      width: baseViewport.width,
      height: baseViewport.height,
      pngBytes: dataUrlToBytes(canvas.toDataURL('image/png')),
    });
  }

  return pageImages;
}

function dataUrlToBytes(dataUrl) {
  const [, base64] = dataUrl.split(',');
  return Uint8Array.from(atob(base64), (ch) => ch.charCodeAt(0));
}

// =========================================================================
// Workspace tabs
// =========================================================================

function addDocument(docState) {
  workspace.documents.push(docState);
  activateDocument(docState.id);
}

function activateDocument(id) {
  const next = workspace.documents.find((doc) => doc.id === id);
  if (!next || next === state) return;
  if (state) {
    state.renderId++;
    state.sidebarDropIndex = null;
    state.previewDropIndex = null;
  }
  state = next;
  state.sidebarDropIndex = null;
  state.previewDropIndex = null;
  workspace.activeId = id;
  hideContextMenu();
  hideDropIndicator();
  hideViewerDropIndicator();
  stopSidebarAutoScroll();
  renderTabs();
  renderSidebar();
  renderMainView();
  renderReviewPane();
  renderToolState();
  updateChrome();
}

async function closeDocument(id) {
  const docState = workspace.documents.find((doc) => doc.id === id);
  if (!docState) return;
  if (!confirmDiscardDocument(docState)) return;

  const wasActive = docState === state;
  workspace.documents = workspace.documents.filter((doc) => doc.id !== id);
  if (docState.pdfjsDoc) {
    try {
      await docState.pdfjsDoc.destroy();
    } catch {
      /* ignore */
    }
  }

  if (workspace.documents.length === 0) {
    state = null;
    const doc = createDocumentState();
    workspace.documents.push(doc);
    state = doc;
    workspace.activeId = doc.id;
    await loadBytes(null, { freshOpen: true });
    return;
  }

  if (wasActive) {
    const fallback = workspace.documents[Math.max(0, workspace.documents.length - 1)];
    state = null;
    activateDocument(fallback.id);
  } else {
    renderTabs();
  }
}

function renderTabs() {
  tabbar.replaceChildren();
  for (const doc of workspace.documents) {
    const tab = document.createElement('button');
    tab.type = 'button';
    tab.className = 'doc-tab';
    tab.classList.toggle('active', doc.id === workspace.activeId);
    tab.title = doc.filePath || doc.fileName;
    tab.dataset.id = String(doc.id);

    const name = document.createElement('span');
    name.className = 'doc-tab-name';
    name.textContent = doc.fileName;
    tab.appendChild(name);

    if (doc.dirty) {
      const dirty = document.createElement('span');
      dirty.className = 'doc-tab-dirty';
      dirty.textContent = '!';
      dirty.title = 'Unsaved changes';
      dirty.setAttribute('aria-label', 'Warning: unsaved changes');
      tab.appendChild(dirty);
    }

    const close = document.createElement('span');
    close.className = 'doc-tab-close';
    close.textContent = '×';
    close.title = `Close ${doc.fileName}`;
    close.setAttribute('role', 'button');
    close.setAttribute('aria-label', `Close ${doc.fileName}`);
    tab.appendChild(close);

    tabbar.appendChild(tab);
  }
}

tabbar.addEventListener('click', async (e) => {
  const tab = e.target.closest('.doc-tab');
  if (!tab) return;
  const id = Number(tab.dataset.id);
  if (e.target.closest('.doc-tab-close')) {
    await closeDocument(id);
  } else {
    activateDocument(id);
  }
});

function confirmDiscardDocument(docState) {
  if (!docState?.dirty) return true;
  return window.confirm(`Discard unsaved changes to "${docState.fileName}"?`);
}

function confirmCloseAllDocuments() {
  const dirty = workspace.documents.filter((doc) => doc.dirty);
  if (dirty.length === 0) return true;
  const first = dirty[0]?.fileName || 'this document';
  const detail =
    dirty.length === 1
      ? `"${first}" has unsaved changes.`
      : `${dirty.length} documents have unsaved changes.`;
  return window.confirm(`${detail}\n\nDiscard changes and close PDF Workbench?`);
}

// =========================================================================
// Selection
// =========================================================================

function onThumbClick(index, e) {
  selectPageFromEvent(index, e);
}

function selectPageFromEvent(index, e, { scrollMain = true } = {}) {
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
  if (scrollMain) scrollMainPageIntoView(index);
  updatePageIndicator();
  updateChrome();
}

function applySelectionStyles({ scrollThumb = true } = {}) {
  thumbList.querySelectorAll('.thumb').forEach((el) => {
    const i = Number(el.dataset.index);
    el.classList.toggle('selected', state.selection.has(i));
    el.classList.toggle('current', i === state.currentPage);
  });
  pageStack.querySelectorAll('.page-view').forEach((el) => {
    const i = Number(el.dataset.index);
    el.classList.toggle('selected', state.selection.has(i));
    el.classList.toggle('current', i === state.currentPage);
  });
  const cur = thumbList.querySelector('.thumb.current');
  if (cur && scrollThumb) cur.scrollIntoView({ block: 'nearest' });
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
  contextCommentTarget = null;
  if (!state.selection.has(index)) {
    state.selection.clear();
    state.selection.add(index);
    state.anchor = index;
    state.currentPage = index;
    applySelectionStyles();
    scrollMainPageIntoView(index);
    updatePageIndicator();
    updateChrome();
  }
  showContextMenu(e.clientX, e.clientY);
}

function showContextMenu(x, y, { commentTarget = null } = {}) {
  if (state.numPages === 0) return;
  contextCommentTarget = commentTarget;
  const commentItem = contextMenu.querySelector('[data-action="add-comment"]');
  if (commentItem) {
    commentItem.classList.toggle('disabled', !commentTarget);
    commentItem.setAttribute('aria-disabled', commentTarget ? 'false' : 'true');
  }
  contextMenu.classList.remove('hidden');
  const rect = contextMenu.getBoundingClientRect();
  const px = Math.min(x, window.innerWidth - rect.width - 4);
  const py = Math.min(y, window.innerHeight - rect.height - 4);
  contextMenu.style.left = `${px}px`;
  contextMenu.style.top = `${py}px`;
}

function hideContextMenu() {
  contextMenu.classList.add('hidden');
  contextCommentTarget = null;
}

contextMenu.addEventListener('click', async (e) => {
  const item = e.target.closest('.ctx-item');
  if (!item || item.classList.contains('disabled')) return;
  const action = item.dataset.action;
  const commentTarget = contextCommentTarget;
  hideContextMenu();
  if (action === 'add-comment' && commentTarget) {
    await addCommentAt(commentTarget.pageIndex, commentTarget.rect, {
      title: `Add comment on page ${commentTarget.pageIndex + 1}`,
    });
  } else if (action === 'delete') await deleteSelectedPages();
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
  const pageView = e.target.closest?.('.page-view');
  let commentTarget = null;
  if (pageView) {
    state.currentPage = Number(pageView.dataset.index);
    commentTarget = commentTargetFromPageEvent(pageView, e);
  }
  if (!state.selection.has(state.currentPage)) {
    state.selection.clear();
    state.selection.add(state.currentPage);
    state.anchor = state.currentPage;
    applySelectionStyles();
    updateChrome();
  }
  showContextMenu(e.clientX, e.clientY, { commentTarget });
});

function commentTargetFromPageEvent(pageView, e) {
  const paper = pageView.querySelector('.page-paper');
  const viewport = paper?.__pdfViewport;
  if (!paper || !viewport) return null;
  const point = pdfPointFromClient(paper, viewport, e.clientX, e.clientY);
  return {
    pageIndex: Number(pageView.dataset.index),
    rect: commentRectFromPoint(point),
  };
}

// =========================================================================
// Mutations
// =========================================================================

async function deleteSelectedPages() {
  if (state.selection.size === 0 || state.numPages === 0) return;
  state.bytes = await getCurrentDocumentBytes();
  const indices = [...state.selection];
  const minIndex = Math.min(...indices);
  const { bytes, outline, remaining } = await deletePages(
    state.bytes,
    indices,
    state.outline
  );
  state.outline = outline;
  state.dirty = true;
  renderTabs();
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
    scrollMainPageIntoView(state.currentPage, 'auto');
    updatePageIndicator();
    updateChrome();
  }
}

async function doInsert({ srcBytes, pageIndices, insertAt, name }) {
  state.bytes = state.bytes ? await getCurrentDocumentBytes() : null;
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
  renderTabs();
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
    scrollMainPageIntoView(at, 'auto');
    updatePageIndicator();
    updateChrome();
  }
}

// =========================================================================
// File operations
// =========================================================================

async function newDocument() {
  const doc = createDocumentState();
  addDocument(doc);
  await loadBytes(null, { freshOpen: true });
}

async function openDocument() {
  const result = await window.api.openFile();
  if (!result) return;
  const doc = createDocumentState({ filePath: result.path, fileName: result.name });
  addDocument(doc);
  state.dirty = false;
  await loadBytes(result.data, {
    path: result.path,
    name: result.name,
    freshOpen: true,
  });
  updateChrome();
  renderTabs();
}

async function saveDocument(forceDialog = false) {
  let targetPath = state.filePath;
  if (forceDialog || !targetPath) {
    targetPath = await window.api.saveFileDialog(state.fileName);
    if (!targetPath) return;
  }
  // An empty document is serialized to a valid zero-page PDF only on save.
  const outBytes = await getCurrentDocumentBytes();
  const res = await window.api.writeFile(targetPath, outBytes);
  state.bytes = outBytes;
  state.filePath = targetPath;
  if (res && res.name) state.fileName = res.name;
  state.dirty = false;
  updateChrome();
  renderTabs();
}

async function exportDetectedTableToExcel() {
  if (!state?.pdfjsDoc || state.numPages === 0) return;

  const docState = state;
  const originalLabel = toolbarTableExport.textContent;
  toolbarTableExport.disabled = true;
  toolbarTableExport.textContent = 'Finding...';
  statusEl.textContent = 'Finding likely table rows';

  let extraction;
  try {
    extraction = await extractTableFromPdf(docState.pdfjsDoc);
  } catch {
    window.alert('Could not inspect this PDF for table text.');
    updateChrome();
    renderToolState();
    return;
  } finally {
    toolbarTableExport.textContent = originalLabel;
    renderToolState();
  }

  if (docState !== state) return;
  if (!extraction.table.length) {
    window.alert('No likely table was found in this PDF.');
    updateChrome();
    return;
  }

  const confirmed = await openTableDialog(extraction);
  if (!confirmed) {
    updateChrome();
    return;
  }

  const rows = trimExportRows(confirmed.rows);
  if (!rows.length || rows.every((row) => row.every((cell) => !cell.trim()))) {
    window.alert('There is no table data to export.');
    return;
  }

  let bytes;
  try {
    bytes = createXlsxWorkbook(rows, {
      sheetName: fileBaseName(state.fileName) || 'Extracted Table',
      firstRowHeader: confirmed.firstRowHeader,
    });
  } catch {
    window.alert('Could not create the Excel workbook.');
    return;
  }

  const targetPath = await window.api.saveXlsxDialog(defaultExcelName());
  if (!targetPath) return;

  try {
    const result = await window.api.writeFile(targetPath, bytes);
    statusEl.textContent = `Exported ${rows.length} rows to ${result?.name || 'Excel'}`;
  } catch {
    window.alert('Could not save the Excel workbook.');
  }
}

function defaultExcelName() {
  return `${fileBaseName(state?.fileName) || 'table'} table.xlsx`;
}

function fileBaseName(fileName) {
  return String(fileName || '').replace(/\.[^.]+$/, '').trim();
}

function trimExportRows(rows) {
  const normalized = rows.map((row) => row.map((cell) => String(cell || '').trim()));
  let lastColumn = Math.max(0, ...normalized.map((row) => row.length)) - 1;
  while (lastColumn >= 0 && normalized.every((row) => !row[lastColumn])) {
    lastColumn--;
  }
  return normalized
    .map((row) => row.slice(0, lastColumn + 1))
    .filter((row) => row.some(Boolean));
}

function openTableDialog(extraction) {
  return new Promise((resolve) => {
    tableDialogResolve = resolve;
    tableDraftRows = cloneTableRows(extraction.table);
    selectedTableCell = { row: 0, col: 0 };
    tableFirstRowHeader.checked = tableDraftRows.length > 1;
    tableError.classList.add('hidden');
    tableSummary.textContent = describeTableExtraction(extraction);
    renderTableGrid();
    tableDialogOverlay.classList.remove('hidden');
    requestAnimationFrame(() => focusTableCell(0, 0));
  });
}

function closeTableDialog(result) {
  tableDialogOverlay.classList.add('hidden');
  const resolve = tableDialogResolve;
  tableDialogResolve = null;
  if (resolve) resolve(result);
}

function describeTableExtraction(extraction) {
  const rows = extraction.table.length;
  const cols = Math.max(0, ...extraction.table.map((row) => row.length));
  const span = extraction.pageSpan
    ? `pages ${extraction.pageSpan.start + 1}-${extraction.pageSpan.end + 1}`
    : 'the document';
  const removed = extraction.stats?.duplicateHeadersRemoved || 0;
  const duplicateText = removed
    ? ` Removed ${removed} repeated header ${removed === 1 ? 'row' : 'rows'}.`
    : '';
  return `Auto-selected ${rows} rows and ${cols} columns across ${span}.${duplicateText}`;
}

function cloneTableRows(rows) {
  const width = Math.max(1, ...rows.map((row) => row.length));
  const source = rows.length ? rows : [['']];
  return source.map((row) =>
    Array.from({ length: width }, (_, index) => String(row[index] || ''))
  );
}

function renderTableGrid() {
  const rows = tableDraftRows;
  const width = Math.max(1, ...rows.map((row) => row.length));
  tableGrid.replaceChildren();

  const thead = document.createElement('thead');
  const header = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'row-number';
  header.appendChild(corner);
  for (let col = 0; col < width; col++) {
    const th = document.createElement('th');
    th.textContent = tableColumnName(col);
    th.classList.toggle('selected', selectedTableCell.col === col);
    header.appendChild(th);
  }
  thead.appendChild(header);
  tableGrid.appendChild(thead);

  const tbody = document.createElement('tbody');
  rows.forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    const rowHeader = document.createElement('th');
    rowHeader.className = 'row-number';
    rowHeader.textContent = String(rowIndex + 1);
    rowHeader.classList.toggle('selected', selectedTableCell.row === rowIndex);
    tr.appendChild(rowHeader);

    for (let col = 0; col < width; col++) {
      const td = document.createElement('td');
      td.classList.toggle(
        'selected',
        selectedTableCell.row === rowIndex && selectedTableCell.col === col
      );
      const input = document.createElement('input');
      input.className = 'table-cell-input';
      input.type = 'text';
      input.value = row[col] || '';
      input.dataset.row = String(rowIndex);
      input.dataset.col = String(col);
      input.addEventListener('input', onTableCellInput);
      input.addEventListener('focus', onTableCellFocus);
      input.addEventListener('click', onTableCellFocus);
      td.appendChild(input);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  tableGrid.appendChild(tbody);
  updateTableEditButtons();
}

function onTableCellInput(e) {
  const row = Number(e.target.dataset.row);
  const col = Number(e.target.dataset.col);
  ensureTableSize(row + 1, col + 1);
  tableDraftRows[row][col] = e.target.value;
}

function onTableCellFocus(e) {
  const row = Number(e.target.dataset.row);
  const col = Number(e.target.dataset.col);
  selectedTableCell = { row, col };
  tableGrid.querySelectorAll('.selected').forEach((el) => el.classList.remove('selected'));
  e.target.closest('td')?.classList.add('selected');
  tableGrid.querySelector(`tbody tr:nth-child(${row + 1}) .row-number`)?.classList.add('selected');
  tableGrid.querySelector(`thead th:nth-child(${col + 2})`)?.classList.add('selected');
  updateTableEditButtons();
}

function focusTableCell(row, col) {
  const input = tableGrid.querySelector(
    `.table-cell-input[data-row="${row}"][data-col="${col}"]`
  );
  input?.focus();
  input?.select();
}

function addTableRow() {
  const width = Math.max(1, ...tableDraftRows.map((row) => row.length));
  const insertAt = Math.min(tableDraftRows.length, selectedTableCell.row + 1);
  tableDraftRows.splice(insertAt, 0, Array.from({ length: width }, () => ''));
  selectedTableCell = { row: insertAt, col: Math.min(selectedTableCell.col, width - 1) };
  renderTableGrid();
  focusTableCell(selectedTableCell.row, selectedTableCell.col);
}

function removeTableRow() {
  if (tableDraftRows.length <= 1) {
    tableDraftRows = [Array.from({ length: Math.max(1, tableDraftRows[0]?.length || 1) }, () => '')];
  } else {
    tableDraftRows.splice(selectedTableCell.row, 1);
  }
  selectedTableCell.row = Math.max(0, Math.min(selectedTableCell.row, tableDraftRows.length - 1));
  renderTableGrid();
  focusTableCell(selectedTableCell.row, selectedTableCell.col);
}

function addTableColumn() {
  const insertAt = Math.min(tableDraftRows[0]?.length || 0, selectedTableCell.col + 1);
  tableDraftRows.forEach((row) => row.splice(insertAt, 0, ''));
  selectedTableCell = { row: selectedTableCell.row, col: insertAt };
  renderTableGrid();
  focusTableCell(selectedTableCell.row, selectedTableCell.col);
}

function removeTableColumn() {
  const width = Math.max(1, ...tableDraftRows.map((row) => row.length));
  if (width <= 1) {
    tableDraftRows.forEach((row) => {
      row[0] = '';
    });
  } else {
    tableDraftRows.forEach((row) => row.splice(selectedTableCell.col, 1));
  }
  selectedTableCell.col = Math.max(0, Math.min(selectedTableCell.col, width - 2));
  renderTableGrid();
  focusTableCell(selectedTableCell.row, selectedTableCell.col);
}

function ensureTableSize(height, width) {
  while (tableDraftRows.length < height) tableDraftRows.push([]);
  tableDraftRows.forEach((row) => {
    while (row.length < width) row.push('');
  });
}

function updateTableEditButtons() {
  const width = Math.max(1, ...tableDraftRows.map((row) => row.length));
  tableRemoveRow.disabled = tableDraftRows.length === 0;
  tableRemoveColumn.disabled = width === 0;
}

function confirmTableExport() {
  const rows = cloneTableRows(tableDraftRows);
  if (!rows.some((row) => row.some((cell) => cell.trim()))) {
    tableError.textContent = 'Add at least one cell before exporting.';
    tableError.classList.remove('hidden');
    return;
  }
  closeTableDialog({
    rows,
    firstRowHeader: tableFirstRowHeader.checked,
  });
}

function tableColumnName(index) {
  let n = index + 1;
  let name = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    n = Math.floor((n - mod) / 26);
  }
  return name;
}

// =========================================================================
// Drag & drop of external PDFs
// =========================================================================

const EDGE_HOVER_DELAY = 650;
const SIDEBAR_EDGE_SIZE = 58;
const SIDEBAR_SCROLL_STEP = 12;

let dragDepth = 0;
let sidebarScrollDirection = 0;
let sidebarScrollTimer = null;
let sidebarScrollFrame = null;
let sidebarScrollActive = false;
let sidebarScrollClientY = 0;
let pendingTabId = null;
let pendingTabTimer = null;

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

  updatePendingTabSwitch(e.target);

  if (tabbar.contains(e.target)) {
    clearSidebarDropTarget();
    clearViewerDropTarget();
    stopSidebarAutoScroll();
    return;
  }

  if (sidebar.contains(e.target)) {
    clearViewerDropTarget();
    updateSidebarDropTarget(e.clientY);
    updateSidebarAutoScroll(e.clientY);
  } else if (viewer.contains(e.target)) {
    clearSidebarDropTarget();
    stopSidebarAutoScroll();
    updateViewerDropTarget(e.clientY);
  } else {
    clearSidebarDropTarget();
    clearViewerDropTarget();
    stopSidebarAutoScroll();
  }
});

window.addEventListener('dragleave', (e) => {
  if (!dragHasFiles(e)) return;
  dragDepth--;
  if (dragDepth <= 0) {
    dragDepth = 0;
    dragOverlay.classList.add('hidden');
    clearDragFeedback();
  }
});

window.addEventListener('drop', async (e) => {
  if (!dragHasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  dragOverlay.classList.add('hidden');

  const dropIndex =
    state.sidebarDropIndex ?? state.previewDropIndex ?? null;
  clearDragFeedback();

  const files = [...e.dataTransfer.files].filter(
    (f) => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf')
  );
  if (files.length === 0) return;

  // Process files one at a time, each through the insert dialog.
  for (const file of files) {
    const srcBytes = new Uint8Array(await file.arrayBuffer());
    let srcCount;
    try {
      const probe = await pdfjsLib.getDocument({
        data: srcBytes.slice(),
        ...PDFJS_DOCUMENT_OPTIONS,
      }).promise;
      srcCount = probe.numPages;
      await probe.destroy();
    } catch {
      window.alert(`Could not read "${file.name}" as a PDF.`);
      continue;
    }
    const choice = await openInsertDialog({
      fileName: file.name,
      srcCount,
      srcBytes,
      dropIndex,
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

function clearDragFeedback() {
  clearSidebarDropTarget();
  clearViewerDropTarget();
  stopSidebarAutoScroll();
  clearPendingTabSwitch();
}

function updateSidebarDropTarget(clientY) {
  if (state.numPages === 0) {
    clearSidebarDropTarget();
    return;
  }
  const idx = computeSidebarInsertIndex(clientY);
  state.sidebarDropIndex = idx;
  showDropIndicator(idx);
}

function clearSidebarDropTarget() {
  state.sidebarDropIndex = null;
  hideDropIndicator();
}

function updateViewerDropTarget(clientY) {
  const idx = computeViewerInsertIndex(clientY);
  state.previewDropIndex = idx;
  showViewerDropIndicator(idx);
}

function clearViewerDropTarget() {
  state.previewDropIndex = null;
  hideViewerDropIndicator();
}

function computeViewerInsertIndex(clientY) {
  if (state.numPages === 0) return 0;
  const pages = [...pageStack.querySelectorAll('.page-view')];
  if (pages.length === 0) return state.currentPage + 1;
  for (const page of pages) {
    const rect = page.getBoundingClientRect();
    if (clientY < rect.top + rect.height / 2) return Number(page.dataset.index);
  }
  return state.numPages;
}

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

function showViewerDropIndicator(index) {
  viewerDropIndicator.replaceChildren();
  const label = document.createElement('span');
  label.textContent =
    state.numPages === 0 ? 'Insert into empty document' : describeDropIndex(index);
  viewerDropIndicator.appendChild(label);
  viewerDropIndicator.classList.remove('hidden', 'empty');

  const viewerRect = viewer.getBoundingClientRect();
  if (state.numPages === 0) {
    viewerDropIndicator.classList.add('empty');
    const width = Math.min(360, viewerRect.width - 48);
    viewerDropIndicator.style.left = `${viewerRect.left + (viewerRect.width - width) / 2}px`;
    viewerDropIndicator.style.top = `${viewerRect.top + viewerRect.height / 2 - 26}px`;
    viewerDropIndicator.style.width = `${width}px`;
    return;
  }

  const beforeEl = pageStack.querySelector(`.page-view[data-index="${index - 1}"]`);
  const afterEl = pageStack.querySelector(`.page-view[data-index="${index}"]`);
  const beforeRect = beforeEl?.getBoundingClientRect();
  const afterRect = afterEl?.getBoundingClientRect();
  let lineY;
  if (beforeRect && afterRect) lineY = (beforeRect.bottom + afterRect.top) / 2;
  else if (afterRect) lineY = afterRect.top - 12;
  else if (beforeRect) lineY = beforeRect.bottom + 12;
  else lineY = viewerRect.top + viewerRect.height / 2;

  const targetRect = afterRect || beforeRect || viewerRect;
  const width = Math.max(180, Math.min(targetRect.width, viewerRect.width - 56));
  const left = Math.max(
    viewerRect.left + 28,
    Math.min(targetRect.left + (targetRect.width - width) / 2, viewerRect.right - width - 28)
  );

  viewerDropIndicator.style.left = `${left}px`;
  viewerDropIndicator.style.top = `${lineY - 14}px`;
  viewerDropIndicator.style.width = `${width}px`;
}

function hideViewerDropIndicator() {
  viewerDropIndicator.classList.add('hidden');
  viewerDropIndicator.replaceChildren();
}

function describeDropIndex(index) {
  if (state.numPages === 0) return 'Insert into empty document';
  if (index <= 0) return 'Insert before page 1';
  if (index >= state.numPages) return `Insert after page ${state.numPages}`;
  return `Insert between pages ${index} and ${index + 1}`;
}

function updateSidebarAutoScroll(clientY) {
  sidebarScrollClientY = clientY;
  const direction = getSidebarScrollDirection(clientY);
  if (!direction) {
    stopSidebarAutoScroll();
    return;
  }
  if (direction === sidebarScrollDirection) {
    showSidebarScrollIndicator(direction, sidebarScrollActive);
    return;
  }

  stopSidebarAutoScroll();
  sidebarScrollDirection = direction;
  showSidebarScrollIndicator(direction, false);
  sidebarScrollTimer = window.setTimeout(() => {
    sidebarScrollTimer = null;
    sidebarScrollActive = true;
    showSidebarScrollIndicator(sidebarScrollDirection, true);
    runSidebarAutoScroll();
  }, EDGE_HOVER_DELAY);
}

function getSidebarScrollDirection(clientY) {
  if (state.numPages === 0) return 0;
  const rect = sidebar.getBoundingClientRect();
  const canScrollUp = sidebar.scrollTop > 0;
  const canScrollDown =
    sidebar.scrollTop + sidebar.clientHeight < sidebar.scrollHeight - 1;

  if (clientY < rect.top + SIDEBAR_EDGE_SIZE && canScrollUp) return -1;
  if (clientY > rect.bottom - SIDEBAR_EDGE_SIZE && canScrollDown) return 1;
  return 0;
}

function runSidebarAutoScroll() {
  if (!sidebarScrollDirection) return;
  const before = sidebar.scrollTop;
  sidebar.scrollTop += sidebarScrollDirection * SIDEBAR_SCROLL_STEP;
  updateSidebarDropTarget(sidebarScrollClientY);

  if (sidebar.scrollTop === before && !getSidebarScrollDirection(sidebarScrollClientY)) {
    stopSidebarAutoScroll();
    return;
  }
  sidebarScrollFrame = window.requestAnimationFrame(runSidebarAutoScroll);
}

function stopSidebarAutoScroll() {
  if (sidebarScrollTimer) window.clearTimeout(sidebarScrollTimer);
  if (sidebarScrollFrame) window.cancelAnimationFrame(sidebarScrollFrame);
  sidebarScrollTimer = null;
  sidebarScrollFrame = null;
  sidebarScrollDirection = 0;
  sidebarScrollActive = false;
  hideSidebarScrollIndicator();
}

function showSidebarScrollIndicator(direction, active) {
  const rect = sidebar.getBoundingClientRect();
  sidebarScrollIndicator.textContent = active
    ? `Scrolling ${direction < 0 ? 'up' : 'down'}`
    : `Hold to scroll ${direction < 0 ? 'up' : 'down'}`;
  sidebarScrollIndicator.className = `edge-scroll-indicator ${
    direction < 0 ? 'up' : 'down'
  } ${active ? 'active' : 'pending'}`;
  sidebarScrollIndicator.style.left = `${rect.left}px`;
  sidebarScrollIndicator.style.top = `${
    direction < 0 ? rect.top + 8 : rect.bottom - 44
  }px`;
  sidebarScrollIndicator.style.width = `${rect.width}px`;
}

function hideSidebarScrollIndicator() {
  sidebarScrollIndicator.className = 'hidden';
  sidebarScrollIndicator.textContent = '';
}

function updatePendingTabSwitch(target) {
  const tab = target.closest?.('.doc-tab');
  const id = tab ? Number(tab.dataset.id) : null;
  if (!id || id === workspace.activeId) {
    clearPendingTabSwitch();
    return;
  }
  if (id === pendingTabId) return;

  clearPendingTabSwitch();
  pendingTabId = id;
  tab.classList.add('drop-hover');
  pendingTabTimer = window.setTimeout(() => {
    const next = pendingTabId;
    clearPendingTabSwitch();
    activateDocument(next);
  }, EDGE_HOVER_DELAY);
}

function clearPendingTabSwitch() {
  if (pendingTabTimer) window.clearTimeout(pendingTabTimer);
  pendingTabTimer = null;
  pendingTabId = null;
  tabbar.querySelectorAll('.doc-tab.drop-hover').forEach((tab) => {
    tab.classList.remove('drop-hover');
  });
}

// =========================================================================
// Insert dialog
// =========================================================================

let dialogResolve = null;
let dialogPreviewDoc = null;
let dialogPreviewObserver = null;
let dialogPreviewRenderId = 0;
let dialogSelectedPages = new Set();
let dialogSelectionAnchor = null;
let suppressRangeInputSync = false;

function openInsertDialog({ fileName, srcCount, srcBytes, dropIndex }) {
  return new Promise((resolve) => {
    resetDialogPreview();
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
    updateDialogPreviewStatus(srcCount);

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
    setupDialogPreview(srcBytes, srcCount);
    $('dialog-overlay').classList.remove('hidden');
  });
}

function closeDialog(result) {
  $('dialog-overlay').classList.add('hidden');
  resetDialogPreview();
  const resolve = dialogResolve;
  dialogResolve = null;
  if (resolve) resolve(result);
}

function resetDialogPreview() {
  dialogPreviewRenderId++;
  if (dialogPreviewObserver) {
    dialogPreviewObserver.disconnect();
    dialogPreviewObserver = null;
  }
  if (dialogPreviewDoc) {
    const doc = dialogPreviewDoc;
    dialogPreviewDoc = null;
    doc.destroy().catch(() => {});
  }
  dialogSelectedPages = new Set();
  dialogSelectionAnchor = null;
  suppressRangeInputSync = false;
  $('dlg-preview-grid').replaceChildren();
  $('dlg-page-preview').classList.add('hidden');
}

function setupDialogPreview(srcBytes, srcCount) {
  const grid = $('dlg-preview-grid');
  grid.replaceChildren();
  for (let i = 0; i < srcCount; i++) {
    const page = document.createElement('button');
    page.type = 'button';
    page.className = 'preview-page';
    page.dataset.index = String(i);
    page.setAttribute('aria-pressed', 'false');
    page.setAttribute('aria-label', `Page ${i + 1}`);

    const thumb = document.createElement('span');
    thumb.className = 'preview-thumb';
    const placeholder = document.createElement('span');
    placeholder.className = 'preview-placeholder';
    placeholder.textContent = String(i + 1);
    thumb.appendChild(placeholder);
    page.appendChild(thumb);

    const label = document.createElement('span');
    label.className = 'preview-page-label';
    label.textContent = `Page ${i + 1}`;
    page.appendChild(label);

    page.addEventListener('click', (e) => onDialogPreviewClick(i, e));
    grid.appendChild(page);
  }

  const renderId = ++dialogPreviewRenderId;
  pdfjsLib.getDocument({
    data: srcBytes.slice(),
    ...PDFJS_DOCUMENT_OPTIONS,
  }).promise
    .then((doc) => {
      if (renderId !== dialogPreviewRenderId) {
        doc.destroy().catch(() => {});
        return;
      }
      dialogPreviewDoc = doc;
      observeDialogPreviewPages();
    })
    .catch(() => {
      if (renderId !== dialogPreviewRenderId) return;
      const error = document.createElement('div');
      error.className = 'preview-error';
      error.textContent = 'Preview unavailable';
      grid.replaceChildren(error);
    });
}

function observeDialogPreviewPages() {
  if (dialogPreviewObserver) dialogPreviewObserver.disconnect();
  if (!dialogPreviewDoc) return;
  const grid = $('dlg-preview-grid');
  dialogPreviewObserver = new IntersectionObserver(
    (entries, obs) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        const el = entry.target;
        obs.unobserve(el);
        renderDialogPreviewPage(Number(el.dataset.index), el, dialogPreviewRenderId);
      });
    },
    { root: grid, rootMargin: '180px' }
  );
  grid.querySelectorAll('.preview-page:not(.rendered)').forEach((el) => {
    dialogPreviewObserver.observe(el);
  });
}

async function renderDialogPreviewPage(index, el, renderId) {
  const doc = dialogPreviewDoc;
  if (!doc) return;
  try {
    const page = await doc.getPage(index + 1);
    if (renderId !== dialogPreviewRenderId || doc !== dialogPreviewDoc) return;
    const dpr = window.devicePixelRatio || 1;
    const base = page.getViewport({ scale: 1 });
    const scale = (DIALOG_THUMB_WIDTH / base.width) * dpr;
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const ctx = canvas.getContext('2d');
    await page.render({
      canvasContext: ctx,
      viewport,
      annotationMode: pdfjsLib.AnnotationMode.ENABLE_STORAGE,
    }).promise;
    if (renderId !== dialogPreviewRenderId || doc !== dialogPreviewDoc) return;

    const thumb = el.querySelector('.preview-thumb');
    thumb.replaceChildren(canvas);
    el.classList.add('rendered');
  } catch {
    /* Leave the numbered placeholder in place. */
  }
}

function onDialogPreviewClick(index, e) {
  document.querySelector('input[name="pages"][value="range"]').checked = true;
  updateDialogPagesMode(false);

  if (e.shiftKey && dialogSelectionAnchor !== null) {
    const [a, b] = [dialogSelectionAnchor, index].sort((x, y) => x - y);
    for (let i = a; i <= b; i++) dialogSelectedPages.add(i);
  } else {
    if (dialogSelectedPages.has(index)) dialogSelectedPages.delete(index);
    else dialogSelectedPages.add(index);
    dialogSelectionAnchor = index;
  }
  syncDialogRangeFromSelection();
}

function syncDialogRangeFromSelection() {
  const indices = [...dialogSelectedPages].sort((a, b) => a - b);
  suppressRangeInputSync = true;
  $('dlg-range').value = formatPageRanges(indices);
  suppressRangeInputSync = false;
  $('dlg-range-error').classList.add('hidden');
  applyDialogPreviewSelection();
  updateDialogPreviewStatus(state.dialogSrcCount);
}

function syncDialogSelectionFromRange() {
  if (suppressRangeInputSync) return;
  const indices = parsePageRanges($('dlg-range').value, state.dialogSrcCount);
  dialogSelectedPages = new Set(indices);
  dialogSelectionAnchor = indices.length ? indices[indices.length - 1] : null;
  applyDialogPreviewSelection();
  updateDialogPreviewStatus(state.dialogSrcCount);
}

function applyDialogPreviewSelection() {
  $('dlg-preview-grid').querySelectorAll('.preview-page').forEach((el) => {
    const index = Number(el.dataset.index);
    const selected = dialogSelectedPages.has(index);
    el.classList.toggle('selected', selected);
    el.setAttribute('aria-pressed', selected ? 'true' : 'false');
  });
}

function updateDialogPreviewStatus(srcCount = state.dialogSrcCount) {
  const selected = dialogSelectedPages.size;
  $('dlg-preview-status').textContent = `${selected} of ${srcCount} selected`;
}

function formatPageRanges(indices) {
  if (indices.length === 0) return '';
  const ranges = [];
  let start = indices[0];
  let prev = indices[0];
  for (let i = 1; i < indices.length; i++) {
    const cur = indices[i];
    if (cur === prev + 1) {
      prev = cur;
      continue;
    }
    ranges.push(start === prev ? `${start + 1}` : `${start + 1}-${prev + 1}`);
    start = cur;
    prev = cur;
  }
  ranges.push(start === prev ? `${start + 1}` : `${start + 1}-${prev + 1}`);
  return ranges.join(', ');
}

function updateDialogPagesMode(focusRange = true) {
  const isRange =
    document.querySelector('input[name="pages"]:checked').value === 'range';
  const input = $('dlg-range');
  input.disabled = !isRange;
  $('dlg-page-preview').classList.toggle('hidden', !isRange);
  if (isRange) {
    if (focusRange) input.focus();
    requestAnimationFrame(observeDialogPreviewPages);
  }
}

// Enable range input and page previews only when "Selected pages" is chosen.
document.querySelectorAll('input[name="pages"]').forEach((radio) => {
  radio.addEventListener('change', () => updateDialogPagesMode());
});

$('dlg-range').addEventListener('input', syncDialogSelectionFromRange);

$('dlg-preview-all').addEventListener('click', () => {
  document.querySelector('input[name="pages"][value="range"]').checked = true;
  updateDialogPagesMode(false);
  dialogSelectedPages = new Set(
    Array.from({ length: state.dialogSrcCount }, (_, i) => i)
  );
  dialogSelectionAnchor = 0;
  syncDialogRangeFromSelection();
});

$('dlg-preview-clear').addEventListener('click', () => {
  document.querySelector('input[name="pages"][value="range"]').checked = true;
  updateDialogPagesMode(false);
  dialogSelectedPages = new Set();
  dialogSelectionAnchor = null;
  syncDialogRangeFromSelection();
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
// Comment dialog
// =========================================================================

function openCommentDialog({ title, actionLabel }) {
  return new Promise((resolve) => {
    state.commentDialogOpen = true;
    commentDialogResolve = resolve;
    commentDialogTitle.textContent = title;
    commentConfirm.textContent = actionLabel;
    commentAuthor.value = localStorage.getItem('pdf-workbench-comment-author') || 'Me';
    commentText.value = '';
    commentError.classList.add('hidden');
    commentDialogOverlay.classList.remove('hidden');
    requestAnimationFrame(() => commentText.focus());
  });
}

function closeCommentDialog(result) {
  commentDialogOverlay.classList.add('hidden');
  if (state) state.commentDialogOpen = false;
  const resolve = commentDialogResolve;
  commentDialogResolve = null;
  if (resolve) resolve(result);
}

function confirmCommentDialog() {
  const author = commentAuthor.value.trim() || 'Me';
  const text = commentText.value.trim();
  if (!text) {
    commentError.textContent = 'Enter a comment before continuing.';
    commentError.classList.remove('hidden');
    commentText.focus();
    return;
  }
  localStorage.setItem('pdf-workbench-comment-author', author);
  closeCommentDialog({ author, text });
}

commentCancel.addEventListener('click', () => closeCommentDialog(null));
commentConfirm.addEventListener('click', confirmCommentDialog);
commentDialogOverlay.addEventListener('click', (e) => {
  if (e.target === commentDialogOverlay) closeCommentDialog(null);
});

// =========================================================================
// Keyboard & menu wiring
// =========================================================================

document.addEventListener('keydown', (e) => {
  // Don't hijack typing inside dialog inputs.
  const inField =
    e.target.tagName === 'INPUT' ||
    e.target.tagName === 'TEXTAREA' ||
    e.target.tagName === 'SELECT';
  const dialogOpen = !$('dialog-overlay').classList.contains('hidden');
  const commentDialogOpen = !commentDialogOverlay.classList.contains('hidden');
  const tableDialogOpen = !tableDialogOverlay.classList.contains('hidden');

  if (tableDialogOpen) {
    if (e.key === 'Escape') closeTableDialog(null);
    else if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') confirmTableExport();
    return;
  }

  if (commentDialogOpen) {
    if (e.key === 'Escape') closeCommentDialog(null);
    else if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') confirmCommentDialog();
    return;
  }

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
    if (state.activeTool === 'redact' || state.activeTool === 'comment') {
      setActiveTool('select');
    }
    hideContextMenu();
  }
});

function navigateTo(index) {
  state.currentPage = index;
  state.selection.clear();
  state.selection.add(index);
  state.anchor = index;
  applySelectionStyles();
  scrollMainPageIntoView(index);
  updatePageIndicator();
  updateChrome();
}

redactionApply?.addEventListener('click', () => {
  applyQueuedRedactions().catch(() => {
    window.alert('Could not apply permanent redactions.');
  });
});

toolbarNew.addEventListener('click', newDocument);
toolbarOpen.addEventListener('click', openDocument);
toolbarSave.addEventListener('click', () => saveDocument(false));
toolbarTableExport.addEventListener('click', exportDetectedTableToExcel);
toolbarComment.addEventListener('click', toggleCommentTool);
toolbarRedact.addEventListener('click', () => {
  if (!state.numPages) return;
  setActiveTool(state.activeTool === 'redact' ? 'select' : 'redact');
});
reviewAddComment.addEventListener('click', () => {
  toggleCommentTool();
});
reviewRescan.addEventListener('click', () => runReviewScan(state, state.renderId));
reviewList.addEventListener('click', onReviewAction);
reviewIgnoredList.addEventListener('click', onReviewAction);
tableAddRow.addEventListener('click', addTableRow);
tableRemoveRow.addEventListener('click', removeTableRow);
tableAddColumn.addEventListener('click', addTableColumn);
tableRemoveColumn.addEventListener('click', removeTableColumn);
tableCancel.addEventListener('click', () => closeTableDialog(null));
tableExport.addEventListener('click', confirmTableExport);
tableDialogOverlay.addEventListener('click', (e) => {
  if (e.target === tableDialogOverlay) closeTableDialog(null);
});

async function onReviewAction(e) {
  const button = e.target.closest('button[data-action]');
  if (!button) return;
  const action = button.dataset.action;
  if (
    action === 'show-comment' ||
    action === 'reply-comment' ||
    action === 'toggle-resolved' ||
    action === 'remove-comment' ||
    action === 'remove-reply'
  ) {
    const commentEl = e.target.closest('.comment-reply, .comment-card');
    const comment = commentEl ? findReviewComment(commentEl.dataset.id) : null;
    if (!comment) return;
    if (action === 'reply-comment') {
      await replyToReviewComment(comment);
      return;
    }
    if (action === 'toggle-resolved') {
      await toggleReviewCommentResolved(comment);
      return;
    }
    if (action === 'remove-comment') {
      await removeReviewComment(comment, { includeReplies: true });
      return;
    }
    if (action === 'remove-reply') {
      await removeReviewComment(comment, { includeReplies: false });
      return;
    }

    button.disabled = true;
    const label = button.textContent;
    button.textContent = 'Showing...';
    try {
      await focusReviewComment(comment);
    } finally {
      button.disabled = false;
      button.textContent = label;
    }
    return;
  }

  const card = e.target.closest('.review-card');
  const finding = card ? findReviewFinding(card.dataset.id) : null;
  if (!finding) return;

  if (action === 'show') {
    button.disabled = true;
    const label = button.textContent;
    button.textContent = 'Showing...';
    try {
      await focusReviewFinding(finding);
    } finally {
      button.disabled = false;
      button.textContent = label;
    }
  } else if (action === 'ignore') {
    state.ignoredReviewFindingIds.add(finding.id);
    if (state.focusedReviewFindingId === finding.id) state.focusedReviewFindingId = null;
    state.focusedReviewCommentId = null;
    renderReviewPane();
    renderAllReviewHighlights();
  } else if (action === 'unignore') {
    state.ignoredReviewFindingIds.delete(finding.id);
    renderReviewPane();
    renderAllReviewHighlights();
  } else if (action === 'unredact') {
    button.disabled = true;
    button.textContent = 'Working...';
    try {
      await repairReviewFinding(finding);
    } catch {
      button.disabled = false;
      button.textContent = finding.repairLabel || 'Unredact';
      window.alert('Could not safely unredact this finding.');
    }
  } else if (action === 'redact-properly') {
    button.disabled = true;
    button.textContent = 'Applying...';
    try {
      await applyPermanentRedactions([{ pageIndex: finding.pageIndex, rect: finding.rect }]);
    } catch {
      button.disabled = false;
      button.textContent = 'Redact Properly';
      window.alert('Could not apply a permanent redaction for this finding.');
    }
  }
}

// Application menu actions from main process.
window.api.onMenuAction((action) => {
  switch (action) {
    case 'request-close':
      if (confirmCloseAllDocuments()) window.api.approveClose();
      break;
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
    case 'export-table':
      exportDetectedTableToExcel();
      break;
    case 'add-comment':
      startCommentTool();
      break;
    case 'select-all':
      selectAll();
      break;
    case 'delete':
      deleteSelectedPages();
      break;
    case 'redact':
      setActiveTool('redact');
      break;
  }
});

// Re-render the main page on resize so it keeps fitting the viewer.
let resizeTimer = null;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => renderMainView(), 120);
});

viewer.addEventListener('scroll', () => {
  clearTimeout(viewerScrollTimer);
  viewerScrollTimer = setTimeout(updateCurrentPageFromViewerScroll, 60);
});

// =========================================================================
// Startup
// =========================================================================
(async function init() {
  state = createDocumentState();
  workspace.documents.push(state);
  workspace.activeId = state.id;
  await loadBytes(null, { freshOpen: true });
})();

// --- Debug-only end-to-end test harness -----------------------------------
if (window.api.debug) {
  const snapshot = () => ({
    numPages: state.numPages,
    currentPage: state.currentPage,
    selection: [...state.selection].sort((a, b) => a - b),
    outline: state.outline.map((e) => ({ title: e.title, pageIndex: e.pageIndex })),
    tabs: workspace.documents.map((doc) => ({
      fileName: doc.fileName,
      dirty: doc.dirty,
      active: doc.id === workspace.activeId,
    })),
    thumbCanvases: thumbList.querySelectorAll('.thumb canvas').length,
    mainRendered:
      pageStack.querySelectorAll('.page-view canvas').length > 0 &&
      !pageStage.classList.contains('hidden'),
    reviewStatus: state.reviewStatus,
    reviewFindings: state.reviewFindings.length,
    reviewComments: state.reviewComments.length,
    reviewReplies: (state.reviewComments || []).reduce(
      (total, comment) => total + (comment.replies?.length || 0),
      0
    ),
    reviewResolved: Boolean(state.reviewComments?.[0]?.resolved),
    reviewSummary: reviewSummary.textContent,
    brokenAnnotationImages: [...document.querySelectorAll('.annotationLayer img')].filter(
      (img) => img.complete && img.naturalWidth === 0
    ).length,
  });
  const b64ToBytes = (b64) =>
    Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  const waitForReview = async () => {
    for (let i = 0; i < 40 && state.reviewStatus === 'scanning'; i++) {
      await new Promise((r) => setTimeout(r, 50));
    }
  };
  window.__test = {
    snapshot,
    async load(b64, name = 'Test.pdf') {
      await loadBytes(b64ToBytes(b64), { name, freshOpen: true });
      await waitForReview();
      return snapshot();
    },
    async addComment(text = 'New comment') {
      await mutatePdfComment((bytes) =>
        addPdfComment(bytes, {
          pageIndex: 0,
          rect: { x: 72, y: 300, width: 22, height: 22 },
          contents: text,
          author: 'Smoke Test',
        })
      );
      await waitForReview();
      return snapshot();
    },
    async replyFirstComment(text = 'Reply') {
      const first = state.reviewComments[0];
      if (!first) return snapshot();
      await mutatePdfComment((bytes) =>
        replyToPdfComment(bytes, {
          annotationId: first.annotationId,
          contents: text,
          author: 'Smoke Reply',
        })
      );
      await waitForReview();
      return snapshot();
    },
    async resolveFirstComment(resolved = true) {
      const first = state.reviewComments[0];
      if (!first) return snapshot();
      await mutatePdfComment((bytes) =>
        setPdfCommentResolved(bytes, {
          annotationId: first.annotationId,
          resolved,
        })
      );
      await waitForReview();
      return snapshot();
    },
    async removeFirstComment() {
      const first = state.reviewComments[0];
      if (!first) return snapshot();
      await mutatePdfComment((bytes) =>
        removePdfComment(bytes, {
          annotationId: first.annotationId,
          includeReplies: true,
        })
      );
      await waitForReview();
      return snapshot();
    },
    async extractTable() {
      const extraction = await extractTableFromPdf(state.pdfjsDoc);
      const bytes = createXlsxWorkbook(extraction.table, {
        sheetName: 'Statement',
        firstRowHeader: true,
      });
      return {
        rows: extraction.table.length,
        cols: Math.max(0, ...extraction.table.map((row) => row.length)),
        duplicateHeadersRemoved: extraction.stats.duplicateHeadersRemoved,
        firstRow: extraction.table[0] || [],
        xlsxBytes: bytes.length,
      };
    },
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

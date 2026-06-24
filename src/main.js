// Electron main process: window creation, application menu, and file I/O over IPC.
const { app, BrowserWindow, Menu, dialog, ipcMain, nativeImage } = require('electron');
const fs = require('fs/promises');
const path = require('path');

const APP_NAME = 'PDF Workbench';
const ICON_PATH = path.join(__dirname, '..', 'assets', 'icon.png');

let mainWindow = null;
let closeApproved = false;

app.setName(APP_NAME);

function createWindow() {
  closeApproved = false;
  const icon = nativeImage.createFromPath(ICON_PATH);

  mainWindow = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 720,
    minHeight: 480,
    title: APP_NAME,
    icon,
    backgroundColor: '#2b2b2b',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  mainWindow.on('close', (event) => {
    if (process.env.PDF_EDITOR_SMOKE) return;
    if (closeApproved) return;
    event.preventDefault();
    send('request-close');
  });
  buildMenu();

  if (process.env.PDF_EDITOR_DEBUG) {
    mainWindow.webContents.on('console-message', (_e, level, message) => {
      console.log(`[renderer:${level}] ${message}`);
    });
    mainWindow.webContents.on('render-process-gone', (_e, details) => {
      console.log(`[render-process-gone] ${JSON.stringify(details)}`);
    });
    mainWindow.webContents.on('did-finish-load', () => {
      console.log('[did-finish-load]');
      if (process.env.PDF_EDITOR_SMOKE) {
        runSmokeTest().finally(() => app.quit());
      }
    });
  }
}

function send(action) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('menu-action', action);
  }
}

function buildMenu() {
  const isMac = process.platform === 'darwin';
  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: 'File',
      submenu: [
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: () => send('new') },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
        { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: () => send('save-as') },
        { type: 'separator' },
        { label: 'Export Table to Excel…', accelerator: 'CmdOrCtrl+E', click: () => send('export-table') },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { label: 'Add Comment', click: () => send('add-comment') },
        { label: 'Redact', click: () => send('redact') },
        { type: 'separator' },
        { label: 'Select All Pages', accelerator: 'CmdOrCtrl+A', click: () => send('select-all') },
        { label: 'Delete Selected Pages', accelerator: 'Delete', click: () => send('delete') },
      ],
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' },
        { role: 'toggleDevTools' },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// --- IPC: file operations -------------------------------------------------

ipcMain.handle('open-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Open PDF',
    properties: ['openFile'],
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  const filePath = result.filePaths[0];
  const data = await fs.readFile(filePath);
  return { path: filePath, name: path.basename(filePath), data: new Uint8Array(data) };
});

ipcMain.handle('save-file-dialog', async (_evt, defaultName) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'Save PDF',
    defaultPath: defaultName || 'document.pdf',
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  });
  if (result.canceled || !result.filePath) return null;
  return result.filePath;
});

ipcMain.handle('save-xlsx-dialog', async (_evt, defaultName) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'Export Table to Excel',
    defaultPath: defaultName || 'table.xlsx',
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }],
  });
  if (result.canceled || !result.filePath) return null;
  return result.filePath;
});

ipcMain.handle('write-file', async (_evt, filePath, data) => {
  await fs.writeFile(filePath, Buffer.from(data));
  return { ok: true, name: path.basename(filePath) };
});

ipcMain.handle('read-file', async (_evt, filePath) => {
  const data = await fs.readFile(filePath);
  return new Uint8Array(data);
});

ipcMain.on('set-title', (_evt, title) => {
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.setTitle(title);
});

ipcMain.on('close-approved', () => {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  closeApproved = true;
  mainWindow.close();
});

// --- Debug-only smoke test driver -----------------------------------------
async function makeSrcPdfBase64(n) {
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 0; i < n; i++) {
    const p = doc.addPage([300, 400]);
    p.drawText(`Source page ${i + 1}`, { x: 30, y: 200, size: 20, font, color: rgb(0, 0, 0) });
  }
  const bytes = await doc.save();
  return Buffer.from(bytes).toString('base64');
}

async function makeCommentPdfBase64() {
  const {
    PDFDocument,
    PDFHexString,
    PDFName,
    PDFString,
    StandardFonts,
    rgb,
  } = require('pdf-lib');
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const page = doc.addPage([360, 460]);
  page.drawText('Comment fixture', { x: 40, y: 390, size: 22, font, color: rgb(0, 0, 0) });
  page.drawText('This page has a PDF sticky-note annotation.', {
    x: 40,
    y: 350,
    size: 12,
    font,
    color: rgb(0, 0, 0),
  });

  const annotation = doc.context.obj({
    Type: PDFName.of('Annot'),
    Subtype: PDFName.of('Text'),
    Rect: [42, 300, 62, 320],
    Contents: PDFHexString.fromText('Please review this section.'),
    T: PDFHexString.fromText('Smoke Test'),
    M: PDFString.of('D:20260623120000Z'),
    C: [1, 0.85, 0],
    Open: false,
  });
  const annotationRef = doc.context.register(annotation);
  page.node.set(PDFName.of('Annots'), doc.context.obj([annotationRef]));

  const bytes = await doc.save();
  return Buffer.from(bytes).toString('base64');
}

async function makeStatementPdfBase64() {
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const columns = [
    { label: 'Date', x: 44 },
    { label: 'Description', x: 112 },
    { label: 'Debit', x: 322 },
    { label: 'Credit', x: 398 },
    { label: 'Balance', x: 478 },
  ];
  const pages = [
    [
      ['2026-01-02', 'Opening balance', '', '1200.00', '1200.00'],
      ['2026-01-03', 'Groceries', '48.15', '', '1151.85'],
      ['2026-01-04', 'Salary', '', '2400.00', '3551.85'],
    ],
    [
      ['2026-01-07', 'Rent', '1300.00', '', '2251.85'],
      ['2026-01-08', 'Coffee', '3.40', '', '2248.45'],
      ['2026-01-09', 'Transfer', '500.00', '', '1748.45'],
    ],
  ];

  pages.forEach((rows, pageIndex) => {
    const page = doc.addPage([612, 792]);
    page.drawText('Statement fixture', {
      x: 44,
      y: 740,
      size: 18,
      font: bold,
      color: rgb(0, 0, 0),
    });
    page.drawText(`Page ${pageIndex + 1}`, {
      x: 520,
      y: 742,
      size: 10,
      font,
      color: rgb(0, 0, 0),
    });

    let y = 690;
    columns.forEach((column) => {
      page.drawText(column.label, {
        x: column.x,
        y,
        size: 11,
        font: bold,
        color: rgb(0, 0, 0),
      });
    });
    y -= 26;
    rows.forEach((row) => {
      row.forEach((value, columnIndex) => {
        if (!value) return;
        page.drawText(value, {
          x: columns[columnIndex].x,
          y,
          size: 10,
          font,
          color: rgb(0, 0, 0),
        });
      });
      y -= 24;
    });
  });

  const bytes = await doc.save();
  return Buffer.from(bytes).toString('base64');
}

async function runSmokeTest() {
  const wc = mainWindow.webContents;
  const run = (expr) => wc.executeJavaScript(expr, true);
  const log = (label, obj) => console.log(`[smoke] ${label}: ${JSON.stringify(obj)}`);
  try {
    const src5 = await makeSrcPdfBase64(5);
    const src3 = await makeSrcPdfBase64(3);
    const src1 = await makeSrcPdfBase64(1);
    const commentPdf = await makeCommentPdfBase64();
    const statementPdf = await makeStatementPdfBase64();

    log('start', await run('window.__test.snapshot()'));

    // Insert all 5 pages into the empty document with a bookmark name.
    let s = await run(`window.__test.insert(${JSON.stringify(src5)}, [0,1,2,3,4], 0, 'Chapter One')`);
    log('after insert 5', s);

    // Insert a selection (src pages 1 & 3) after current page with a name.
    s = await run(`window.__test.insert(${JSON.stringify(src3)}, [0,2], 2, 'Section Two')`);
    log('after insert selection', s);

    // Delete pages 0 and 1.
    s = await run('window.__test.del([0,1])');
    log('after delete', s);
    const workflowOk =
      s.numPages === 5 &&
      s.outline.some((e) => e.title === 'Section Two') &&
      s.thumbCanvases > 0 &&
      s.mainRendered;

    s = await run(`window.__test.load(${JSON.stringify(commentPdf)}, 'Commented.pdf')`);
    log('after comment load', s);
    const commentReadOk =
      s.numPages === 1 &&
      s.reviewComments === 1 &&
      s.brokenAnnotationImages === 0 &&
      /comment/i.test(s.reviewSummary) &&
      s.thumbCanvases > 0 &&
      s.mainRendered;

    s = await run(`window.__test.load(${JSON.stringify(src1)}, 'Comment Lifecycle.pdf')`);
    log('comment lifecycle start', s);
    s = await run('window.__test.addComment("Lifecycle parent")');
    log('after add comment', s);
    const addOk =
      s.reviewComments === 1 &&
      s.reviewReplies === 0 &&
      !s.reviewResolved &&
      s.brokenAnnotationImages === 0;
    s = await run('window.__test.replyFirstComment("Lifecycle reply")');
    log('after reply comment', s);
    const replyOk = s.reviewComments === 1 && s.reviewReplies === 1;
    s = await run('window.__test.resolveFirstComment(true)');
    log('after resolve comment', s);
    const resolveOk = s.reviewComments === 1 && s.reviewResolved;
    s = await run('window.__test.removeFirstComment()');
    log('after remove comment', s);
    const removeOk = s.reviewComments === 0 && s.reviewReplies === 0;

    s = await run(`window.__test.load(${JSON.stringify(statementPdf)}, 'Statement.pdf')`);
    log('statement load', s);
    const table = await run('window.__test.extractTable()');
    log('table extraction', table);
    const tableOk =
      table.rows === 7 &&
      table.cols === 5 &&
      table.duplicateHeadersRemoved === 1 &&
      table.xlsxBytes > 1200 &&
      table.firstRow.join('|') === 'Date|Description|Debit|Credit|Balance';

    // Verdict
    const ok = workflowOk && commentReadOk && addOk && replyOk && resolveOk && removeOk && tableOk;
    console.log(`[smoke] RESULT: ${ok ? 'PASS' : 'FAIL'} (comments=${s.reviewComments}, thumbs rendered=${s.thumbCanvases}, main=${s.mainRendered})`);
  } catch (e) {
    console.log(`[smoke] ERROR: ${e && e.message}`);
  }
}

app.whenReady().then(() => {
  if (process.platform === 'darwin' && !nativeImage.createFromPath(ICON_PATH).isEmpty()) {
    app.dock.setIcon(ICON_PATH);
  }
  app.setAboutPanelOptions({
    applicationName: APP_NAME,
    applicationVersion: app.getVersion(),
    iconPath: ICON_PATH,
  });
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

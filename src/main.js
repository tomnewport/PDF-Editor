// Electron main process: window creation, application menu, and file I/O over IPC.
const { app, BrowserWindow, Menu, dialog, ipcMain } = require('electron');
const fs = require('fs/promises');
const path = require('path');

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 720,
    minHeight: 480,
    title: 'PDF Editor',
    backgroundColor: '#2b2b2b',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
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
        isMac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    {
      label: 'Edit',
      submenu: [
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

async function runSmokeTest() {
  const wc = mainWindow.webContents;
  const run = (expr) => wc.executeJavaScript(expr, true);
  const log = (label, obj) => console.log(`[smoke] ${label}: ${JSON.stringify(obj)}`);
  try {
    const src5 = await makeSrcPdfBase64(5);
    const src3 = await makeSrcPdfBase64(3);

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

    // Verdict
    const ok =
      s.numPages === 5 &&
      s.outline.some((e) => e.title === 'Section Two') &&
      s.thumbCanvases > 0 &&
      s.mainRendered;
    console.log(`[smoke] RESULT: ${ok ? 'PASS' : 'FAIL'} (thumbs rendered=${s.thumbCanvases}, main=${s.mainRendered})`);
  } catch (e) {
    console.log(`[smoke] ERROR: ${e && e.message}`);
  }
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

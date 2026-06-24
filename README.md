# PDF Editor

A desktop (Electron) app for assembling and editing PDF files. Open an existing
PDF or start a new, empty one, then build it up by dragging other PDFs in, and
reorder or remove pages with a live thumbnail preview.

## Download

Pre-built installers are published on the
[**latest release**](https://github.com/tomnewport/PDF-Editor/releases/latest)
page:

| Platform | File |
| -------- | ---- |
| Windows  | `.exe` installer (NSIS) |
| macOS    | `.dmg` |
| Linux    | `.AppImage` or `.deb` |

Installers are built and published automatically by
[GitHub Actions](.github/workflows/release.yml) whenever a version tag (`v*`) is
pushed. The macOS and Windows builds are unsigned, so you may need to approve
them past Gatekeeper / SmartScreen on first launch.

## Features

- **Open or create** PDFs. A new document starts empty (zero pages).
- **Drag & drop PDFs to add pages.** Dropping a PDF opens a dialog that lets you:
  - add the **entire PDF** or just a **selection of pages** (e.g. `1-3, 5, 8-10`);
  - choose the **insert position** (before / after the current page);
  - optionally **name the inserted section**, which adds a bookmark to the
    document's outline.
- **Drop directly into the thumbnail sidebar.** While dragging over the
  sidebar, an insertion indicator appears at the gap nearest the cursor, and the
  drop dialog pre-selects that position.
- **Thumbnail preview sidebar.** Click a thumbnail to jump to that page.
- **Multi-select** pages with `Shift` (range) and `Ctrl`/`Cmd`/`Alt` (toggle).
- **Delete pages** from the right-click context menu (in either the sidebar or
  the main page view), or with the `Delete` key. Removing the last page returns
  the document to the empty state.
- **Save / Save As** to write the edited PDF back to disk.

## Keyboard shortcuts

| Action            | Shortcut                 |
| ----------------- | ------------------------ |
| New               | `Ctrl/Cmd+N`             |
| Open              | `Ctrl/Cmd+O`             |
| Save              | `Ctrl/Cmd+S`             |
| Save As           | `Ctrl/Cmd+Shift+S`       |
| Select all pages  | `Ctrl/Cmd+A`             |
| Delete selection  | `Delete` / `Backspace`   |
| Navigate pages    | Arrow keys               |

## Getting started

```bash
npm install
npm start        # builds the renderer bundle and launches Electron
```

`npm start` runs `npm run build` first, which bundles the renderer with esbuild
and copies static assets and the pdf.js worker into `dist/`.

## Project layout

```
build.js            esbuild bundling + asset copy into dist/
src/main.js         Electron main process: window, menu, file I/O over IPC
src/preload.js      contextBridge API exposed to the renderer
src/renderer.js     UI: rendering, selection, drag & drop, dialog
src/pdf-ops.js      pdf-lib operations: insert/delete pages, outline bookmarks
public/             index.html + styles.css (copied to dist/ at build time)
```

## How it works

- **Rendering** uses [pdf.js](https://mozilla.github.io/pdf.js/) (`pdfjs-dist`)
  to draw page thumbnails and the main preview to `<canvas>` elements.
- **Editing** uses [pdf-lib](https://pdf-lib.js.org/) to copy pages between
  documents, insert/remove pages, and write the outline (bookmarks).
- Because a genuine zero-page PDF cannot be reliably round-tripped, the empty
  document is represented in memory as the absence of bytes and is only
  serialized to a real (zero-page) PDF when saved.

## Debugging

Set `PDF_EDITOR_DEBUG=1` to forward renderer console output to the terminal and
expose a small test harness. Set `PDF_EDITOR_SMOKE=1` as well to run an
automated end-to-end check (insert / select / delete) on launch and exit.

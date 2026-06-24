import * as pdfjsLib from 'pdfjs-dist';

const ROW_GAP_FACTOR = 0.58;
const COLUMN_CLUSTER_TOLERANCE = 18;
const MIN_COLUMN_SUPPORT = 2;

export async function extractTableFromPdf(pdfjsDoc) {
  if (!pdfjsDoc?.numPages) {
    return emptyExtraction();
  }

  const pageRows = [];
  for (let pageIndex = 0; pageIndex < pdfjsDoc.numPages; pageIndex++) {
    const page = await pdfjsDoc.getPage(pageIndex + 1);
    const rows = await extractPageRows(page, pageIndex);
    pageRows.push(...rows);
  }

  const candidateRows = pageRows.filter((row) => row.fragments.length >= 2);
  if (candidateRows.length === 0) {
    return emptyExtraction({ pages: pdfjsDoc.numPages });
  }

  const anchors = inferColumnAnchors(candidateRows);
  if (anchors.length < 2) {
    return emptyExtraction({
      pages: pdfjsDoc.numPages,
      stats: {
        sourceRows: pageRows.length,
        candidateRows: candidateRows.length,
        selectedRows: 0,
        duplicateHeadersRemoved: 0,
      },
    });
  }

  const assignedRows = candidateRows
    .map((row) => assignRowToColumns(row, anchors))
    .filter((row) => filledCellCount(row.cells) >= 2);

  const selectedRows = selectLargestTableRun(assignedRows);
  const { rows, duplicateHeadersRemoved } = removeRepeatedHeaders(selectedRows);
  const table = normalizeTable(rows.map((row) => row.cells));
  const pageIndices = rows.map((row) => row.pageIndex);

  return {
    table,
    rows,
    pages: pdfjsDoc.numPages,
    pageSpan: pageIndices.length
      ? {
          start: Math.min(...pageIndices),
          end: Math.max(...pageIndices),
        }
      : null,
    columnAnchors: anchors,
    stats: {
      sourceRows: pageRows.length,
      candidateRows: candidateRows.length,
      selectedRows: selectedRows.length,
      duplicateHeadersRemoved,
    },
  };
}

async function extractPageRows(page, pageIndex) {
  const textContent = await page.getTextContent();
  const viewport = page.getViewport({ scale: 1 });
  const items = textContent.items
    .map((item) => textItemToFragment(item, viewport))
    .filter((item) => item && item.text);

  if (items.length === 0) return [];

  const medianHeight = median(items.map((item) => item.height).filter((value) => value > 0)) || 10;
  const tolerance = Math.max(2.5, medianHeight * ROW_GAP_FACTOR);
  const rowBuckets = [];

  for (const item of items.sort((a, b) => a.y - b.y || a.x - b.x)) {
    let bucket = rowBuckets.find(
      (row) => Math.abs(row.y - item.y) <= Math.max(tolerance, row.height * 0.44)
    );
    if (!bucket) {
      bucket = {
        pageIndex,
        y: item.y,
        height: item.height,
        items: [],
      };
      rowBuckets.push(bucket);
    }
    bucket.items.push(item);
    bucket.y = weightedAverage(bucket.items.map((entry) => entry.y));
    bucket.height = Math.max(bucket.height, item.height);
  }

  return rowBuckets
    .sort((a, b) => a.y - b.y)
    .map((row, rowIndex) => {
      const fragments = mergeNearbyItems(row.items, row.height);
      return {
        pageIndex,
        rowIndex,
        y: row.y,
        height: row.height,
        fragments,
      };
    })
    .filter((row) => row.fragments.length > 0);
}

function textItemToFragment(item, viewport) {
  const text = String(item.str || '').replace(/\s+/g, ' ').trim();
  if (!text) return null;

  const transform = pdfjsLib.Util.transform(viewport.transform, item.transform);
  const x = transform[4];
  const y = transform[5];
  const height = Math.max(1, Math.hypot(transform[2], transform[3]) || item.height || 10);
  const width = Math.max(1, Math.abs(item.width || transform[0] || text.length * height * 0.45));

  return {
    text,
    x,
    y,
    width,
    height,
    right: x + width,
  };
}

function mergeNearbyItems(items, rowHeight) {
  const sorted = [...items].sort((a, b) => a.x - b.x);
  const fragments = [];
  const smallGap = Math.max(3, rowHeight * 0.55);

  for (const item of sorted) {
    const previous = fragments[fragments.length - 1];
    if (previous && item.x - previous.right <= smallGap) {
      previous.text = joinCellText(previous.text, item.text);
      previous.right = Math.max(previous.right, item.right);
      previous.width = previous.right - previous.x;
      previous.height = Math.max(previous.height, item.height);
    } else {
      fragments.push({ ...item });
    }
  }

  return fragments;
}

function inferColumnAnchors(rows) {
  const positions = rows.flatMap((row) => row.fragments.map((fragment) => fragment.x));
  const clusters = clusterNumbers(positions, COLUMN_CLUSTER_TOLERANCE)
    .map((cluster) => ({
      x: weightedAverage(cluster.values),
      count: cluster.values.length,
    }))
    .sort((a, b) => a.x - b.x);

  const minSupport = Math.max(
    MIN_COLUMN_SUPPORT,
    Math.ceil(rows.length * 0.04)
  );
  let supported = clusters.filter((cluster) => cluster.count >= minSupport);

  if (supported.length < 2) {
    supported = [...clusters].sort((a, b) => b.count - a.count).slice(0, 8);
  }

  return supported
    .sort((a, b) => a.x - b.x)
    .map((cluster) => Math.round(cluster.x * 100) / 100);
}

function assignRowToColumns(row, anchors) {
  const cells = Array.from({ length: anchors.length }, () => '');
  for (const fragment of row.fragments) {
    const index = nearestAnchorIndex(fragment.x, anchors);
    cells[index] = joinCellText(cells[index], fragment.text);
  }
  return {
    ...row,
    cells: cells.map((cell) => cell.trim()),
  };
}

function nearestAnchorIndex(x, anchors) {
  let bestIndex = 0;
  let bestDistance = Infinity;
  anchors.forEach((anchor, index) => {
    const distance = Math.abs(anchor - x);
    if (distance < bestDistance) {
      bestIndex = index;
      bestDistance = distance;
    }
  });
  return bestIndex;
}

function selectLargestTableRun(rows) {
  if (rows.length <= 3) return rows;

  const byPage = new Map();
  for (const row of rows) {
    if (!byPage.has(row.pageIndex)) byPage.set(row.pageIndex, []);
    byPage.get(row.pageIndex).push(row);
  }

  const bestRuns = [];
  for (const pageRows of byPage.values()) {
    const runs = splitPageRuns(pageRows);
    const best = runs.sort((a, b) => scoreRun(b) - scoreRun(a))[0];
    if (best) bestRuns.push(best);
  }

  if (bestRuns.length <= 1) return bestRuns[0] || rows;
  const strongest = [...bestRuns].sort((a, b) => scoreRun(b) - scoreRun(a))[0];
  return bestRuns
    .filter((run) => compatibleColumnProfile(run, strongest))
    .flat();
}

function splitPageRuns(rows) {
  const runs = [];
  let current = [];
  for (const row of rows) {
    const previous = current[current.length - 1];
    const gap = previous ? row.y - previous.y : 0;
    const continuation =
      previous && gap <= Math.max(34, Math.max(previous.height, row.height) * 2.4);
    if (!previous || continuation) {
      current.push(row);
    } else {
      runs.push(current);
      current = [row];
    }
  }
  if (current.length) runs.push(current);
  return runs;
}

function compatibleColumnProfile(run, reference) {
  const a = columnProfile(run);
  const b = columnProfile(reference);
  if (!a.size || !b.size) return false;
  const overlap = [...a].filter((index) => b.has(index)).length;
  return overlap / Math.min(a.size, b.size) >= 0.6;
}

function columnProfile(run) {
  const counts = new Map();
  run.forEach((row) => {
    row.cells.forEach((cell, index) => {
      if (cell) counts.set(index, (counts.get(index) || 0) + 1);
    });
  });
  const threshold = Math.max(1, Math.ceil(run.length * 0.35));
  return new Set(
    [...counts.entries()]
      .filter(([, count]) => count >= threshold)
      .map(([index]) => index)
  );
}

function scoreRun(run) {
  const filled = run.reduce((total, row) => total + filledCellCount(row.cells), 0);
  const numeric = run.reduce(
    (total, row) => total + row.cells.filter((cell) => looksNumeric(cell)).length,
    0
  );
  const pageBonus = new Set(run.map((row) => row.pageIndex)).size * 1.5;
  return run.length * 4 + filled + numeric * 1.2 + pageBonus;
}

function removeRepeatedHeaders(rows) {
  const output = [];
  const headerSignatures = new Set();
  let duplicateHeadersRemoved = 0;

  rows.forEach((row, index) => {
    const signature = rowSignature(row.cells);
    const canBeHeader = looksLikeHeader(row.cells);
    const nearPageTop = row.rowIndex <= 6;
    const repeatedHeader =
      index > 0 && canBeHeader && nearPageTop && headerSignatures.has(signature);

    if (repeatedHeader) {
      duplicateHeadersRemoved++;
      return;
    }

    output.push(row);
    if (canBeHeader && (index < 3 || nearPageTop)) {
      headerSignatures.add(signature);
    }
  });

  return { rows: output, duplicateHeadersRemoved };
}

function normalizeTable(rows) {
  const width = Math.max(0, ...rows.map((row) => row.length));
  const padded = rows.map((row) =>
    Array.from({ length: width }, (_, index) => String(row[index] || '').trim())
  );

  let lastColumn = width - 1;
  while (
    lastColumn >= 0 &&
    padded.every((row) => !row[lastColumn])
  ) {
    lastColumn--;
  }

  return padded.map((row) => row.slice(0, lastColumn + 1));
}

function looksLikeHeader(cells) {
  const filled = cells.filter(Boolean);
  if (filled.length < 2) return false;
  const alphaCount = filled.filter((cell) => /[A-Za-z]/.test(cell)).length;
  const numericCount = filled.filter(looksNumeric).length;
  return alphaCount >= Math.max(2, filled.length - 1) && numericCount <= 1;
}

function looksNumeric(cell) {
  const text = String(cell || '').trim();
  return /^[-+]?[$£€]?\s*\(?\d[\d,]*(?:\.\d+)?\)?%?$/.test(text);
}

function rowSignature(cells) {
  return cells
    .map((cell) => String(cell || '').toLowerCase().replace(/[^a-z0-9]+/g, ''))
    .filter(Boolean)
    .join('|');
}

function filledCellCount(cells) {
  return cells.filter((cell) => String(cell || '').trim()).length;
}

function clusterNumbers(values, tolerance) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  const clusters = [];
  for (const value of sorted) {
    const current = clusters[clusters.length - 1];
    if (!current || Math.abs(weightedAverage(current.values) - value) > tolerance) {
      clusters.push({ values: [value] });
    } else {
      current.values.push(value);
    }
  }
  return clusters;
}

function joinCellText(a, b) {
  const left = String(a || '').trim();
  const right = String(b || '').trim();
  if (!left) return right;
  if (!right) return left;
  if (/[-/([{]$/.test(left) || /^[,.;:)\]}%]/.test(right)) return `${left}${right}`;
  return `${left} ${right}`;
}

function weightedAverage(values) {
  if (!values.length) return 0;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function median(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2
    ? sorted[mid]
    : (sorted[mid - 1] + sorted[mid]) / 2;
}

function emptyExtraction(overrides = {}) {
  const { stats: overrideStats, ...rest } = overrides;
  const stats = {
    sourceRows: 0,
    candidateRows: 0,
    selectedRows: 0,
    duplicateHeadersRemoved: 0,
    ...(overrideStats || {}),
  };
  return {
    table: [],
    rows: [],
    pages: 0,
    pageSpan: null,
    columnAnchors: [],
    stats,
    ...rest,
  };
}

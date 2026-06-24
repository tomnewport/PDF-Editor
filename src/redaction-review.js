import * as pdfjsLib from 'pdfjs-dist';

const { OPS } = pdfjsLib;
const MIN_RECT_SIZE = 3;
const PAGE_EDGE_GUARD = 36;
const PAGE_COUNTER_RE = /^\d+\s+(?:of|\/)\s+\d+$/i;

export async function analyzeSuspiciousRedactions(pdfjsDoc) {
  if (!pdfjsDoc) return [];

  const findings = [];
  for (let pageIndex = 0; pageIndex < pdfjsDoc.numPages; pageIndex++) {
    const page = await pdfjsDoc.getPage(pageIndex + 1);
    const [opList, textContent] = await Promise.all([
      page.getOperatorList(),
      page.getTextContent({ disableCombineTextItems: false }),
    ]);
    findings.push(...analyzePage(pageIndex, opList, textContent, pageBox(page)));
  }

  return findings;
}

function analyzePage(pageIndex, opList, textContent, pageBounds) {
  const { rectangles, textRuns } = extractOperatorFeatures(opList);
  const textItems = assignTextSequences(textContent.items || [], textRuns);
  const findings = [];

  for (const rect of rectangles) {
    const covered = textItems.filter(
      (item) =>
        item.seq < rect.seq &&
        hasMeaningfulText(item.text) &&
        intersects(item.bbox, rect.bbox, 0.15)
    );
    if (covered.length === 0) continue;
    if (isPageCounterArtifact(covered, rect.bbox, pageBounds)) continue;
    if (hasLaterVisibleText(textItems, covered, rect.bbox, rect.seq)) continue;

    const textRect = coveredTextBBox(covered, rect.bbox);
    const sampleText = summarizeText(extractCoveredText(covered, rect.bbox));
    findings.push({
      id: makeFindingId('overlay', pageIndex, rect.bbox, sampleText),
      kind: 'overlay-redaction',
      severity: 'high',
      title: 'Possible fake redaction',
      pageIndex,
      rect: rect.bbox,
      textRect,
      sampleText,
      annotationId: rect.annotation?.id || null,
      annotationRect: rect.annotation?.rect || null,
      reason:
        rect.annotation
          ? 'A dark annotation rectangle is drawn over selectable text in the same area.'
          : 'A solid dark rectangle is drawn after selectable text in the same area.',
      repairable: true,
      repairLabel: 'Unredact',
    });
  }

  for (const item of textItems) {
    if (!hasMeaningfulText(item.text)) continue;
    const invisible = item.renderingMode === 3;
    const sameAsPaper = isNearWhite(item.color);
    if (!invisible && !sameAsPaper) continue;
    if (hasLaterVisibleText(textItems, [item], item.bbox, item.seq)) continue;

    const reason = invisible
      ? 'Text uses PDF rendering mode 3, which extracts but does not paint.'
      : 'Text is painted nearly white, which may hide it on a white page.';
    findings.push({
      id: makeFindingId('hidden-text', pageIndex, item.bbox, item.text),
      kind: 'hidden-text',
      severity: invisible ? 'high' : 'medium',
      title: invisible ? 'Invisible text remains in PDF' : 'Text matches page background',
      pageIndex,
      rect: item.bbox,
      textRect: item.bbox,
      sampleText: summarizeText(item.text),
      reason,
      repairable: true,
      repairLabel: 'Reveal text',
    });
  }

  return findings;
}

function extractOperatorFeatures(opList) {
  const rectangles = [];
  const textRuns = [];
  const stack = [];
  const annotationStack = [];
  let state = defaultGraphicsState();
  let pendingPath = null;

  for (let i = 0; i < opList.fnArray.length; i++) {
    const fn = opList.fnArray[i];
    const args = opList.argsArray[i];

    switch (fn) {
      case OPS.beginAnnotation: {
        annotationStack.push(cloneGraphicsState(state));
        const next = cloneGraphicsState(state);
        const [, rect, transform] = Array.isArray(args) ? args : [];
        next.annotation = {
          id: normalizeAnnotationId(args?.[0]),
          rect: rectFromArray(rect),
        };
        if (isMatrix(transform)) next.ctm = multiplyMatrix(next.ctm, transform);
        state = next;
        pendingPath = null;
        break;
      }
      case OPS.endAnnotation:
        state = annotationStack.pop() || defaultGraphicsState();
        pendingPath = null;
        break;
      case OPS.save:
        stack.push(cloneGraphicsState(state));
        break;
      case OPS.restore:
        state = stack.pop() || defaultGraphicsState();
        pendingPath = null;
        break;
      case OPS.transform:
        state.ctm = multiplyMatrix(state.ctm, args);
        break;
      case OPS.setFillRGBColor:
        state.fillColor = rgbFromArgs(args);
        break;
      case OPS.setFillGray:
        state.fillColor = grayFromArgs(args);
        break;
      case OPS.setFillCMYKColor:
        state.fillColor = cmykFromArgs(args);
        break;
      case OPS.setFillColor:
      case OPS.setFillColorN:
        state.fillColor = genericColorFromArgs(args, state.fillColor);
        break;
      case OPS.setTextRenderingMode:
        state.textRenderingMode = Array.isArray(args) ? args[0] : args;
        break;
      case OPS.constructPath:
        pendingPath = pathBBox(args, state.ctm);
        break;
      case OPS.fill:
      case OPS.eoFill:
      case OPS.fillStroke:
      case OPS.eoFillStroke:
      case OPS.closeFillStroke:
      case OPS.closeEOFillStroke:
        if (pendingPath && isDark(state.fillColor) && isCandidateRect(pendingPath)) {
          rectangles.push({
            seq: i,
            bbox: pendingPath,
            color: state.fillColor,
            annotation: state.annotation,
          });
        }
        pendingPath = null;
        break;
      case OPS.showText:
      case OPS.showSpacedText:
      case OPS.nextLineShowText:
      case OPS.nextLineSetSpacingShowText: {
        const text = extractText(args);
        if (text) {
          textRuns.push({
            seq: i,
            text,
            color: state.fillColor,
            renderingMode: state.textRenderingMode,
          });
        }
        break;
      }
      default:
        break;
    }
  }

  return { rectangles, textRuns };
}

function defaultGraphicsState() {
  return {
    ctm: [1, 0, 0, 1, 0, 0],
    fillColor: [0, 0, 0],
    textRenderingMode: 0,
    annotation: null,
  };
}

function cloneGraphicsState(state) {
  return {
    ctm: state.ctm.slice(),
    fillColor: state.fillColor.slice(),
    textRenderingMode: state.textRenderingMode,
    annotation: state.annotation ? { ...state.annotation } : null,
  };
}

function assignTextSequences(items, textRuns) {
  const out = [];
  let runCursor = 0;
  for (const item of items) {
    const text = item.str || '';
    if (!text.trim()) continue;

    let runIndex = -1;
    for (let i = runCursor; i < textRuns.length; i++) {
      if (textMatches(textRuns[i].text, text)) {
        runIndex = i;
        break;
      }
    }
    if (runIndex === -1 && runCursor < textRuns.length) runIndex = runCursor;
    const run = textRuns[runIndex] || {};
    if (runIndex >= 0) runCursor = runIndex + 1;

    out.push({
      text,
      bbox: textItemBBox(item),
      seq: run.seq ?? Number.MAX_SAFE_INTEGER,
      color: run.color || [0, 0, 0],
      renderingMode: run.renderingMode ?? 0,
    });
  }
  return out;
}

function textItemBBox(item) {
  const transform = item.transform || [1, 0, 0, 1, 0, 0];
  const x = transform[4] || 0;
  const y = transform[5] || 0;
  const width = Math.abs(item.width || 0);
  const height =
    Math.abs(item.height || 0) || Math.hypot(transform[2] || 0, transform[3] || 0);
  return {
    x,
    y,
    width: Math.max(0, width),
    height: Math.max(0, height),
  };
}

function pageBox(page) {
  const view = page.view || [0, 0, 0, 0];
  return rectFromArray(view);
}

function rectFromArray(values) {
  if (!Array.isArray(values) || values.length < 4) return null;
  const [x1, y1, x2, y2] = values.map((v) => Number(v) || 0);
  const minX = Math.min(x1, x2);
  const minY = Math.min(y1, y2);
  return {
    x: minX,
    y: minY,
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1),
  };
}

function pathBBox(args, ctm) {
  const minMax = Array.isArray(args?.[2]) ? args[2] : null;
  let points = [];
  if (minMax && minMax.length === 4) {
    const [x1, y1, x2, y2] = minMax;
    points = [
      [x1, y1],
      [x1, y2],
      [x2, y1],
      [x2, y2],
    ];
  } else {
    const coords = Array.isArray(args?.[1]) ? args[1] : [];
    for (let i = 0; i + 1 < coords.length; i += 2) {
      points.push([coords[i], coords[i + 1]]);
    }
  }
  if (points.length === 0) return null;

  const transformed = points.map(([x, y]) => transformPoint(ctm, x, y));
  const xs = transformed.map((p) => p[0]);
  const ys = transformed.map((p) => p[1]);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}

function isCandidateRect(bbox) {
  return bbox.width > MIN_RECT_SIZE && bbox.height > MIN_RECT_SIZE;
}

function multiplyMatrix(m1, m2) {
  const [a1, b1, c1, d1, e1, f1] = m1;
  const [a2, b2, c2, d2, e2, f2] = m2;
  return [
    a1 * a2 + c1 * b2,
    b1 * a2 + d1 * b2,
    a1 * c2 + c1 * d2,
    b1 * c2 + d1 * d2,
    a1 * e2 + c1 * f2 + e1,
    b1 * e2 + d1 * f2 + f1,
  ];
}

function transformPoint(m, x, y) {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function isMatrix(value) {
  return Array.isArray(value) && value.length === 6 && value.every(Number.isFinite);
}

function rgbFromArgs(args) {
  const values = Array.from(args || []).slice(0, 3);
  if (values.length < 3) return [0, 0, 0];
  return values.map((v) => clampColor(v > 1 ? v / 255 : v));
}

function grayFromArgs(args) {
  const raw = Array.isArray(args) ? args[0] : args;
  const v = clampColor((raw || 0) > 1 ? raw / 255 : raw || 0);
  return [v, v, v];
}

function cmykFromArgs(args) {
  const [c = 0, m = 0, y = 0, k = 0] = Array.from(args || []).map((v) =>
    clampColor(v > 1 ? v / 255 : v)
  );
  return [
    clampColor(1 - Math.min(1, c + k)),
    clampColor(1 - Math.min(1, m + k)),
    clampColor(1 - Math.min(1, y + k)),
  ];
}

function genericColorFromArgs(args, fallback) {
  const values = flattenNumbers(args);
  if (values.length === 1) return grayFromArgs(values);
  if (values.length === 3) return rgbFromArgs(values);
  if (values.length >= 4) return cmykFromArgs(values.slice(0, 4));
  return fallback || [0, 0, 0];
}

function flattenNumbers(value) {
  const out = [];
  const visit = (item) => {
    if (item == null) return;
    if (typeof item === 'number') out.push(item);
    else if (Array.isArray(item) || ArrayBuffer.isView(item)) item.forEach(visit);
    else if (typeof item === 'object') Object.values(item).forEach(visit);
  };
  visit(value);
  return out;
}

function clampColor(v) {
  return Math.max(0, Math.min(1, Number(v) || 0));
}

function isDark(color) {
  return luminance(color) < 0.08;
}

function isNearWhite(color) {
  return color.every((v) => v > 0.92);
}

function luminance([r, g, b]) {
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function extractText(args) {
  const strings = [];
  const visit = (value) => {
    if (!value) return;
    if (typeof value === 'string') strings.push(value);
    else if (Array.isArray(value)) value.forEach(visit);
    else if (typeof value.unicode === 'string') strings.push(value.unicode);
  };
  visit(args);
  return strings.join('');
}

function hasMeaningfulText(text) {
  return /[A-Za-z0-9]/.test(text || '');
}

function normalizeText(text) {
  return (text || '').replace(/\s+/g, '').toLowerCase();
}

function textMatches(a, b) {
  const aa = normalizeText(a);
  const bb = normalizeText(b);
  return aa === bb || aa.includes(bb) || bb.includes(aa);
}

function intersects(a, b, minCoverage = 0) {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.width, b.x + b.width);
  const y2 = Math.min(a.y + a.height, b.y + b.height);
  if (x2 <= x1 || y2 <= y1) return false;
  if (minCoverage <= 0) return true;
  const overlap = (x2 - x1) * (y2 - y1);
  const area = Math.max(1, a.width * a.height);
  return overlap / area >= minCoverage;
}

function hasLaterVisibleText(textItems, coveredItems, rect, afterSeq) {
  return coveredItems.every((covered) =>
    textItems.some(
      (item) =>
        item.seq > afterSeq &&
        item.renderingMode !== 3 &&
        !isNearWhite(item.color) &&
        textMatches(item.text, covered.text) &&
        intersects(item.bbox, rect, 0.15)
    )
  );
}

function isPageCounterArtifact(covered, rect, pageBounds) {
  if (!pageBounds) return false;
  const text = summarizeText(covered.map((item) => item.text).join(' '));
  if (!PAGE_COUNTER_RE.test(text)) return false;
  return (
    isNearPageEdge(rect, pageBounds) ||
    covered.some((item) => isNearPageEdge(item.bbox, pageBounds))
  );
}

function isNearPageEdge(rect, pageBounds) {
  const left = rect.x - pageBounds.x;
  const bottom = rect.y - pageBounds.y;
  const right = pageBounds.x + pageBounds.width - (rect.x + rect.width);
  const top = pageBounds.y + pageBounds.height - (rect.y + rect.height);
  return Math.min(left, bottom, right, top) <= PAGE_EDGE_GUARD;
}

function coveredTextBBox(items, rect) {
  const intersections = items.map((item) => intersectionBBox(item.bbox, rect)).filter(Boolean);
  return unionBBoxes(intersections) || unionBBoxes(items.map((item) => item.bbox));
}

function intersectionBBox(a, b) {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.width, b.x + b.width);
  const y2 = Math.min(a.y + a.height, b.y + b.height);
  if (x2 <= x1 || y2 <= y1) return null;
  return { x: x1, y: y1, width: x2 - x1, height: y2 - y1 };
}

function extractCoveredText(items, rect) {
  return [...items]
    .sort((a, b) => b.bbox.y - a.bbox.y || a.bbox.x - b.bbox.x)
    .map((item) => coveredWords(item, rect))
    .filter(Boolean)
    .join(' ');
}

function coveredWords(item, rect) {
  const text = item.text || '';
  if (!text.trim() || item.bbox.width <= 0) return text;

  const x1 = Math.max(item.bbox.x, rect.x);
  const x2 = Math.min(item.bbox.x + item.bbox.width, rect.x + rect.width);
  if (x2 <= x1) return '';

  const startRatio = clampUnit((x1 - item.bbox.x) / item.bbox.width);
  const endRatio = clampUnit((x2 - item.bbox.x) / item.bbox.width);
  if (endRatio - startRatio > 0.85) return text.trim();

  const words = [...text.matchAll(/\S+/g)];
  const selected = words
    .filter((match) => {
      const center = (match.index + match[0].length / 2) / text.length;
      return center >= startRatio && center <= endRatio;
    })
    .map((match) => match[0]);

  if (selected.length) return selected.join(' ');

  const start = Math.floor(startRatio * text.length);
  const end = Math.ceil(endRatio * text.length);
  return text.slice(start, end).trim();
}

function clampUnit(v) {
  return Math.max(0, Math.min(1, Number(v) || 0));
}

function unionBBoxes(boxes) {
  if (boxes.length === 0) return null;
  const minX = Math.min(...boxes.map((b) => b.x));
  const minY = Math.min(...boxes.map((b) => b.y));
  const maxX = Math.max(...boxes.map((b) => b.x + b.width));
  const maxY = Math.max(...boxes.map((b) => b.y + b.height));
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}

function summarizeText(text) {
  const clean = (text || '').replace(/\s+/g, ' ').trim();
  if (clean.length <= 80) return clean;
  return `${clean.slice(0, 77)}...`;
}

function makeFindingId(kind, pageIndex, bbox, text) {
  const rounded = [bbox.x, bbox.y, bbox.width, bbox.height]
    .map((n) => Math.round(n * 10) / 10)
    .join(',');
  return `${kind}:p${pageIndex + 1}:${rounded}:${hashText(text)}`;
}

function normalizeAnnotationId(value) {
  if (!value) return null;
  const match = String(value).match(/(\d+)\s*(?:\d+\s*)?R/i);
  return match ? `${match[1]}R` : String(value);
}

function hashText(text) {
  let hash = 0;
  for (let i = 0; i < text.length; i++) {
    hash = (hash * 31 + text.charCodeAt(i)) | 0;
  }
  return Math.abs(hash).toString(36);
}

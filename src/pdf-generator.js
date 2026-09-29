import { PDFDocument } from "pdf-lib";
import html2canvas from "html2canvas";
import {
  PAGE,
  LOGO,
  DEFAULT_SIZE,
  NUMBERED_LIST,
  buildDocumentModel,
  attachmentLayout,
} from "./document-model.js";

// Lays out the same document model the .docx is built from, the way Word
// lays it out, so the PDF looks like the .docx exported to PDF. All DOM
// measurements are CSS px (96 dpi).

const twipsToPx = (twips) => (twips * 96) / 1440;
const halfPointsToPx = (halfPoints) => (halfPoints / 2) * (96 / 72);
const pointsToPx = (points) => (points * 96) / 72;

const RENDER_SCALE = 2;
const PAGE_W = twipsToPx(PAGE.width);
const PAGE_H = twipsToPx(PAGE.height);
const BODY_LEFT = twipsToPx(PAGE.marginLeft);
const COLUMN_W = PAGE_W - BODY_LEFT - twipsToPx(PAGE.marginRight);
const BODY_BOTTOM = PAGE_H - twipsToPx(PAGE.marginBottom);

// An empty paragraph (and the header's one empty paragraph) takes its
// height from Word's fallback 10pt default font.
const EMPTY_LINE = pointsToPx(11.5);
const MARKER_SIZE = pointsToPx(10);
const MARKER_FONT = "'Times New Roman', 'Liberation Serif', serif";

// Word starts the body below the header when the header runs past the top
// margin: header distance plus that empty header paragraph.
const BODY_TOP = Math.max(twipsToPx(PAGE.marginTop), twipsToPx(PAGE.marginHeader) + EMPTY_LINE);

const LOGO_TOP = LOGO.topOffsetEmu / 9525;

const FONT_FAMILY = "TH Sarabun New";
// Single line spacing = (winAscent + winDescent) / unitsPerEm, as in Word.
const LINE_HEIGHT_FACTOR = { regular: (2619 + 1359) / 2048, bold: (2644 + 1472) / 2048 };

const FONT_FILES = [
  { file: "THSarabunNew.woff", weight: "400" },
  { file: "THSarabunNew-Bold.woff", weight: "700" },
];

let fontBytesPromise;
function fetchFontBytes() {
  if (!fontBytesPromise) {
    const base = import.meta.env.BASE_URL;
    fontBytesPromise = Promise.all(
      FONT_FILES.map(async ({ file, weight }) => {
        const res = await fetch(`${base}fonts/${file}`);
        if (!res.ok) throw new Error(`Failed to load font ${file}`);
        return { weight, bytes: await res.arrayBuffer() };
      }),
    ).catch((err) => {
      fontBytesPromise = null;
      throw err;
    });
  }
  return fontBytesPromise;
}

// Registers the font in a document. html2canvas renders from a cloned copy
// of the page, which doesn't inherit fonts added through the FontFace API,
// so this runs on both the real document (used to measure page breaks) and
// the clone; otherwise text is laid out in a fallback font but drawn in TH
// Sarabun, and the letters overlap.
const installedFonts = new WeakMap();
function installFonts(doc) {
  if (!installedFonts.has(doc)) {
    const promise = fetchFontBytes().then((fonts) => {
      const FontFaceCtor = doc.defaultView.FontFace;
      return Promise.all(
        fonts.map(async ({ weight, bytes }) => {
          const face = new FontFaceCtor(FONT_FAMILY, bytes.slice(0), { weight });
          await face.load();
          doc.fonts.add(face);
        }),
      );
    });
    promise.catch(() => installedFonts.delete(doc));
    installedFonts.set(doc, promise);
  }
  return installedFonts.get(doc);
}

function el(tag, style, children) {
  const node = document.createElement(tag);
  if (style) Object.assign(node.style, style);
  for (const child of children ?? []) {
    if (child == null) continue;
    node.appendChild(typeof child === "string" ? document.createTextNode(child) : child);
  }
  return node;
}

function renderParagraph(p, marker) {
  const hasText = p.runs.some((r) => r.text);
  const size = Math.max(...p.runs.map((r) => r.size ?? DEFAULT_SIZE));
  const bold = p.runs.some((r) => r.bold && r.text);
  const lineHeight = hasText
    ? halfPointsToPx(size) * (bold ? LINE_HEIGHT_FACTOR.bold : LINE_HEIGHT_FACTOR.regular)
    : EMPTY_LINE;
  const hanging = p.list?.type === "number" ? NUMBERED_LIST.hanging : (p.indent?.hanging ?? 0);
  const spaceBefore = twipsToPx(p.spacing?.before ?? 0);
  const spaceAfter = twipsToPx(p.spacing?.after ?? 0);

  const children = [];
  if (marker) {
    children.push(
      el("span", {
        display: "inline-block",
        width: `${twipsToPx(hanging)}px`,
        textIndent: "0",
        fontFamily: MARKER_FONT,
        fontSize: `${MARKER_SIZE}px`,
        fontWeight: "400",
      }, [marker]),
    );
  }
  if (hasText) {
    for (const r of p.runs) {
      children.push(
        el("span", {
          fontSize: `${halfPointsToPx(r.size ?? DEFAULT_SIZE)}px`,
          fontWeight: r.bold ? "700" : "400",
        }, [r.text ?? ""]),
      );
    }
  } else {
    children.push(el("div", { height: `${EMPTY_LINE}px` }));
  }

  // Margins (not padding) so adjacent spacing collapses to the larger of
  // space-after and space-before, which is how Word combines them.
  const node = el("div", {
    fontFamily: `'${FONT_FAMILY}', sans-serif`,
    fontSize: `${halfPointsToPx(size)}px`,
    lineHeight: `${lineHeight}px`,
    whiteSpace: "pre-wrap",
    color: "#000",
    textAlign: p.align ?? "left",
    marginTop: `${spaceBefore}px`,
    marginBottom: `${spaceAfter}px`,
    paddingLeft: `${twipsToPx(p.indent?.left ?? 0)}px`,
    textIndent: `${-twipsToPx(hanging)}px`,
  }, children);

  return { node, lineHeight };
}

function buildColumn(data) {
  const listCounters = new Map();
  const paragraphs = buildDocumentModel(data).map((p) => {
    let marker = null;
    if (p.list?.type === "bullet") marker = "●";
    if (p.list?.type === "number") {
      const n = (listCounters.get(p.list.group) ?? 0) + 1;
      listCounters.set(p.list.group, n);
      marker = `${n}.`;
    }
    return renderParagraph(p, marker);
  });
  // flow-root keeps the title's space-before inside the column.
  const column = el("div", {
    width: `${COLUMN_W}px`,
    background: "#ffffff",
    display: "flow-root",
  }, paragraphs.map((p) => p.node));
  return { column, paragraphs };
}

// Every line's vertical extent within the column (excluding paragraph
// spacing), derived from each paragraph's fixed line height.
function measureLines(column, paragraphs) {
  const columnTop = column.getBoundingClientRect().top;
  const lines = [];
  for (const { node, lineHeight } of paragraphs) {
    const rect = node.getBoundingClientRect();
    const top = rect.top - columnTop;
    const count = Math.max(1, Math.round(rect.height / lineHeight));
    for (let i = 0; i < count; i++) {
      lines.push({ top: top + i * lineHeight, bottom: top + (i + 1) * lineHeight });
    }
  }
  return lines;
}

// Breaks pages between lines, never through one. The first page keeps the
// title's space-before; later pages drop the space-before of the paragraph
// that starts them, as Word does after a natural page break.
function paginate(lines) {
  const available = BODY_BOTTOM - BODY_TOP;
  const pages = [];
  let start = 0;
  let end = 0;
  for (const line of lines) {
    if (line.bottom - start > available && end > start) {
      pages.push({ start, end });
      start = line.top;
    }
    end = line.bottom;
  }
  pages.push({ start, end });
  return pages;
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error(`Failed to load image: ${url}`));
    img.src = url;
  });
}

function canvasToPngBytes(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Failed to encode PDF page image"));
        return;
      }
      blob.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)));
    }, "image/png");
  });
}

function newPageCanvas() {
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(PAGE_W * RENDER_SCALE);
  canvas.height = Math.round(PAGE_H * RENDER_SCALE);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  return { canvas, ctx };
}

// The logo banner the docx header puts on every page. Drawn last: it floats
// in front of the body, whose white background would otherwise cover it.
function drawLogo(ctx, logoImg) {
  ctx.drawImage(logoImg, 0, LOGO_TOP * RENDER_SCALE, ctx.canvas.width, LOGO.heightPx * RENDER_SCALE);
}

async function renderMainContentPages(data, logoImg) {
  // If the font can't load (e.g. offline), neither copy gets it, so layout
  // and drawing still agree on the fallback font.
  const fontReady = await installFonts(document).then(
    () => true,
    (err) => {
      console.warn("TH Sarabun font failed to load; PDF will use a fallback font.", err);
      return false;
    },
  );
  const { column, paragraphs } = buildColumn(data);
  const host = el("div", { position: "fixed", top: "0", left: "-10000px", zIndex: "-1" }, [column]);
  document.body.appendChild(host);

  let contentCanvas;
  let pages;
  try {
    pages = paginate(measureLines(column, paragraphs));
    contentCanvas = await html2canvas(column, {
      scale: RENDER_SCALE,
      useCORS: true,
      backgroundColor: "#ffffff",
      windowWidth: Math.ceil(PAGE_W),
      onclone: (clonedDoc) => (fontReady ? installFonts(clonedDoc) : undefined),
    });
  } finally {
    document.body.removeChild(host);
  }

  const result = [];
  for (const { start, end } of pages) {
    const { canvas, ctx } = newPageCanvas();
    const height = (end - start) * RENDER_SCALE;
    ctx.drawImage(
      contentCanvas,
      0, start * RENDER_SCALE, contentCanvas.width, height,
      BODY_LEFT * RENDER_SCALE, BODY_TOP * RENDER_SCALE, contentCanvas.width, height,
    );
    drawLogo(ctx, logoImg);
    result.push(await canvasToPngBytes(canvas));
  }
  return result;
}

// Mirrors the docx attachment page: image centred in the column at the top
// of the body.
function renderAttachmentPage(att, logoImg) {
  const { canvas, ctx } = newPageCanvas();
  const { width, height } = attachmentLayout(att);
  const x = BODY_LEFT + (COLUMN_W - width) / 2;
  const y = BODY_TOP;

  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = async () => {
      ctx.drawImage(img, x * RENDER_SCALE, y * RENDER_SCALE, width * RENDER_SCALE, height * RENDER_SCALE);
      drawLogo(ctx, logoImg);
      resolve(await canvasToPngBytes(canvas));
    };
    img.onerror = () => reject(new Error(`Failed to load attachment image: ${att.label}`));
    img.src = URL.createObjectURL(new Blob([att.bytes], { type: "image/png" }));
  });
}

export async function generateReferralFormPdf(data) {
  const pdfDoc = await PDFDocument.create();
  const logoImg = await loadImage(`${import.meta.env.BASE_URL}logo.png`);

  const mainPageBytes = await renderMainContentPages(data, logoImg);
  const attachmentPageBytes = await Promise.all(
    (data.attachments ?? []).map((att) => renderAttachmentPage(att, logoImg)),
  );

  const pageWidthPt = PAGE.width / 20;
  const pageHeightPt = PAGE.height / 20;
  for (const bytes of [...mainPageBytes, ...attachmentPageBytes]) {
    const page = pdfDoc.addPage([pageWidthPt, pageHeightPt]);
    const png = await pdfDoc.embedPng(bytes);
    page.drawImage(png, { x: 0, y: 0, width: pageWidthPt, height: pageHeightPt });
  }

  return pdfDoc.save();
}

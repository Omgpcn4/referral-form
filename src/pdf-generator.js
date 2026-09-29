import {
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFOperator,
  PDFOperatorNames,
  StandardFonts,
  beginText,
  endText,
  setFontAndSize,
  setTextMatrix,
  rgb,
} from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import * as hb from "harfbuzzjs";
import { thaiFallbackShape, thaiForPua } from "./thai-shaping.js";
import {
  PAGE,
  LOGO,
  DEFAULT_SIZE,
  NUMBERED_LIST,
  buildDocumentModel,
  attachmentLayout,
} from "./document-model.js";

// Builds a PDF with real, selectable text laid out the way Word lays out the
// .docx. The browser lays the letter out (line breaks, word positions) in TH
// Sarabun; HarfBuzz, the shaping engine browsers use, picks each word's
// glyphs, including the font's substitutions that place Thai tone marks and
// vowels correctly; pdf-lib writes those glyphs with the font embedded.
// DOM measurements are CSS px (96 dpi); PDF coordinates are points.

const twipsToPx = (twips) => (twips * 96) / 1440;
const halfPointsToPx = (halfPoints) => (halfPoints / 2) * (96 / 72);
const pointsToPx = (points) => (points * 96) / 72;
const pxToPt = (px) => px * 0.75;

const PAGE_W = twipsToPx(PAGE.width);
const PAGE_H = twipsToPx(PAGE.height);
const PAGE_W_PT = PAGE.width / 20;
const PAGE_H_PT = PAGE.height / 20;
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
const FONT_FILES = {
  400: { file: "THSarabunNew.ttf", ascent: 2619, descent: 1359 },
  700: { file: "THSarabunNew-Bold.ttf", ascent: 2644, descent: 1472 },
};
const UNITS_PER_EM = 2048;
// Single line spacing = (winAscent + winDescent) / unitsPerEm, as in Word.
const lineHeightFactor = (weight) =>
  (FONT_FILES[weight].ascent + FONT_FILES[weight].descent) / UNITS_PER_EM;
const ascentFactor = (weight) => FONT_FILES[weight].ascent / UNITS_PER_EM;

// Kerning and ligatures are off, as in Word; the layout and the PDF glyphs
// must use the same settings so they agree on every advance.
const SHAPING_FEATURES = ["kern", "liga", "clig", "calt"].map((tag) => new hb.Feature(tag, 0));

// ---------------------------------------------------------------- fonts --

let fontsPromise;
function loadFonts() {
  if (!fontsPromise) {
    const base = import.meta.env.BASE_URL;
    fontsPromise = Promise.all(
      Object.entries(FONT_FILES).map(async ([weight, { file }]) => {
        const res = await fetch(`${base}fonts/${file}`);
        if (!res.ok) throw new Error(`Failed to load font ${file}`);
        const bytes = await res.arrayBuffer();
        const face = new FontFace(FONT_FAMILY, bytes.slice(0), { weight });
        await face.load();
        document.fonts.add(face);
        const hbFont = new hb.Font(new hb.Face(new hb.Blob(bytes.slice(0))));
        return [weight, { bytes, hbFont }];
      }),
    )
      .then((entries) => Object.fromEntries(entries))
      .catch((err) => {
        fontsPromise = null;
        throw err;
      });
  }
  return fontsPromise;
}

function shape(hbFont, text) {
  const buffer = new hb.Buffer();
  buffer.addText(thaiFallbackShape(text, (u) => hbFont.nominalGlyph(u) !== undefined));
  buffer.guessSegmentProperties();
  hb.shape(hbFont, buffer, SHAPING_FEATURES);
  const infos = buffer.getGlyphInfos();
  const positions = buffer.getGlyphPositions();
  const glyphs = infos.map((info, i) => ({ gid: info.codepoint, ...positions[i] }));
  return text.includes("ำ") ? recomposeSaraAm(hbFont, glyphs) : glyphs;
}

// Shaping splits SARA AM into NIKHAHIT + SARA AA so tone marks can sit
// between them. In TH Sarabun the SARA AM glyph is exactly those two
// outlines with the same advance, so drawing it instead looks identical and
// lets the PDF's text read back as "ำ" rather than "ํา".
function recomposeSaraAm(hbFont, glyphs) {
  const nikhahit = hbFont.nominalGlyph(0x0e4d);
  const saraAa = hbFont.nominalGlyph(0x0e32);
  const saraAm = hbFont.nominalGlyph(0x0e33);
  const out = [];
  for (let i = 0; i < glyphs.length; i++) {
    if (glyphs[i].gid === nikhahit) {
      let j = i + 1;
      while (j < glyphs.length && glyphs[j].xAdvance === 0 && glyphs[j].gid !== saraAa) j++;
      if (j < glyphs.length && glyphs[j].gid === saraAa) {
        out.push(...glyphs.slice(i + 1, j), { ...glyphs[j], gid: saraAm });
        i = j;
        continue;
      }
    }
    out.push(glyphs[i]);
  }
  return out;
}

// pdf-lib builds the font's ToUnicode table from each glyph's code points.
// The Thai variant glyphs are mapped from Private Use Area code points, so
// point them at the Thai character they stand for.
const thaiAwareFontkit = {
  create(data) {
    const font = fontkit.create(data);
    const getGlyph = font.getGlyph.bind(font);
    font.getGlyph = (id, codePoints) => {
      const glyph = getGlyph(id, codePoints);
      const [first] = glyph.codePoints;
      const thai = first !== undefined && thaiForPua(first);
      if (thai) glyph.codePoints = thai;
      return glyph;
    };
    return font;
  },
};

// --------------------------------------------------------------- layout --

function el(tag, style, children, attrs) {
  const node = document.createElement(tag);
  if (style) Object.assign(node.style, style);
  if (attrs) Object.assign(node.dataset, attrs);
  for (const child of children ?? []) {
    if (child == null) continue;
    node.appendChild(typeof child === "string" ? document.createTextNode(child) : child);
  }
  return node;
}

function renderParagraph(p, marker) {
  const hasText = p.runs.some((r) => r.text);
  const size = halfPointsToPx(Math.max(...p.runs.map((r) => r.size ?? DEFAULT_SIZE)));
  const weight = p.runs.some((r) => r.bold && r.text) ? 700 : 400;
  const lineHeight = hasText ? size * lineHeightFactor(weight) : EMPTY_LINE;
  const hanging = p.list?.type === "number" ? NUMBERED_LIST.hanging : (p.indent?.hanging ?? 0);

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
      }, [marker], { role: "marker", size: String(MARKER_SIZE) }),
    );
  }
  if (hasText) {
    for (const r of p.runs) {
      const runWeight = r.bold ? 700 : 400;
      children.push(
        el("span", {
          fontSize: `${halfPointsToPx(r.size ?? DEFAULT_SIZE)}px`,
          fontWeight: String(runWeight),
        }, [r.text ?? ""], {
          role: "text",
          size: String(halfPointsToPx(r.size ?? DEFAULT_SIZE)),
          weight: String(runWeight),
        }),
      );
    }
  } else {
    children.push(el("div", { height: `${EMPTY_LINE}px` }));
  }

  // Margins (not padding) so adjacent spacing collapses to the larger of
  // space-after and space-before, which is how Word combines them.
  const node = el("div", {
    fontFamily: `'${FONT_FAMILY}', sans-serif`,
    fontSize: `${size}px`,
    lineHeight: `${lineHeight}px`,
    fontKerning: "none",
    fontVariantLigatures: "none",
    whiteSpace: "pre-wrap",
    textAlign: p.align ?? "left",
    marginTop: `${twipsToPx(p.spacing?.before ?? 0)}px`,
    marginBottom: `${twipsToPx(p.spacing?.after ?? 0)}px`,
    paddingLeft: `${twipsToPx(p.indent?.left ?? 0)}px`,
    textIndent: `${-twipsToPx(hanging)}px`,
  }, children);

  return { node, lineHeight, baseline: hasText ? size * ascentFactor(weight) : 0 };
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
    display: "flow-root",
    color: "#000",
  }, paragraphs.map((p) => p.node));
  return { column, paragraphs };
}

// Every line's vertical extent within the column (excluding paragraph
// spacing), derived from each paragraph's fixed line height.
function measureLines(columnTop, paragraphs) {
  const lines = [];
  for (const p of paragraphs) {
    const rect = p.node.getBoundingClientRect();
    p.top = rect.top - columnTop;
    const count = Math.max(1, Math.round(rect.height / p.lineHeight));
    for (let i = 0; i < count; i++) {
      lines.push({ top: p.top + i * p.lineHeight, bottom: p.top + (i + 1) * p.lineHeight });
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

const graphemes = new Intl.Segmenter("th", { granularity: "grapheme" });
const SPACE = /^\s+$/;

// Splits the laid-out text into fragments: runs of non-space characters on
// one line within one span, each with its position and the line's baseline.
function collectFragments(columnLeft, paragraphs) {
  const fragments = [];
  const range = document.createRange();
  for (const p of paragraphs) {
    const paragraphTop = p.node.getBoundingClientRect().top;
    for (const span of p.node.querySelectorAll("span[data-role]")) {
      const text = span.firstChild;
      if (!text) continue;
      let current = null;
      for (const { segment, index } of graphemes.segment(text.data)) {
        if (SPACE.test(segment)) {
          if (current) current.trailing += segment;
          current = null;
          continue;
        }
        range.setStart(text, index);
        range.setEnd(text, index + segment.length);
        const rect = range.getBoundingClientRect();
        const middle = rect.top + rect.height / 2 - paragraphTop;
        const line = Math.max(0, Math.floor(middle / p.lineHeight));
        if (!current || current.line !== line) {
          current = {
            text: "",
            trailing: "",
            line,
            lineTop: p.top + line * p.lineHeight,
            lineBottom: p.top + (line + 1) * p.lineHeight,
            baseline: p.top + line * p.lineHeight + p.baseline,
            left: rect.left - columnLeft,
            width: 0,
            role: span.dataset.role,
            size: Number(span.dataset.size),
            weight: Number(span.dataset.weight ?? 400),
            fontFamily: getComputedStyle(span).fontFamily,
          };
          fragments.push(current);
        }
        current.text += segment;
        current.width = rect.right - columnLeft - current.left;
      }
    }
  }
  return fragments;
}

// ------------------------------------------------------------- drawing --

const BLACK = rgb(0, 0, 0);

function pdfX(px) {
  return pxToPt(BODY_LEFT + px);
}
function pdfY(page, px) {
  return PAGE_H_PT - pxToPt(BODY_TOP + px - page.start);
}

// Real text: HarfBuzz glyphs written as-is (Identity-H, so each code is a
// glyph id), wrapped in ActualText so copying gives the original characters
// rather than the font's alternate glyphs.
function drawShapedText(pdfPage, context, fontKey, hbFont, fragment, glyphs, x, y) {
  const sizePt = pxToPt(fragment.size);
  const scale = sizePt / UNITS_PER_EM;
  const ops = [
    PDFOperator.of(PDFOperatorNames.BeginMarkedContentSequence, [
      PDFName.of("Span"),
      context.obj({ ActualText: PDFHexString.fromText(fragment.text + fragment.trailing) }),
    ]),
    beginText(),
    setFontAndSize(fontKey, sizePt),
  ];
  const hex = (g) => PDFHexString.of(g.gid.toString(16).padStart(4, "0"));
  if (glyphs.every((g) => g.xOffset === 0 && g.yOffset === 0)) {
    // One run of glyphs; TJ adjustments cover any advance that differs from
    // the font's own width for the glyph.
    const items = [];
    for (const g of glyphs) {
      items.push(hex(g));
      const adjust = ((hbFont.glyphHAdvance(g.gid) - g.xAdvance) * 1000) / UNITS_PER_EM;
      if (Math.abs(adjust) > 1e-6) items.push(adjust);
    }
    ops.push(
      setTextMatrix(1, 0, 0, 1, x, y),
      PDFOperator.of(PDFOperatorNames.ShowTextAdjusted, [context.obj(items)]),
    );
  } else {
    let penX = x;
    for (const g of glyphs) {
      ops.push(
        setTextMatrix(1, 0, 0, 1, penX + g.xOffset * scale, y + g.yOffset * scale),
        PDFOperator.of(PDFOperatorNames.ShowText, [hex(g)]),
      );
      penX += g.xAdvance * scale;
    }
  }
  ops.push(endText(), PDFOperator.of(PDFOperatorNames.EndMarkedContent));
  pdfPage.pushOperators(...ops);
}

// ● and ○ aren't in TH Sarabun; drawn as circles, sized like a typical
// symbol font's glyph.
const CIRCLES = { "●": true, "○": false };

function drawCircleGlyph(pdfPage, fragment, filled, x, y) {
  const sizePt = pxToPt(fragment.size);
  const radius = sizePt * 0.34;
  const cx = x + pxToPt(fragment.width) / 2;
  const cy = y + sizePt * 0.33;
  pdfPage.drawCircle({
    x: cx,
    y: cy,
    size: filled ? radius : radius - sizePt * 0.03,
    color: filled ? BLACK : undefined,
    borderColor: filled ? undefined : BLACK,
    borderWidth: filled ? 0 : sizePt * 0.06,
  });
}

// Anything the font can't draw (emoji, other scripts) is drawn by the
// browser at high resolution and placed as a small image.
async function drawRasterFallback(pdfDoc, pdfPage, fragment, page) {
  const scale = 4;
  const height = fragment.lineBottom - fragment.lineTop;
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.ceil(fragment.width * scale));
  canvas.height = Math.ceil(height * scale);
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  // Built by hand: the computed `font` shorthand is empty when ligatures are
  // turned off, and the canvas would fall back to 10px sans-serif.
  ctx.font = `${fragment.weight} ${fragment.size}px ${fragment.fontFamily}`;
  ctx.fillStyle = "#000";
  ctx.textBaseline = "alphabetic";
  ctx.fillText(fragment.text, 0, fragment.baseline - fragment.lineTop);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  const image = await pdfDoc.embedPng(await blob.arrayBuffer());
  pdfPage.drawImage(image, {
    x: pdfX(fragment.left),
    y: pdfY(page, fragment.lineBottom),
    width: pxToPt(fragment.width),
    height: pxToPt(height),
  });
}

async function drawFragment(pdfDoc, pdfPage, page, fragment, fonts, fontKeys, markerFont) {
  const x = pdfX(fragment.left);
  const y = pdfY(page, fragment.baseline);

  if (fragment.text in CIRCLES) {
    drawCircleGlyph(pdfPage, fragment, CIRCLES[fragment.text], x, y);
    return;
  }
  if (fragment.role === "marker") {
    pdfPage.drawText(fragment.text, { x, y, size: pxToPt(fragment.size), font: markerFont, color: BLACK });
    return;
  }
  const { hbFont } = fonts[fragment.weight];
  const glyphs = shape(hbFont, fragment.text);
  if (glyphs.some((g) => g.gid === 0)) {
    await drawRasterFallback(pdfDoc, pdfPage, fragment, page);
    return;
  }
  drawShapedText(pdfPage, pdfDoc.context, fontKeys[fragment.weight], hbFont, fragment, glyphs, x, y);
}

async function fetchPng(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to load image: ${url}`);
  return res.arrayBuffer();
}

function addPage(pdfDoc, logo) {
  const pdfPage = pdfDoc.addPage([PAGE_W_PT, PAGE_H_PT]);
  return { pdfPage, drawLogo: () => drawLogo(pdfPage, logo) };
}

// The logo banner the docx header puts on every page, full page width.
function drawLogo(pdfPage, logo) {
  const height = pxToPt(LOGO.heightPx);
  pdfPage.drawImage(logo, {
    x: 0,
    y: PAGE_H_PT - pxToPt(LOGO_TOP) - height,
    width: PAGE_W_PT,
    height,
  });
}

export async function generateReferralFormPdf(data) {
  const fonts = await loadFonts();
  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(thaiAwareFontkit);
  pdfDoc.setTitle(`Referral letter ${data.petName ?? ""}`.trim());
  const pdfFonts = {
    400: await pdfDoc.embedFont(fonts[400].bytes, { subset: false }),
    700: await pdfDoc.embedFont(fonts[700].bytes, { subset: false }),
  };
  const markerFont = await pdfDoc.embedFont(StandardFonts.TimesRoman);
  const logo = await pdfDoc.embedPng(await fetchPng(`${import.meta.env.BASE_URL}logo.png`));

  const { column, paragraphs } = buildColumn(data);
  const host = el("div", { position: "fixed", top: "0", left: "-10000px", zIndex: "-1" }, [column]);
  document.body.appendChild(host);
  let pages;
  let fragments;
  try {
    const columnRect = column.getBoundingClientRect();
    pages = paginate(measureLines(columnRect.top, paragraphs));
    fragments = collectFragments(columnRect.left, paragraphs);
  } finally {
    document.body.removeChild(host);
  }

  for (const page of pages) {
    const { pdfPage, drawLogo: finishPage } = addPage(pdfDoc, logo);
    const fontKeys = {
      400: pdfPage.node.newFontDictionary(pdfFonts[400].name, pdfFonts[400].ref),
      700: pdfPage.node.newFontDictionary(pdfFonts[700].name, pdfFonts[700].ref),
    };
    const onPage = fragments.filter((f) => f.lineTop >= page.start - 0.5 && f.lineBottom <= page.end + 0.5);
    for (const fragment of onPage) {
      await drawFragment(pdfDoc, pdfPage, page, fragment, fonts, fontKeys, markerFont);
    }
    finishPage();
  }

  // Attachment pages mirror the docx: image centred in the column at the
  // top of the body, embedded at full resolution.
  for (const att of data.attachments ?? []) {
    const { pdfPage, drawLogo: finishPage } = addPage(pdfDoc, logo);
    const image = await pdfDoc.embedPng(att.bytes);
    const { width, height } = attachmentLayout(att);
    pdfPage.drawImage(image, {
      x: pxToPt(BODY_LEFT + (COLUMN_W - width) / 2),
      y: PAGE_H_PT - pxToPt(BODY_TOP + height),
      width: pxToPt(width),
      height: pxToPt(height),
    });
    finishPage();
  }

  return pdfDoc.save();
}

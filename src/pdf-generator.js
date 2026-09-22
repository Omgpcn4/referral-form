import { PDFDocument } from "pdf-lib";
import html2canvas from "html2canvas";

// A4 at 96dpi (CSS px) for the on-screen render, and in PDF points for the
// final pages.
const PAGE_WIDTH_PX = 794;
const PAGE_HEIGHT_PX = 1123;
const PAGE_WIDTH_PT = 595.28;
const PAGE_HEIGHT_PT = 841.89;
const RENDER_SCALE = 2;

// Reserve a header band (for the logo) on every page — matching the docx,
// where the header repeats automatically. FOOTER_BAND_PX is just a bottom
// margin now (no footer content, but keeps content off the page edge).
// Logo aspect ratio matches src/assets/logo.png (2639x270).
const HEADER_BAND_PX = 96;
const FOOTER_BAND_PX = 28;
const LOGO_HEIGHT_PX = Math.round(PAGE_WIDTH_PX * (270 / 2639));

const UNSELECTED = "○";
const SELECTED = "●";

function s(value) {
  return value == null ? "" : String(value).trim();
}

function formatDate(isoDate) {
  if (!isoDate) return "";
  const [year, month, day] = isoDate.split("-");
  if (!year || !month || !day) return isoDate;
  return `${day}/${month}/${year}`;
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

const BASE_TEXT = { fontFamily: "'Sarabun', sans-serif", fontSize: "15px", color: "#111" };

function labeledLine(fields) {
  const parts = [];
  fields.forEach(([label, value], i) => {
    if (i > 0) parts.push("     ");
    parts.push(`${label}: ${s(value)}`);
  });
  return el("div", { ...BASE_TEXT, marginBottom: "8px" }, parts);
}

function dateHnLine(label, value) {
  return el("div", { ...BASE_TEXT, fontSize: "13px", textAlign: "right" }, [`${label}: ${s(value)}`]);
}

function blockLineStyle(block) {
  return { ...BASE_TEXT, fontSize: "13px", fontWeight: block.bold ? "700" : "400" };
}

function renderBlocks(blocks) {
  const nodes = [];
  let i = 0;
  while (i < blocks.length) {
    const block = blocks[i];
    if (block.type === "bullet" || block.type === "number") {
      const groupType = block.type;
      const items = [];
      while (i < blocks.length && blocks[i].type === groupType) {
        items.push(el("li", blockLineStyle(blocks[i]), [blocks[i].text]));
        i++;
      }
      nodes.push(
        el(groupType === "bullet" ? "ul" : "ol", {
          marginLeft: "18px",
          marginTop: "0",
          marginBottom: "0",
        }, items),
      );
    } else {
      nodes.push(el("div", { ...blockLineStyle(block), marginLeft: "18px" }, [block.text]));
      i++;
    }
  }
  return nodes;
}

function numberedSection(number, label, blocks) {
  return el("div", { marginBottom: "6px" }, [
    el("div", { ...BASE_TEXT, marginLeft: "18px", marginBottom: "2px" }, [`${number}. ${label}`]),
    ...(blocks && blocks.length ? renderBlocks(blocks) : [el("div", { height: "14px" })]),
  ]);
}

function purposeLine(purpose) {
  const options = [
    { key: "diagnosis", th: "รับการวินิจฉัย", en: "Diagnosis" },
    { key: "treatment", th: "รับการรักษา", en: "Treatment" },
    { key: "ownerRequest", th: "ตามความต้องการของเจ้าของ", en: "Owner request" },
  ];
  return el("div", { ...BASE_TEXT, marginBottom: "10px" }, [
    "เพื่อ For :  ",
    ...options.map((opt, i) => {
      const mark = purpose === opt.key ? SELECTED : UNSELECTED;
      return `${i > 0 ? "     " : ""}${mark} ${opt.th} ${opt.en}`;
    }),
  ]);
}

function buildPrintableDom(data) {
  const content = el("div", {
    width: `${PAGE_WIDTH_PX}px`,
    background: "#ffffff",
    paddingLeft: "48px",
    paddingRight: "48px",
    boxSizing: "border-box",
  }, [
    el("div", { ...BASE_TEXT, fontSize: "21px", fontWeight: "700", textAlign: "center" }, [
      "ใบส่งตัวสัตว์ป่วย",
    ]),
    el("div", { ...BASE_TEXT, fontSize: "21px", fontWeight: "700", textAlign: "center", marginBottom: "10px" }, [
      "Referral Form",
    ]),
    dateHnLine("วันที่ Date", formatDate(data.date)),
    el("div", { marginBottom: "10px" }, [dateHnLine("HN", s(data.hn))]),
    el("div", { ...BASE_TEXT, marginBottom: "10px" }, [
      "เรียน สัตวแพทย์ผู้เกี่ยวข้อง To Whom it may concern",
    ]),
    labeledLine([
      ["ชื่อสัตว์เลี้ยง Pet's name", data.petName],
      ["ชนิด Species", data.species],
      ["เพศ Gender", data.gender],
    ]),
    labeledLine([
      ["พันธุ์ Breed", data.breed],
      ["อายุ Age", data.age],
    ]),
    labeledLine([["ชื่อเจ้าของสัตว์เลี้ยง Owner's name", data.ownerName]]),
    purposeLine(data.purpose),
    numberedSection(1, "ประวัติอาการ History", data.history),
    numberedSection(2, "อาการป่วยปัจจุบัน/ผลการตรวจร่างกาย Physical Examination", data.physicalExam),
    numberedSection(3, "ผลการตรวจทางห้องปฏิบัติการ Laboratory", data.laboratory),
    numberedSection(4, "การวินิจฉัยเบื้องต้น Diagnosis", data.diagnosis),
    numberedSection(5, "การรักษา Treatment", data.treatment),
    numberedSection(6, "รายละเอียดอื่นๆ Others", data.others),
    el("div", { ...BASE_TEXT, textAlign: "center", marginTop: "20px" }, ["เรียนมาเพื่อทราบ Sincerely,"]),
    el("div", { height: "36px" }),
    el("div", { ...BASE_TEXT, fontSize: "13px", textAlign: "center" }, [`(${s(data.vetName)})`]),
    el("div", { ...BASE_TEXT, fontSize: "13px", textAlign: "center", marginBottom: "24px" }, [
      `ใบอนุญาตเลขที่ Veterinary License No. ${s(data.licenseNo)}`,
    ]),
  ]);
  return content;
}

async function waitForFonts() {
  if (!document.fonts) return;
  try {
    await Promise.all([
      document.fonts.load('400 16px "Sarabun"'),
      document.fonts.load('700 16px "Sarabun"'),
    ]);
    await document.fonts.ready;
  } catch (err) {
    // Web font failed to load (offline, blocked, etc.) — proceed with whatever renders.
  }
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

// Draws the logo header band that appears on every page — mirroring the
// docx, where this comes from a real repeating header.
function drawPageHeader(ctx, logoImg) {
  const pageWidthPxScaled = PAGE_WIDTH_PX * RENDER_SCALE;
  const logoHeightScaled = LOGO_HEIGHT_PX * RENDER_SCALE;
  ctx.drawImage(logoImg, 0, 0, pageWidthPxScaled, logoHeightScaled);
}

async function renderMainContentPages(data, logoImg) {
  await waitForFonts();
  const dom = buildPrintableDom(data);
  const host = el("div", {
    position: "fixed",
    top: "0",
    left: "-10000px",
    zIndex: "-1",
  }, [dom]);
  document.body.appendChild(host);

  let contentCanvas;
  try {
    contentCanvas = await html2canvas(dom, {
      scale: RENDER_SCALE,
      useCORS: true,
      backgroundColor: "#ffffff",
      width: PAGE_WIDTH_PX,
      windowWidth: PAGE_WIDTH_PX,
    });
  } finally {
    document.body.removeChild(host);
  }

  const pageWidthPxScaled = PAGE_WIDTH_PX * RENDER_SCALE;
  const pageHeightPxScaled = PAGE_HEIGHT_PX * RENDER_SCALE;
  const headerBandScaled = HEADER_BAND_PX * RENDER_SCALE;
  const footerBandScaled = FOOTER_BAND_PX * RENDER_SCALE;
  const availableContentHeightScaled = pageHeightPxScaled - headerBandScaled - footerBandScaled;
  const pageCount = Math.max(1, Math.ceil(contentCanvas.height / availableContentHeightScaled));

  const pages = [];
  for (let i = 0; i < pageCount; i++) {
    const pageCanvas = document.createElement("canvas");
    pageCanvas.width = pageWidthPxScaled;
    pageCanvas.height = pageHeightPxScaled;
    const ctx = pageCanvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
    ctx.drawImage(
      contentCanvas,
      0,
      i * availableContentHeightScaled,
      pageWidthPxScaled,
      availableContentHeightScaled,
      0,
      headerBandScaled,
      pageWidthPxScaled,
      availableContentHeightScaled,
    );
    drawPageHeader(ctx, logoImg);
    pages.push(await canvasToPngBytes(pageCanvas));
  }
  return pages;
}

function renderAttachmentPage(att, logoImg) {
  const pageWidthPxScaled = PAGE_WIDTH_PX * RENDER_SCALE;
  const pageHeightPxScaled = PAGE_HEIGHT_PX * RENDER_SCALE;
  const headerBandScaled = HEADER_BAND_PX * RENDER_SCALE;
  const footerBandScaled = FOOTER_BAND_PX * RENDER_SCALE;
  const availableHeight = pageHeightPxScaled - headerBandScaled - footerBandScaled;
  const availableWidth = pageWidthPxScaled - 48 * RENDER_SCALE * 2;

  const pageCanvas = document.createElement("canvas");
  pageCanvas.width = pageWidthPxScaled;
  pageCanvas.height = pageHeightPxScaled;
  const ctx = pageCanvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);

  const scale = Math.min(1, availableWidth / att.width, availableHeight / att.height);
  const width = att.width * scale;
  const height = att.height * scale;
  const x = (pageWidthPxScaled - width) / 2;
  const y = headerBandScaled + (availableHeight - height) / 2;

  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = async () => {
      ctx.drawImage(img, x, y, width, height);
      drawPageHeader(ctx, logoImg);
      resolve(await canvasToPngBytes(pageCanvas));
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

  for (const bytes of [...mainPageBytes, ...attachmentPageBytes]) {
    const page = pdfDoc.addPage([PAGE_WIDTH_PT, PAGE_HEIGHT_PT]);
    const png = await pdfDoc.embedPng(bytes);
    page.drawImage(png, { x: 0, y: 0, width: PAGE_WIDTH_PT, height: PAGE_HEIGHT_PT });
  }

  return pdfDoc.save();
}

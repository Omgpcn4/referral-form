import { PDFDocument } from "pdf-lib";
import html2canvas from "html2canvas";

// A4 at 96dpi (CSS px) for the on-screen render, and in PDF points for the
// final pages.
const PAGE_WIDTH_PX = 794;
const PAGE_HEIGHT_PX = 1123;
const PAGE_WIDTH_PT = 595.28;
const PAGE_HEIGHT_PT = 841.89;
const RENDER_SCALE = 2;

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
    parts.push(`${label} : ${s(value)}`);
  });
  return el("div", { ...BASE_TEXT, marginBottom: "8px" }, parts);
}

function dateHnLine(label, value) {
  return el("div", { ...BASE_TEXT, fontSize: "13px", textAlign: "right" }, [`${label} : ${s(value)}`]);
}

function numberedSection(number, label, text) {
  const lines = s(text)
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  return el("div", { marginBottom: "6px" }, [
    el("div", { ...BASE_TEXT, marginLeft: "18px", marginBottom: "2px" }, [`${number}. ${label}`]),
    ...(lines.length
      ? lines.map((line) =>
          el("div", { ...BASE_TEXT, fontSize: "13px", marginLeft: "18px" }, [line]),
        )
      : [el("div", { height: "14px" })]),
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
  const wrapper = el("div", {
    position: "relative",
    width: `${PAGE_WIDTH_PX}px`,
    background: "#ffffff",
    boxSizing: "border-box",
  });

  const logo = document.createElement("img");
  logo.src = `${import.meta.env.BASE_URL}logo.png`;
  Object.assign(logo.style, { position: "absolute", top: "0", left: "0", width: "100%", display: "block" });
  wrapper.appendChild(logo);

  const content = el("div", {
    paddingTop: "96px",
    paddingLeft: "48px",
    paddingRight: "48px",
    paddingBottom: "40px",
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
    el("div", { ...BASE_TEXT, fontSize: "11px", textAlign: "right", color: "#555" }, [
      "FM-HP-013 : Rev.01 : 1/12/2024",
    ]),
  ]);
  wrapper.appendChild(content);
  return wrapper;
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

async function renderMainContentPages(data) {
  await waitForFonts();
  const dom = buildPrintableDom(data);
  const host = el("div", {
    position: "fixed",
    top: "0",
    left: "-10000px",
    zIndex: "-1",
  }, [dom]);
  document.body.appendChild(host);

  let canvas;
  try {
    canvas = await html2canvas(dom, {
      scale: RENDER_SCALE,
      useCORS: true,
      backgroundColor: "#ffffff",
      width: PAGE_WIDTH_PX,
      windowWidth: PAGE_WIDTH_PX,
    });
  } finally {
    document.body.removeChild(host);
  }

  const pageHeightPxScaled = PAGE_HEIGHT_PX * RENDER_SCALE;
  const pageWidthPxScaled = PAGE_WIDTH_PX * RENDER_SCALE;
  const pageCount = Math.max(1, Math.ceil(canvas.height / pageHeightPxScaled));

  const pages = [];
  for (let i = 0; i < pageCount; i++) {
    const sliceCanvas = document.createElement("canvas");
    sliceCanvas.width = pageWidthPxScaled;
    sliceCanvas.height = pageHeightPxScaled;
    const ctx = sliceCanvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, sliceCanvas.width, sliceCanvas.height);
    ctx.drawImage(
      canvas,
      0,
      i * pageHeightPxScaled,
      pageWidthPxScaled,
      pageHeightPxScaled,
      0,
      0,
      pageWidthPxScaled,
      pageHeightPxScaled,
    );
    const bytes = await new Promise((resolve, reject) => {
      sliceCanvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error("Failed to encode PDF page image"));
          return;
        }
        blob.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)));
      }, "image/png");
    });
    pages.push(bytes);
  }
  return pages;
}

export async function generateReferralFormPdf(data) {
  const pdfDoc = await PDFDocument.create();

  const mainPageBytes = await renderMainContentPages(data);
  for (const bytes of mainPageBytes) {
    const page = pdfDoc.addPage([PAGE_WIDTH_PT, PAGE_HEIGHT_PT]);
    const png = await pdfDoc.embedPng(bytes);
    page.drawImage(png, { x: 0, y: 0, width: PAGE_WIDTH_PT, height: PAGE_HEIGHT_PT });
  }

  const marginPt = 36;
  const maxWidthPt = PAGE_WIDTH_PT - marginPt * 2;
  const maxHeightPt = PAGE_HEIGHT_PT - marginPt * 2;
  for (const att of data.attachments ?? []) {
    const page = pdfDoc.addPage([PAGE_WIDTH_PT, PAGE_HEIGHT_PT]);
    const png = await pdfDoc.embedPng(att.bytes);
    const scale = Math.min(1, maxWidthPt / att.width, maxHeightPt / att.height);
    const width = att.width * scale;
    const height = att.height * scale;
    page.drawImage(png, {
      x: (PAGE_WIDTH_PT - width) / 2,
      y: (PAGE_HEIGHT_PT - height) / 2,
      width,
      height,
    });
  }

  return pdfDoc.save();
}

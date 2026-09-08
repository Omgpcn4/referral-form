import {
  Document,
  Paragraph,
  TextRun,
  ImageRun,
  Header,
  Footer,
  AlignmentType,
} from "docx";
import logoUrl from "./assets/logo.png";

// Latin runs render in Calibri Light, Thai runs in Angsana New — matches the
// "majorBidi" theme fonts used throughout Referral_form_PK.docx.
const RUN_FONT = { ascii: "Calibri Light", hAnsi: "Calibri Light", cs: "Angsana New", eastAsia: "Angsana New" };

const PAGE_WIDTH = 11906;
const PAGE_HEIGHT = 16838;
const MARGIN_TOP = 431;
const MARGIN_BOTTOM = 289;
const MARGIN_LEFT = 720;
const MARGIN_RIGHT = 720;
const MARGIN_HEADER = 454;
const MARGIN_FOOTER = 284;

// Source logo is 2639x270px; keep that aspect ratio in the header banner.
const LOGO_SOURCE_WIDTH = 2639;
const LOGO_SOURCE_HEIGHT = 270;
const LOGO_DISPLAY_WIDTH = 600;
const LOGO_DISPLAY_HEIGHT = Math.round(
  (LOGO_DISPLAY_WIDTH * LOGO_SOURCE_HEIGHT) / LOGO_SOURCE_WIDTH,
);

const UNSELECTED = "○";
const SELECTED = "●";

let logoBytesPromise;
function getLogoBytes() {
  if (!logoBytesPromise) {
    logoBytesPromise = fetch(logoUrl)
      .then((res) => res.arrayBuffer())
      .then((buf) => new Uint8Array(buf));
  }
  return logoBytesPromise;
}

function s(value) {
  return value == null ? "" : String(value).trim();
}

function formatDate(isoDate) {
  if (!isoDate) return "";
  const [year, month, day] = isoDate.split("-");
  if (!year || !month || !day) return isoDate;
  return `${day}/${month}/${year}`;
}

function run(text, opts = {}) {
  return new TextRun({ text: text ?? "", font: RUN_FONT, size: 20, ...opts });
}

function labeledLine(parts, opts = {}) {
  const children = [];
  parts.forEach((part, i) => {
    if (i > 0) children.push(run("     "));
    children.push(run(`${part.label} : `, opts));
    children.push(run(s(part.value), opts));
  });
  return new Paragraph({ spacing: { after: 120 }, children });
}

function purposeLine(purpose) {
  const options = [
    { key: "diagnosis", th: "รับการวินิจฉัย", en: "Diagnosis" },
    { key: "treatment", th: "รับการรักษา", en: "Treatment" },
    { key: "ownerRequest", th: "ตามความต้องการของเจ้าของ", en: "Owner request" },
  ];
  const children = [run("เพื่อ For :  ")];
  options.forEach((opt, i) => {
    if (i > 0) children.push(run("     "));
    const mark = purpose === opt.key ? SELECTED : UNSELECTED;
    children.push(run(`${mark} ${opt.th} ${opt.en}`));
  });
  return new Paragraph({ spacing: { after: 120 }, children });
}

function numberedSection(number, label, text) {
  const paragraphs = [
    new Paragraph({
      indent: { left: 360 },
      spacing: { before: 120, after: 40 },
      children: [run(`${number}. ${label}`)],
    }),
  ];
  const lines = s(text)
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  if (lines.length === 0) {
    paragraphs.push(
      new Paragraph({ indent: { left: 360 }, children: [run("", { size: 18 })] }),
    );
  } else {
    lines.forEach((line) => {
      paragraphs.push(
        new Paragraph({ indent: { left: 360 }, children: [run(line, { size: 18 })] }),
      );
    });
  }
  return paragraphs;
}

async function buildHeader() {
  const logoBytes = await getLogoBytes();
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            type: "png",
            data: logoBytes,
            transformation: { width: LOGO_DISPLAY_WIDTH, height: LOGO_DISPLAY_HEIGHT },
            altText: {
              title: "Arak Animal Hospital Phuket",
              description: "Arak Animal Hospital Phuket logo",
              name: "logo",
            },
          }),
        ],
      }),
    ],
  });
}

function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [run("FM-HP-013 : Rev.01 : 1/12/2024", { size: 24 })],
      }),
    ],
  });
}

export async function generateReferralFormDocx(data) {
  const children = [];

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0 },
      children: [run("ใบส่งตัวสัตว์ป่วย", { bold: true, size: 32 })],
    }),
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
      children: [run("Referral Form", { bold: true, size: 32 })],
    }),
  );

  children.push(
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { after: 0 },
      children: [run("วันที่ Date : ", { size: 18 }), run(formatDate(data.date), { size: 18 })],
    }),
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { after: 160 },
      children: [run("HN : ", { size: 18 }), run(s(data.hn), { size: 18 })],
    }),
  );

  children.push(
    new Paragraph({
      spacing: { after: 160 },
      children: [run("เรียน สัตวแพทย์ผู้เกี่ยวข้อง To Whom it may concern")],
    }),
  );

  children.push(
    labeledLine([
      { label: "ชื่อสัตว์เลี้ยง Pet's name", value: data.petName },
      { label: "ชนิด Species", value: data.species },
      { label: "เพศ Gender", value: data.gender },
    ]),
  );
  children.push(
    labeledLine([
      { label: "พันธุ์ Breed", value: data.breed },
      { label: "อายุ Age", value: data.age },
    ]),
  );
  children.push(labeledLine([{ label: "ชื่อเจ้าของสัตว์เลี้ยง Owner's name", value: data.ownerName }]));
  children.push(purposeLine(data.purpose));

  children.push(...numberedSection(1, "ประวัติอาการ History", data.history));
  children.push(
    ...numberedSection(2, "อาการป่วยปัจจุบัน/ผลการตรวจร่างกาย Physical Examination", data.physicalExam),
  );
  children.push(...numberedSection(3, "ผลการตรวจทางห้องปฏิบัติการ Laboratory", data.laboratory));
  children.push(...numberedSection(4, "การวินิจฉัยเบื้องต้น Diagnosis", data.diagnosis));
  children.push(...numberedSection(5, "การรักษา Treatment", data.treatment));
  children.push(...numberedSection(6, "รายละเอียดอื่นๆ Others", data.others));

  children.push(
    new Paragraph({
      spacing: { before: 300, after: 100 },
      alignment: AlignmentType.CENTER,
      children: [run("เรียนมาเพื่อทราบ Sincerely,")],
    }),
  );
  children.push(new Paragraph({ children: [run("")] }));
  children.push(new Paragraph({ children: [run("")] }));
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [run(`(${s(data.vetName)})`, { size: 18 })],
    }),
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        run("ใบอนุญาตเลขที่ Veterinary License No. ", { size: 18 }),
        run(s(data.licenseNo), { size: 18 }),
      ],
    }),
  );

  return new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
            margin: {
              top: MARGIN_TOP,
              bottom: MARGIN_BOTTOM,
              left: MARGIN_LEFT,
              right: MARGIN_RIGHT,
              header: MARGIN_HEADER,
              footer: MARGIN_FOOTER,
            },
          },
        },
        headers: { default: await buildHeader() },
        footers: { default: buildFooter() },
        children,
      },
    ],
  });
}

export function buildFilename(data) {
  const petName = s(data.petName).replace(/[^\p{L}\p{N}]+/gu, "_") || "Pet";
  const date = s(data.date) || new Date().toISOString().slice(0, 10);
  return `Referral_Form_${petName}_${date}.docx`;
}

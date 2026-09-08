import {
  Document,
  Paragraph,
  TextRun,
  ImageRun,
  Header,
  Footer,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  TabStopType,
  BorderStyle,
  VerticalAlign,
} from "docx";
import logoUrl from "./assets/logo.png";

const FONT = "Sarabun";

// Source logo is 2639x270px; keep that aspect ratio when placed in the header.
const LOGO_SOURCE_WIDTH = 2639;
const LOGO_SOURCE_HEIGHT = 270;
const LOGO_DISPLAY_WIDTH = 220;
const LOGO_DISPLAY_HEIGHT = Math.round(
  (LOGO_DISPLAY_WIDTH * LOGO_SOURCE_HEIGHT) / LOGO_SOURCE_WIDTH,
);

let logoBytesPromise;
function getLogoBytes() {
  if (!logoBytesPromise) {
    logoBytesPromise = fetch(logoUrl)
      .then((res) => res.arrayBuffer())
      .then((buf) => new Uint8Array(buf));
  }
  return logoBytesPromise;
}

const PAGE_WIDTH = 11909;
const PAGE_HEIGHT = 16834;
const MARGIN_TOP = 864;
const MARGIN_BOTTOM = 864;
const MARGIN_LEFT = 1080;
const MARGIN_RIGHT = 1080;
const MARGIN_HEADER = 720;
const MARGIN_FOOTER = 720;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT;

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NO_BORDERS = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
};

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
  return new TextRun({ text: text ?? "", font: FONT, size: 22, ...opts });
}

function labelValueRuns(label, value, opts = {}) {
  return [run(`${label} `, opts), run(s(value))];
}

function fieldLine(parts) {
  const children = [];
  parts.forEach((part, i) => {
    if (i > 0) children.push(run("\t"));
    children.push(...labelValueRuns(part.label, part.value));
  });
  return new Paragraph({
    spacing: { after: 80 },
    tabStops: [
      { type: TabStopType.LEFT, position: Math.round((CONTENT_WIDTH * 1) / 3) },
      { type: TabStopType.LEFT, position: Math.round((CONTENT_WIDTH * 2) / 3) },
    ],
    children,
  });
}

async function buildHeader() {
  const logoBytes = await getLogoBytes();
  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    borders: {
      top: NO_BORDER,
      bottom: NO_BORDER,
      left: NO_BORDER,
      right: NO_BORDER,
      insideHorizontal: NO_BORDER,
      insideVertical: NO_BORDER,
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: Math.round(CONTENT_WIDTH * 0.38), type: WidthType.DXA },
            borders: NO_BORDERS,
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    type: "png",
                    data: logoBytes,
                    transformation: {
                      width: LOGO_DISPLAY_WIDTH,
                      height: LOGO_DISPLAY_HEIGHT,
                    },
                    altText: {
                      title: "Arak Animal Hospital Phuket",
                      description: "Arak Animal Hospital Phuket logo",
                      name: "logo",
                    },
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: Math.round(CONTENT_WIDTH * 0.62), type: WidthType.DXA },
            borders: NO_BORDERS,
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  run(
                    "100/16 Mu 7, Si Sunthon Subdistrict, Thalang District, Phuket 83110",
                    { size: 18 },
                  ),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  run(
                    "100/16 หมู่ 7 ตำบลศรีสุนทร อำเภอถลาง จังหวัดภูเก็ต 83110",
                    { size: 18 },
                  ),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  return new Header({
    children: [table, new Paragraph({ children: [] })],
  });
}

function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [run("FM-HP-002 : Rev.01 : 1/12/2024", { size: 16, font: FONT })],
      }),
    ],
  });
}

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 200, after: 120 },
    children: [run(text, { bold: true, size: 24 })],
  });
}

function multiLineSection(number, label, text) {
  const paragraphs = [
    new Paragraph({
      spacing: { before: 160, after: 60 },
      children: [run(`${number}. ${label}`, { size: 22 })],
    }),
  ];
  const lines = s(text)
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  if (lines.length === 0) {
    paragraphs.push(new Paragraph({ children: [run("")] }));
  } else {
    lines.forEach((line) => {
      paragraphs.push(new Paragraph({ children: [run(line)] }));
    });
  }
  return paragraphs;
}

export async function generateHealthCertificateDocx(data) {
  const children = [];

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [run("ใบรายงานผลการตรวจ / Health certificate", { bold: true, size: 32 })],
    }),
  );

  children.push(
    new Paragraph({
      spacing: { after: 200 },
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH }],
      children: [
        run("เรียน เจ้าของสัตว์เลี้ยง/To Whom It May Concern"),
        data.recipientName ? run(` ${s(data.recipientName)}`) : run(""),
        run("\t"),
        run("วันที่/Date : "),
        run(formatDate(data.date)),
      ],
    }),
  );

  children.push(sectionHeader("ข้อมูลเจ้าของสัตว์/Owner details"));
  children.push(fieldLine([{ label: "ชื่อ/Name :", value: data.ownerName }]));
  children.push(fieldLine([{ label: "ที่อยู่/Address :", value: data.ownerAddress }]));
  children.push(fieldLine([{ label: "โทรศัพท์/Phone :", value: data.ownerPhone }]));

  children.push(sectionHeader("ข้อมูลสัตว์ป่วย/Animal's details"));
  children.push(
    fieldLine([
      { label: "ชื่อสัตว์เลี้ยง/Name :", value: data.animalName },
      { label: "HN :", value: data.hn },
      { label: "ชนิด/Species :", value: data.species },
      { label: "เพศ/Gender :", value: data.gender },
    ]),
  );

  children.push(
    new Paragraph({
      spacing: { after: 80 },
      children: [
        run("อายุ/Age : "),
        run(s(data.ageYears)),
        run(" ปี(Years) "),
        run(s(data.ageMonths)),
        run(" เดือน(Months) "),
        run(s(data.ageDays)),
        run(" วัน(Days)"),
      ],
    }),
  );
  children.push(
    fieldLine([
      { label: "พันธุ์/Breed:", value: data.breed },
      { label: "Microchip number :", value: data.microchip },
    ]),
  );

  children.push(...multiLineSection(1, "ประวัติสัตว์ป่วย/History", data.history));
  children.push(...multiLineSection(2, "ผลการตรวจร่างกาย/Physical examination", data.physicalExam));
  children.push(...multiLineSection(3, "ผลการตรวจทางห้องปฏิบัติการ/Laboratory", data.laboratory));
  children.push(...multiLineSection(4, "การวินิจฉัย/Diagnosis", data.diagnosis));
  children.push(...multiLineSection(5, "การรักษา/Treatment", data.treatment));
  children.push(...multiLineSection(6, "หมายเหตุหรือรายละเอียดอื่นๆ/Others", data.others));

  children.push(
    new Paragraph({
      spacing: { before: 320 },
      alignment: AlignmentType.CENTER,
      children: [run("เรียนมาเพื่อทราบ/Yours sincerely")],
    }),
  );
  children.push(new Paragraph({ children: [run("")] }));
  children.push(new Paragraph({ children: [run("")] }));
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [run("Veterinarian")],
    }),
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [run(`(${s(data.vetName)})`)],
    }),
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [run("ใบอนุญาตเลขที่/License No. "), run(s(data.licenseNo))],
    }),
  );

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: FONT, size: 22 },
        },
      },
    },
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
  const animalName = s(data.animalName).replace(/[^\p{L}\p{N}]+/gu, "_") || "Animal";
  const date = s(data.date) || new Date().toISOString().slice(0, 10);
  return `Health_Certificate_${animalName}_${date}.docx`;
}

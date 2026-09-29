// Single description of the referral letter's layout, shared by the .docx
// and .pdf generators so the two outputs can't drift apart. Units follow
// Word: lengths in twips (1/20 pt), font sizes in half-points.

export const FONT_NAME = "TH Sarabun PSK";
export const DEFAULT_SIZE = 26;

const PAGE_WIDTH = 11906;

// Source logo is 2639x270px. It floats relative to the physical page and
// spans its full width, so it bleeds edge-to-edge like the original
// template's banner.
const EMU_PER_PX = 9525;
const EMU_PER_TWIP = 635;
const LOGO_WIDTH_PX = Math.round(PAGE_WIDTH / 15);
export const LOGO = {
  widthPx: LOGO_WIDTH_PX,
  heightPx: Math.round((LOGO_WIDTH_PX * 270) / 2639),
  topOffsetEmu: 150000,
};
const LOGO_BOTTOM_TWIPS = Math.round(
  (LOGO.topOffsetEmu + LOGO.heightPx * EMU_PER_PX) / EMU_PER_TWIP,
);

// The top margin clears the logo, so text never runs under it on any page
// (continuation pages and attachment pages included).
export const PAGE = {
  width: PAGE_WIDTH,
  height: 16838,
  marginTop: LOGO_BOTTOM_TWIPS + 360,
  marginBottom: 289,
  marginLeft: 720,
  marginRight: 720,
  marginHeader: 454,
  marginFooter: 284,
};

const CONTENT_WIDTH = PAGE.width - PAGE.marginLeft - PAGE.marginRight;
const CONTENT_HEIGHT = PAGE.height - PAGE.marginTop - PAGE.marginBottom;
export const ATTACHMENT_MAX_WIDTH_PX = Math.round(CONTENT_WIDTH / 15);
export const ATTACHMENT_MAX_HEIGHT_PX = Math.round((CONTENT_HEIGHT - 150) / 15);

// Numbered-list level: the paragraph sets its own left indent, and Word
// takes the hanging indent from here.
export const NUMBERED_LIST = { left: 720, hanging: 360 };

const UNSELECTED = "○";
const SELECTED = "●";

function s(value) {
  return value == null ? "" : String(value).trim();
}

export function formatDate(isoDate) {
  if (!isoDate) return "";
  const [year, month, day] = isoDate.split("-");
  if (!year || !month || !day) return isoDate;
  return `${day}/${month}/${year}`;
}

// A paragraph: { runs: [{ text, size?, bold? }], align?, indent?: { left,
// hanging? }, spacing?: { before?, after? }, list?: { type: "bullet" } |
// { type: "number", group } }. Absent keys mean "Word default".
function para(runs, props = {}) {
  return { runs, ...props };
}

function labeledLine(parts) {
  const runs = [];
  parts.forEach((part, i) => {
    if (i > 0) runs.push({ text: "     " });
    runs.push({ text: `${part.label}: ` });
    runs.push({ text: s(part.value) });
  });
  return para(runs, { spacing: { after: 120 } });
}

function purposeLine(purpose) {
  const options = [
    { key: "diagnosis", th: "รับการวินิจฉัย", en: "Diagnosis" },
    { key: "treatment", th: "รับการรักษา", en: "Treatment" },
    { key: "ownerRequest", th: "ตามความต้องการของเจ้าของ", en: "Owner request" },
  ];
  const runs = [{ text: "เพื่อ For :  " }];
  options.forEach((opt, i) => {
    if (i > 0) runs.push({ text: "     " });
    const mark = purpose === opt.key ? SELECTED : UNSELECTED;
    runs.push({ text: `${mark} ${opt.th} ${opt.en}` });
  });
  return para(runs, { spacing: { after: 120 } });
}

function numberedSection(number, label, blocks, nextListGroup) {
  const paragraphs = [
    para([{ text: `${number}. ${label}` }], {
      indent: { left: 360 },
      spacing: { before: 120, after: 40 },
    }),
  ];
  if (!blocks || blocks.length === 0) {
    paragraphs.push(para([{ text: "", size: 22 }], { indent: { left: 360 } }));
    return paragraphs;
  }
  let previousType = null;
  let group = null;
  for (const block of blocks) {
    const runs = [{ text: block.text, size: 22, bold: block.bold || undefined }];
    if (block.type === "bullet") {
      paragraphs.push(para(runs, { indent: { left: 1080, hanging: 360 }, list: { type: "bullet" } }));
    } else if (block.type === "number") {
      // Each run of consecutive numbered items restarts at 1.
      if (previousType !== "number") group = nextListGroup();
      paragraphs.push(para(runs, { indent: { left: 360 }, list: { type: "number", group } }));
    } else {
      paragraphs.push(para(runs, { indent: { left: 360 } }));
    }
    previousType = block.type;
  }
  return paragraphs;
}

export function buildDocumentModel(data) {
  let listGroups = 0;
  const nextListGroup = () => ++listGroups;

  return [
    para([{ text: "ใบส่งตัวสัตว์ป่วย", bold: true, size: 32 }], {
      align: "center",
      spacing: { after: 0 },
    }),
    para([{ text: "Referral Form", bold: true, size: 32 }], {
      align: "center",
      spacing: { after: 160 },
    }),
    para([{ text: "วันที่ Date: ", size: 22 }, { text: formatDate(data.date), size: 22 }], {
      align: "right",
      spacing: { after: 0 },
    }),
    para([{ text: "HN: ", size: 22 }, { text: s(data.hn), size: 22 }], {
      align: "right",
      spacing: { after: 160 },
    }),
    para([{ text: "เรียน สัตวแพทย์ผู้เกี่ยวข้อง To Whom it may concern" }], {
      spacing: { after: 160 },
    }),
    labeledLine([
      { label: "ชื่อสัตว์เลี้ยง Pet's name", value: data.petName },
      { label: "ชนิด Species", value: data.species },
      { label: "เพศ Gender", value: data.gender },
    ]),
    labeledLine([
      { label: "พันธุ์ Breed", value: data.breed },
      { label: "อายุ Age", value: data.age },
    ]),
    labeledLine([{ label: "ชื่อเจ้าของสัตว์เลี้ยง Owner's name", value: data.ownerName }]),
    purposeLine(data.purpose),
    ...numberedSection(1, "ประวัติอาการ History", data.history, nextListGroup),
    ...numberedSection(
      2,
      "อาการป่วยปัจจุบัน/ผลการตรวจร่างกาย Physical Examination",
      data.physicalExam,
      nextListGroup,
    ),
    ...numberedSection(3, "ผลการตรวจทางห้องปฏิบัติการ Laboratory", data.laboratory, nextListGroup),
    ...numberedSection(4, "การวินิจฉัยเบื้องต้น Diagnosis", data.diagnosis, nextListGroup),
    ...numberedSection(5, "การรักษา Treatment", data.treatment, nextListGroup),
    ...numberedSection(6, "รายละเอียดอื่นๆ Others", data.others, nextListGroup),
    para([{ text: "เรียนมาเพื่อทราบ Sincerely," }], {
      align: "center",
      spacing: { before: 300, after: 100 },
    }),
    para([{ text: "" }]),
    para([{ text: "" }]),
    para([{ text: `(${s(data.vetName)})`, size: 22 }], {
      align: "center",
      spacing: { after: 40 },
    }),
    para(
      [
        { text: "ใบอนุญาตเลขที่ Veterinary License No. ", size: 22 },
        { text: s(data.licenseNo), size: 22 },
      ],
      { align: "center" },
    ),
  ];
}

// Attachment pages: each image on its own page, below the logo, scaled to
// fit the content area. Sizes in CSS px (96dpi), as the docx expects.
export function attachmentLayout(att) {
  const scale = Math.min(1, ATTACHMENT_MAX_WIDTH_PX / att.width, ATTACHMENT_MAX_HEIGHT_PX / att.height);
  return { width: Math.round(att.width * scale), height: Math.round(att.height * scale) };
}

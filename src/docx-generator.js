import {
  Document,
  Paragraph,
  TextRun,
  ImageRun,
  Header,
  AlignmentType,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  TextWrappingType,
  LevelFormat,
} from "docx";
import logoUrl from "./assets/logo.png";
import {
  FONT_NAME,
  DEFAULT_SIZE,
  PAGE,
  LOGO,
  NUMBERED_LIST,
  buildDocumentModel,
  attachmentLayout,
} from "./document-model.js";

const NUMBERED_LIST_REFERENCE = "clinical-numbered-list";

const ALIGNMENT = {
  left: undefined,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
};

let logoBytesPromise;
function getLogoBytes() {
  if (!logoBytesPromise) {
    logoBytesPromise = fetch(logoUrl)
      .then((res) => res.arrayBuffer())
      .then((buf) => new Uint8Array(buf));
  }
  return logoBytesPromise;
}

function toParagraph(p) {
  const options = {
    children: p.runs.map(
      (r) =>
        new TextRun({
          text: r.text ?? "",
          font: FONT_NAME,
          size: r.size ?? DEFAULT_SIZE,
          ...(r.bold ? { bold: true } : {}),
        }),
    ),
  };
  if (ALIGNMENT[p.align]) options.alignment = ALIGNMENT[p.align];
  if (p.indent) options.indent = p.indent;
  if (p.spacing) options.spacing = p.spacing;
  if (p.list?.type === "bullet") options.bullet = { level: 0 };
  if (p.list?.type === "number") {
    options.numbering = { reference: NUMBERED_LIST_REFERENCE, level: 0, instance: p.list.group };
  }
  return new Paragraph(options);
}

function buildAttachmentPages(attachments) {
  return (attachments ?? []).map((att) => {
    const { width, height } = attachmentLayout(att);
    return new Paragraph({
      pageBreakBefore: true,
      alignment: AlignmentType.CENTER,
      children: [
        new ImageRun({
          type: "png",
          data: att.bytes,
          transformation: { width, height },
          altText: { title: att.label, description: att.label, name: att.label },
        }),
      ],
    });
  });
}

async function buildHeader() {
  const logoBytes = await getLogoBytes();
  return new Header({
    children: [
      new Paragraph({
        children: [
          new ImageRun({
            type: "png",
            data: logoBytes,
            transformation: { width: LOGO.widthPx, height: LOGO.heightPx },
            floating: {
              horizontalPosition: {
                relative: HorizontalPositionRelativeFrom.PAGE,
                offset: 0,
              },
              verticalPosition: {
                relative: VerticalPositionRelativeFrom.PAGE,
                offset: LOGO.topOffsetEmu,
              },
              wrap: { type: TextWrappingType.NONE },
              allowOverlap: true,
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
  });
}

export async function generateReferralFormDocx(data) {
  const children = [
    ...buildDocumentModel(data).map(toParagraph),
    ...buildAttachmentPages(data.attachments),
  ];

  return new Document({
    numbering: {
      config: [
        {
          reference: NUMBERED_LIST_REFERENCE,
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.START,
              style: { paragraph: { indent: NUMBERED_LIST } },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE.width, height: PAGE.height },
            margin: {
              top: PAGE.marginTop,
              bottom: PAGE.marginBottom,
              left: PAGE.marginLeft,
              right: PAGE.marginRight,
              header: PAGE.marginHeader,
              footer: PAGE.marginFooter,
            },
          },
        },
        headers: { default: await buildHeader() },
        children,
      },
    ],
  });
}

export function buildFilename(data) {
  const petName =
    String(data.petName ?? "")
      .replace(/[\\/:*?"<>|]+/g, " ")
      .replace(/\s+/g, " ")
      .trim() || "Pet";
  const isoDate = String(data.date ?? "").trim() || new Date().toISOString().slice(0, 10);
  const [year, month, day] = isoDate.split("-");
  // "/" isn't allowed in filenames, so DD/MM/YY is written as DD-MM-YY.
  const date = `${day}-${month}-${year.slice(-2)}`;
  return `Referral letter ${petName} ${date}.docx`;
}

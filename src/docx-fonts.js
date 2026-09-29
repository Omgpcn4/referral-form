import JSZip from "jszip";

// Embeds TH Sarabun New (regular and bold) in a packed .docx so Word shows
// the same font as the PDF on any computer, installed or not. The docx
// library can only embed a regular weight and labels the font as
// Latin-only, which could make Word skip it for Thai, so the font table is
// written here with the font's own OS/2 values.

export const EMBEDDED_FONT_NAME = "TH Sarabun New";

const FONT_FILES = { regular: "THSarabunNew.ttf", bold: "THSarabunNew-Bold.ttf" };

// From the font's OS/2 table: PANOSE, Unicode ranges and code pages
// (the code pages include Thai, 874).
const FONT_TABLE_ENTRY = {
  panose1: "020B0500040200020003",
  charset: "DE",
  sig: {
    usb0: "A100006F",
    usb1: "5000205A",
    usb2: "00000000",
    usb3: "00000000",
    csb0: "60010183",
    csb1: "80000000",
  },
};

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const FONT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";

let fontBytesPromise;
function fetchFontBytes() {
  if (!fontBytesPromise) {
    const base = import.meta.env.BASE_URL;
    fontBytesPromise = Promise.all(
      Object.entries(FONT_FILES).map(async ([style, file]) => {
        const res = await fetch(`${base}fonts/${file}`);
        if (!res.ok) throw new Error(`Failed to load font ${file}`);
        return [style, new Uint8Array(await res.arrayBuffer())];
      }),
    )
      .then(Object.fromEntries)
      .catch((err) => {
        fontBytesPromise = null;
        throw err;
      });
  }
  return fontBytesPromise;
}

// ECMA-376 Part 2 font obfuscation: XOR the first 32 bytes with the GUID
// key's bytes in reverse order.
function obfuscate(bytes, guid) {
  const key = guid.replace(/-/g, "").match(/../g).map((h) => parseInt(h, 16)).reverse();
  const out = bytes.slice();
  for (let i = 0; i < 32; i++) out[i] ^= key[i % key.length];
  return out;
}

function newGuid() {
  return crypto.randomUUID().toUpperCase();
}

function fontTableXml(keys) {
  const { panose1, charset, sig } = FONT_TABLE_ENTRY;
  const sigAttrs = Object.entries(sig).map(([k, v]) => `w:${k}="${v}"`).join(" ");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:fonts xmlns:w="${W_NS}" xmlns:r="${R_NS}">` +
    `<w:font w:name="${EMBEDDED_FONT_NAME}">` +
    `<w:panose1 w:val="${panose1}"/><w:charset w:val="${charset}"/>` +
    `<w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig ${sigAttrs}/>` +
    `<w:embedRegular r:id="rIdFontRegular" w:fontKey="{${keys.regular}}"/>` +
    `<w:embedBold r:id="rIdFontBold" w:fontKey="{${keys.bold}}"/>` +
    `</w:font></w:fonts>`
  );
}

function fontTableRelsXml() {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rIdFontRegular" Type="${FONT_REL_TYPE}" Target="fonts/THSarabunNew.odttf"/>` +
    `<Relationship Id="rIdFontBold" Type="${FONT_REL_TYPE}" Target="fonts/THSarabunNew-Bold.odttf"/>` +
    `</Relationships>`
  );
}

// <w:embedTrueTypeFonts/> keeps the fonts embedded if the file is re-saved
// in Word. Schema order puts it after <w:displayBackgroundShape/>.
function withEmbedSetting(settings) {
  if (settings.includes("<w:embedTrueTypeFonts")) return settings;
  if (settings.includes("<w:displayBackgroundShape/>")) {
    return settings.replace("<w:displayBackgroundShape/>", "<w:displayBackgroundShape/><w:embedTrueTypeFonts/>");
  }
  return settings.replace(/(<w:settings\b[^>]*>)/, "$1<w:embedTrueTypeFonts/>");
}

const ODTTF_TYPE = "application/vnd.openxmlformats-officedocument.obfuscatedFont";

function withOdttfContentType(types) {
  if (types.includes('Extension="odttf"')) return types;
  return types.replace("</Types>", `<Default Extension="odttf" ContentType="${ODTTF_TYPE}"/></Types>`);
}

export async function embedFonts(docxBlob) {
  const [fonts, zip] = await Promise.all([fetchFontBytes(), JSZip.loadAsync(docxBlob)]);
  const keys = { regular: newGuid(), bold: newGuid() };

  zip.file("word/fonts/THSarabunNew.odttf", obfuscate(fonts.regular, keys.regular));
  zip.file("word/fonts/THSarabunNew-Bold.odttf", obfuscate(fonts.bold, keys.bold));
  zip.file("word/fontTable.xml", fontTableXml(keys));
  zip.file("word/_rels/fontTable.xml.rels", fontTableRelsXml());
  zip.file("word/settings.xml", withEmbedSetting(await zip.file("word/settings.xml").async("string")));
  zip.file("[Content_Types].xml", withOdttfContentType(await zip.file("[Content_Types].xml").async("string")));

  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
  });
}

// Port of HarfBuzz's Thai fallback shaping (hb-ot-shaper-thai.cc), which
// the harfbuzzjs WebAssembly build leaves out. Older Thai fonts, TH Sarabun
// included, have no OpenType rules for Thai; they instead carry positioned
// variants of each mark in the Private Use Area (lowered tone marks,
// left-shifted marks for tall consonants like ป ฝ ฟ, ญ/ฐ without the
// descender). Browsers and word processors pick those variants with this
// state machine, so the PDF must too, or marks collide with letters.
//
// Runs on the text before shaping: decomposes SARA AM and swaps in PUA
// variants, returning text that HarfBuzz then maps straight to glyphs.

const NC = 0; // plain consonant
const AC = 1; // ascender consonant
const RC = 2; // removable-descender consonant
const DC = 3; // strict-descender consonant
const NOT_CONSONANT = 4;

function consonantType(u) {
  if (u === 0x0e1b || u === 0x0e1d || u === 0x0e1f) return AC;
  if (u === 0x0e0d || u === 0x0e10) return RC;
  if (u === 0x0e0e || u === 0x0e0f) return DC;
  if (u >= 0x0e01 && u <= 0x0e2e) return NC;
  return NOT_CONSONANT;
}

const AV = 0; // above vowel
const BV = 1; // below vowel
const T = 2; // tone mark
const NOT_MARK = 3;

function markType(u) {
  if (u === 0x0e31 || (u >= 0x0e34 && u <= 0x0e37) || u === 0x0e47 || u === 0x0e4d || u === 0x0e4e) {
    return AV;
  }
  if (u >= 0x0e38 && u <= 0x0e3a) return BV;
  if (u >= 0x0e48 && u <= 0x0e4c) return T;
  return NOT_MARK;
}

const NOP = 0;
const SD = 1; // shift down
const SL = 2; // shift left
const SDL = 3; // shift down-left
const RD = 4; // remove descender

// [unicode, Windows PUA, Mac PUA]
const PUA = {
  [SD]: [
    [0x0e48, 0xf70a, 0xf88b], [0x0e49, 0xf70b, 0xf88e], [0x0e4a, 0xf70c, 0xf891],
    [0x0e4b, 0xf70d, 0xf894], [0x0e4c, 0xf70e, 0xf897], [0x0e38, 0xf718, 0xf89b],
    [0x0e39, 0xf719, 0xf89c], [0x0e3a, 0xf71a, 0xf89d],
  ],
  [SDL]: [
    [0x0e48, 0xf705, 0xf88c], [0x0e49, 0xf706, 0xf88f], [0x0e4a, 0xf707, 0xf892],
    [0x0e4b, 0xf708, 0xf895], [0x0e4c, 0xf709, 0xf898],
  ],
  [SL]: [
    [0x0e48, 0xf713, 0xf88a], [0x0e49, 0xf714, 0xf88d], [0x0e4a, 0xf715, 0xf890],
    [0x0e4b, 0xf716, 0xf893], [0x0e4c, 0xf717, 0xf896], [0x0e31, 0xf710, 0xf884],
    [0x0e34, 0xf701, 0xf885], [0x0e35, 0xf702, 0xf886], [0x0e36, 0xf703, 0xf887],
    [0x0e37, 0xf704, 0xf888], [0x0e47, 0xf712, 0xf889], [0x0e4d, 0xf711, 0xf899],
  ],
  [RD]: [[0x0e0d, 0xf70f, 0xf89a], [0x0e10, 0xf700, 0xf89e]],
};

const PUA_TO_THAI = new Map([
  ...Object.values(PUA).flatMap((mappings) =>
    mappings.flatMap(([base, win, mac]) => [[win, [base]], [mac, [base]]]),
  ),
  // Ligature glyphs these fonts map from the Mac PUA: ฤๅ, ฦๅ, ๏"
  [0xf881, [0x0e24, 0x0e45]],
  [0xf882, [0x0e26, 0x0e45]],
  [0xf880, [0x0e4f, 0x22]],
]);

// The Thai text a PUA variant or ligature glyph stands for, so text copied
// from a PDF reads as Thai rather than private-use codes.
export function thaiForPua(u) {
  return PUA_TO_THAI.get(u);
}

function puaVariant(u, action, hasGlyph) {
  const mapping = PUA[action]?.find(([base]) => base === u);
  if (!mapping) return u;
  const [, win, mac] = mapping;
  if (hasGlyph(win)) return win;
  if (hasGlyph(mac)) return mac;
  return u;
}

const T0 = 0;
const T1 = 1;
const T2 = 2;
const T3 = 3;
const ABOVE_START = [T0, T1, T0, T0, T3]; // by consonant type
const ABOVE_MACHINE = [
  //  AV          BV          T
  [[NOP, T3], [NOP, T0], [SD, T3]], // T0
  [[SL, T2], [NOP, T1], [SDL, T2]], // T1
  [[NOP, T3], [NOP, T2], [SL, T3]], // T2
  [[NOP, T3], [NOP, T3], [NOP, T3]], // T3
];

const B0 = 0;
const B1 = 1;
const B2 = 2;
const BELOW_START = [B0, B0, B1, B2, B2];
const BELOW_MACHINE = [
  //  AV          BV          T
  [[NOP, B0], [NOP, B2], [NOP, B0]], // B0
  [[NOP, B1], [RD, B2], [NOP, B1]], // B1
  [[NOP, B2], [SD, B2], [NOP, B2]], // B2
];

const SARA_AM = 0x0e33;
const NIKHAHIT = 0x0e4d;
const SARA_AA = 0x0e32;

function isAboveBaseMark(u) {
  const x = u & ~0x0080; // Lao shares the layout
  return (x >= 0x0e34 && x <= 0x0e37) || (x >= 0x0e47 && x <= 0x0e4e) || x === 0x0e31 || x === 0x0e3b;
}

// SARA AM becomes NIKHAHIT + SARA AA, with the NIKHAHIT moved back before
// any above-base marks, e.g. <0E14 0E4B 0E33> -> <0E14 0E4D 0E4B 0E32>.
function decomposeSaraAm(codepoints) {
  const out = [];
  for (const u of codepoints) {
    if (u !== SARA_AM) {
      out.push(u);
      continue;
    }
    let start = out.length;
    while (start > 0 && isAboveBaseMark(out[start - 1])) start--;
    out.splice(start, 0, NIKHAHIT);
    out.push(SARA_AA);
  }
  return out;
}

export function thaiFallbackShape(text, hasGlyph) {
  const cps = decomposeSaraAm(Array.from(text, (ch) => ch.codePointAt(0)));
  let above = ABOVE_START[NOT_CONSONANT];
  let below = BELOW_START[NOT_CONSONANT];
  let base = 0;
  for (let i = 0; i < cps.length; i++) {
    const mt = markType(cps[i]);
    if (mt === NOT_MARK) {
      const ct = consonantType(cps[i]);
      above = ABOVE_START[ct];
      below = BELOW_START[ct];
      base = i;
      continue;
    }
    const [aboveAction, aboveNext] = ABOVE_MACHINE[above][mt];
    const [belowAction, belowNext] = BELOW_MACHINE[below][mt];
    above = aboveNext;
    below = belowNext;
    // At most one of the two actions is not NOP.
    const action = aboveAction !== NOP ? aboveAction : belowAction;
    if (action === RD) cps[base] = puaVariant(cps[base], action, hasGlyph);
    else if (action !== NOP) cps[i] = puaVariant(cps[i], action, hasGlyph);
  }
  return String.fromCodePoint(...cps);
}

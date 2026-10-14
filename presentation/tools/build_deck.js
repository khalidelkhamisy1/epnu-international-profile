// Builds EPNU_International_Profile.pptx — design & copy per ../PLAN.md, facts per ../../EPNU_International_Relations_Briefing.md
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa6");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, process.env.NO_NOTES ? "EPNU_International_Profile_SHARE.pptx" : "EPNU_International_Profile.pptx");
const ASSETS = path.join(ROOT, "assets"); // selection paths are relative to assets/
const SEL = readJson(path.join(ROOT, "assets", "image_selection.json"), {});
const FACSEL = readJson(path.join(ROOT, "assets", "faculty_image_selection.json"), {});
for (const [k, v] of Object.entries(FACSEL)) SEL["fac_" + k] = v;
const FACPAGES = readJson(path.join(ROOT, "content", "faculty_pages.json"), {});
const RESEARCH = readJson(path.join(ROOT, "content", "research_slide.json"), null);
const ALT = { campus: "EPNU campus building", president: "Prof. Dr. Atef Alam El-Din, University President", lab: "Students in a modern EPNU laboratory", tech: "EPNU computer science students", health: "EPNU health sciences students in clinical training", pharma: "EPNU pharmacy students in the lab", h57357: "EPNU students at Children's Cancer Hospital 57357", students1: "EPNU student team at a competition", students2: "EPNU students with certificates", students3: "EPNU students with a certificate", national: "Partnership meeting with EPNU delegation" };
const NOTES = Object.assign(readJson(path.join(ROOT, "content", "speaker_notes.json"), {}), readJson(path.join(ROOT, "content", "speaker_notes_extra.json"), {}));

const HEX = {
  navy: "071C35", deep: "1B3C73", blue: "0077FF", cyan: "00D4FF",
  ice: "EEF6FF", off: "F8FAFC", slate: "64748B", ink: "0F172A", white: "FFFFFF", line: "D6E4F5",
};
const THEME = {
  name: "EPNU",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.off,
    accent1: HEX.blue, accent2: HEX.cyan, accent3: HEX.deep, accent4: HEX.ice,
    accent5: HEX.slate, accent6: "10283F", hlink: HEX.blue, folHlink: HEX.deep,
  },
};
const W = 13.333, H = 7.5, MX = 0.6;

function readJson(p, dflt) { try { return JSON.parse(fs.readFileSync(p, "utf8")); } catch { return dflt; } }

// ---------- asset helpers ----------
const iconCache = {};
async function icon(name, color = HEX.white, px = 256) {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: px }));
  const buf = await sharp(Buffer.from(svg)).resize(px, px).png().toBuffer();
  return (iconCache[key] = "image/png;base64," + buf.toString("base64"));
}

function imgFile(key) {
  const sel = SEL[key];
  if (!sel) return null;
  const f = path.join(ASSETS, typeof sel === "string" ? sel : sel.file);
  return fs.existsSync(f) ? f : null;
}
function imgFocus(key) { const s = SEL[key]; return (s && s.focus) || "attention"; }

// cover-crop to w×h inches, optional rounded corners (motif). Returns data URI.
async function framed(key, w, h, radius = 0.18, dpi = 110) {
  const f = imgFile(key);
  if (!f) return null;
  const pw = Math.round(w * dpi), ph = Math.round(h * dpi);
  let img = sharp(f).rotate().resize(pw, ph, { fit: "cover", position: imgFocus(key) });
  if (radius > 0) {
    const r = Math.round(radius * dpi);
    const mask = Buffer.from(`<svg width="${pw}" height="${ph}"><rect x="0" y="0" width="${pw}" height="${ph}" rx="${r}" ry="${r}"/></svg>`);
    const buf = await img.composite([{ input: mask, blend: "dest-in" }]).png().toBuffer();
    return "image/png;base64," + buf.toString("base64");
  }
  const buf = await img.jpeg({ quality: 82 }).toBuffer();
  return "image/jpeg;base64," + buf.toString("base64");
}

async function logoData() {
  const f = imgFile("logo");
  if (!f) return null;
  const buf = await sharp(f).resize(400, 400, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- build ----------
(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "EPNU";
  pres.company = "East Port Said National University";
  pres.title = "East Port Said National University — International Profile";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const LOGO = await logoData();

  // ----- layouts -----
  const footerObjs = [
    { text: { text: "East Port Said National University  |  International Profile", options: { x: MX, y: 7.0, w: 7, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true } } },
  ];
  pres.defineSlideMaster({
    title: "CONTENT_LIGHT",
    background: { color: C.background2 },
    objects: [
      ...footerObjs,
      ...(LOGO ? [{ image: { data: LOGO, x: W - MX - 0.42, y: 0.42, w: 0.5, h: 0.5 } }] : []),
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.38, w: 11.0, h: 0.62, fontSize: 30, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, valign: "middle", align: "left" }, text: "" } },
      { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 1.02, w: 11.0, h: 0.4, fontSize: 15, color: C.accent1, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
    slideNumber: { x: W - MX - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
  });
  pres.defineSlideMaster({
    title: "SECTION_DARK",
    background: { color: C.text2 },
    objects: [
      ...(LOGO ? [{ image: { data: LOGO, x: W - MX - 0.42, y: 0.42, w: 0.5, h: 0.5 } }] : []),
      { placeholder: { options: { name: "title", type: "title", x: 1.2, y: 3.05, w: 9.5, h: 1.0, fontSize: 44, bold: true, color: C.background1, fontFace: THEME.headFontFace, margin: 0, valign: "top", align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 1.2, y: 4.15, w: 7.3, h: 0.8, fontSize: 18, color: C.accent2, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
    slideNumber: { x: W - MX - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: "9FB3CC", align: "right" },
  });
  pres.defineSlideMaster({ title: "COVER_DARK", background: { color: C.text2 }, objects: [] });

  const sections = { start: "Introduction", who: "Who We Are", gov: "Leadership & Governance", acad: "Academics", ir: "International Relations", fut: "Looking Ahead" };
  let currentSection = null;
  function newSlide(master, sec, id) {
    if (sec !== currentSection) { pres.addSection({ title: sec }); currentSection = sec; }
    const s = pres.addSlide({ masterName: master, sectionTitle: sec });
    const n = process.env.NO_NOTES ? null : NOTES[id];
    if (n) {
      const lines = [n.talk || ""];
      if (n.caveats && n.caveats.length) lines.push("", "CAVEATS:", ...n.caveats.map((c) => "- " + c));
      if (n.likely_question) lines.push("", "LIKELY QUESTION: " + n.likely_question, "ANSWER: " + (n.suggested_answer || ""));
      if (n.sources && n.sources.length) lines.push("", "Sources (briefing): " + n.sources.join(", "));
      s.addNotes(lines.join("\n"));
    }
    return s;
  }
  function content(sec, id, title, kicker) {
    const s = newSlide("CONTENT_LIGHT", sec, id);
    s.addText(title, { placeholder: "title" });
    if (kicker) s.addText(kicker, { placeholder: "kicker" });
    return s;
  }
  function divider(sec, id, num, title, sub) {
    const s = newSlide("SECTION_DARK", sec, id);
    if (LOGO) s.addImage({ data: LOGO, x: 8.75, y: 2.0, w: 4.3, h: 4.3, transparency: 88, objectName: "seal-watermark" });
    s.addText(num, { x: 1.2, y: 2.05, w: 3, h: 0.9, fontSize: 54, bold: true, color: C.accent1, fontFace: THEME.headFontFace, margin: 0, isTextBox: true, objectName: "section-number" });
    s.addText(title, { placeholder: "title" });
    s.addText(sub, { placeholder: "body" });
    return s;
  }
  // icon in a filled circle (motif)
  async function iconCircle(s, name, x, y, d, fill = C.accent1, fg = HEX.white, objName = "icon") {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: objName + "-circle" });
    const p = d * 0.26;
    s.addImage({ data: await icon(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: objName });
  }
  function card(s, x, y, w, h, fill = C.background1, objName = "card") {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: HEX.line, width: 0.75 },
      shadow: { type: "outer", color: "0F172A", opacity: 0.08, blur: 6, offset: 2, angle: 90 }, objectName: objName,
    });
  }
  async function photo(s, key, x, y, w, h, radius = 0.18, objName = "photo") {
    const d = await framed(key, w, h, radius);
    if (d) s.addImage({ data: d, x, y, w, h, objectName: objName, altText: ALT[key] || (SEL[key] && SEL[key].alt) || "East Port Said National University photo" });
    else {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: radius, fill: { color: C.accent4 }, line: { type: "none" }, objectName: objName + "-fallback" });
      s.addImage({ data: await icon("FaBuildingColumns", HEX.blue), x: x + w / 2 - 0.45, y: y + h / 2 - 0.45, w: 0.9, h: 0.9 });
    }
  }
  const T = (text, opts) => ({ text, options: opts || {} });

  // ===== 1. Title =====
  {
    const s = newSlide("COVER_DARK", sections.start, "title");
    const hero = await framed("hero", 6.2, H, 0);
    if (hero) s.addImage({ data: hero, x: W - 6.2, y: 0, w: 6.2, h: H, objectName: "hero-photo" });
    s.addShape(pres.shapes.RECTANGLE, { x: W - 6.2, y: 0, w: 6.2, h: H, fill: { color: C.text2, transparency: 78 }, line: { type: "none" }, objectName: "hero-tint" });
    if (LOGO) s.addImage({ data: LOGO, x: MX + 0.1, y: 0.7, w: 1.25, h: 1.25, objectName: "logo" });
    s.addText("We Live the Future Today", { x: MX + 0.1, y: 2.3, w: 6.3, h: 0.45, fontSize: 18, italic: true, color: C.accent2, margin: 0, isTextBox: true });
    s.addText("East Port Said National University", { x: MX + 0.1, y: 2.8, w: 6.3, h: 1.6, fontSize: 40, bold: true, color: C.background1, fontFace: THEME.headFontFace, margin: 0, valign: "top", isTextBox: true });
    s.addText("Building International Bridges from the Suez Canal", { x: MX + 0.1, y: 4.5, w: 6.3, h: 0.9, fontSize: 20, color: "C9D8EC", margin: 0, valign: "top", isTextBox: true });
    s.addText("International Profile  |  2026", { x: MX + 0.1, y: 6.5, w: 6.3, h: 0.35, fontSize: 12, color: "9FB3CC", margin: 0, isTextBox: true });
  }

  // ===== 2. At a glance =====
  {
    const s = content(sections.who, "glance", "EPNU at a Glance", "A non-profit national university in Sinai, created under Egypt Vision 2030");
    const stats = [["7", "Faculties"], ["13+", "Academic programs"], ["5,200+", "Students"], ["80+", "International publications"]];
    const gap = 0.3, cw = (W - 2 * MX - 3 * gap) / 4;
    stats.forEach(([n, l], i) => {
      const x = MX + i * (cw + gap);
      card(s, x, 1.75, cw, 1.75, C.background1, "stat-card-" + i);
      s.addText(n, { x: x + 0.25, y: 1.95, w: cw - 0.5, h: 0.85, fontSize: 40, bold: true, color: C.accent1, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
      s.addText(l, { x: x + 0.25, y: 2.8, w: cw - 0.5, h: 0.5, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });
    });
    await photo(s, "campus", MX, 3.85, 5.2, 2.9, 0.18, "campus-photo");
    const rows = [
      ["FaLandmark", "Non-profit national (Ahliya) university", "Part of Egypt's new generation of national universities"],
      ["FaLocationDot", "East Port Said, Sinai", "A 44-feddan campus on the Suez Canal corridor, driving regional development"],
      ["FaFlag", "Aligned with Egypt Vision 2030", "A \"fourth-generation\" university focused on research and innovation"],
    ];
    for (let i = 0; i < rows.length; i++) {
      const y = 3.95 + i * 0.95;
      await iconCircle(s, rows[i][0], 6.2, y, 0.62, C.accent1, HEX.white, "glance-icon-" + i);
      s.addText([T(rows[i][1], { bold: true, fontSize: 16, color: C.text2, breakLine: true }), T(rows[i][2], { fontSize: 13, color: C.accent5 })],
        { x: 7.0, y: y - 0.06, w: 5.7, h: 0.8, margin: 0, valign: "middle", isTextBox: true });
    }
  }

  // ===== 3. Why partner =====
  {
    const s = content(sections.who, "why", "Why Partner with EPNU", "A new-generation university building its international partnerships");
    const items = [
      ["FaShip", "Strategic location", "East Port Said on the Suez Canal, where Africa meets Asia, at the heart of Sinai's development."],
      ["FaGraduationCap", "Global academic standards", "English-medium instruction, a credit-hour system with ECTS equivalence, and bilingual graduation certificates."],
      ["FaMicrochip", "Future-focused programs", "Artificial intelligence across faculties, a Huawei ICT Academy, and modern health sciences and clinical pharmacy."],
      ["FaHandshake", "Strong national backing", "A Board chaired by a former Prime Minister, and a national mandate to expand global partnerships."],
    ];
    const cw = (W - 2 * MX - 0.3) / 2, ch = 2.325;
    for (let i = 0; i < 4; i++) {
      const x = MX + (i % 2) * (cw + 0.3), y = 1.75 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch, C.background1, "why-card-" + i);
      await iconCircle(s, items[i][0], x + 0.35, y + 0.35, 0.8, C.accent1, HEX.white, "why-icon-" + i);
      s.addText(items[i][1], { x: x + 1.35, y: y + 0.38, w: cw - 1.65, h: 0.7, fontSize: 20, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, valign: "middle", isTextBox: true });
      s.addText(items[i][2], { x: x + 1.35, y: y + 1.12, w: cw - 1.65, h: 1.1, fontSize: 15, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    }
  }

  // ===== 4. Divider: governance =====
  divider(sections.gov, "sec-gov", "01", "Leadership & Governance", "Experienced leadership with national reach and a global outlook");

  // ===== 5. President =====
  {
    const s = content(sections.gov, "president", "Leadership with a Global Outlook", "Prof. Dr. Atef Alam El-Din, President of East Port Said National University");
    await photo(s, "president", MX, 1.75, 3.6, 4.95, 0.2, "president-photo");
    const rows = [
      ["FaEarthEurope", "Internationally educated", "PhD in Engineering, Leningrad University (former Soviet Union); BSc and MSc, Menoufia University"],
      ["FaHandshakeAngle", "Egyptian-Japanese Education Initiative 2025", "Member of the initiative's Steering Committee; has organized several international scientific conferences"],
      ["FaBuildingColumns", "National policy experience", "Has served on the Egyptian Senate's Committee on Education, Scientific Research & ICT"],
      ["FaAward", "Academic leadership and quality", "Former Acting President of Port Said University; certified NAQAAE quality reviewer; about 40 research papers"],
    ];
    for (let i = 0; i < rows.length; i++) {
      const y = 1.8 + i * 1.22;
      await iconCircle(s, rows[i][0], 4.65, y + 0.05, 0.7, C.accent1, HEX.white, "pres-icon-" + i);
      s.addText([T(rows[i][1], { bold: true, fontSize: 17, color: C.text2, breakLine: true }), T(rows[i][2], { fontSize: 14, color: C.text1 })],
        { x: 5.55, y, w: 7.15, h: 1.05, margin: 0, valign: "top", isTextBox: true });
    }
  }

  // ===== 6. Board =====
  {
    const s = content(sections.gov, "board", "A Board Bridging Government, Academia & Industry", "The Board of Trustees supports the university's strategy and institutional partnerships");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 1.75, w: 3.7, h: 4.95, rectRadius: 0.15, fill: { color: C.text2 }, line: { type: "none" }, objectName: "chair-card" });
    await iconCircle(s, "FaUserTie", MX + 0.35, 2.1, 0.9, C.accent1, HEX.white, "chair-icon");
    s.addText([T("Chairman", { fontSize: 13, color: HEX.cyan, breakLine: true }), T("Eng. Ibrahim Mahlab", { fontSize: 22, bold: true, color: HEX.white, breakLine: true }), T("Former Prime Minister of Egypt", { fontSize: 15, color: "C9D8EC" })],
      { x: MX + 0.35, y: 3.2, w: 3.1, h: 1.6, margin: 0, valign: "top", isTextBox: true });
    s.addText("Members bring together government, universities, medicine, and trade and industry.", { x: MX + 0.35, y: 5.2, w: 3.1, h: 1.2, fontSize: 14, color: "C9D8EC", margin: 0, valign: "top", isTextBox: true });
    const groups = [
      ["FaLandmarkFlag", "Government & public service", ["Maj. Gen. Ahmed Abdullah, former Governor of the Red Sea and Port Said", "Maj. Gen. Ibrahim Abu Limoun, Governor of Port Said"]],
      ["FaGraduationCap", "Academia & medicine", ["Prof. Dr. Atef Alam El-Din, University President", "Prof. Dr. Sherif Youssef Saleh, President of Port Said University", "Prof. Dr. Mohamed Hassan El-Kenawy, former President of Mansoura University", "Prof. Dr. El-Sayed Hussein Hegazy", "Prof. Dr. Mamdouh Radwan El-Nahas, President, Egyptian Society of Diabetes and Lipidology"]],
      ["FaIndustry", "Industry & trade", ["Eng. Lotfy Hassan El-Badrawy, KAPCI Coatings", "Mr. Mohamed El-Said Saada, Port Said Chamber of Commerce", "Dr. Hassan Hassanein Khalil Hassan, pharmaceutical sector"]],
    ];
    const gx = MX + 4.0, gw = W - MX - gx;
    const boxes = [
      [groups[1], gx, 1.75, gw, 2.3],
      [groups[0], gx, 4.25, (gw - 0.2) / 2, 2.45],
      [groups[2], gx + (gw - 0.2) / 2 + 0.2, 4.25, (gw - 0.2) / 2, 2.45],
    ];
    for (let i = 0; i < boxes.length; i++) {
      const [g, x, y, w, h] = boxes[i];
      card(s, x, y, w, h, C.background1, "board-group-" + i);
      await iconCircle(s, g[0], x + 0.25, y + 0.2, 0.55, C.accent1, HEX.white, "board-icon-" + i);
      s.addText(g[1], { x: x + 0.95, y: y + 0.2, w: w - 1.2, h: 0.55, fontSize: 16, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
      s.addText(g[2].map((t, j, a) => T(t, { bullet: true, breakLine: j < a.length - 1 })),
        { x: x + 0.3, y: y + 0.88, w: w - 0.55, h: h - 1.05, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 3, isTextBox: true });
    }
  }

  // ===== 6b. Academic leadership =====
  {
    const s = content(sections.gov, "leadership", "Academic Leadership", "Program directors and university offices leading teaching, quality and international relations");
    const fac = [
      ["FaGears", "Engineering", [["Assoc. Prof. Mohamed El-Sakka", "Mechatronics & Automation"], ["Assoc. Prof. Asmaa Rady", "AI & Machine Learning · Construction"]]],
      ["FaLaptopCode", "Computer Science & IT", [["Prof. Osama El-Baksawy", "Director of the Faculty's Programs"]]],
      ["FaUserDoctor", "Medicine", [["Assoc. Prof. Sherine El-Khouly", "Medicine and Surgery"]]],
      ["FaPrescriptionBottleMedical", "Pharmacy", [["Assoc. Prof. Mohamed Yasser", "Clinical Pharmacy (PharmD)"]]],
      ["FaFlask", "Applied Health Sciences", [["Dr. Salma Khairallah", "Director of the Faculty's Programs"]]],
      ["FaChartLine", "Business Sciences", [["Dr. Ali Gomaa", "Accounting & Software Applications"]]],
    ];
    const gw = 8.35, gap = 0.25, cw = (gw - 2 * gap) / 3, ch = (4.95 - gap) / 2;
    for (let i = 0; i < fac.length; i++) {
      const x = MX + (i % 3) * (cw + gap), y = 1.75 + Math.floor(i / 3) * (ch + gap);
      card(s, x, y, cw, ch, C.background1, "lead-card-" + i);
      await iconCircle(s, fac[i][0], x + 0.22, y + 0.22, 0.55, C.accent1, HEX.white, "lead-icon-" + i);
      s.addText(fac[i][1], { x: x + 0.9, y: y + 0.2, w: cw - 1.05, h: 0.6, fontSize: 14.5, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
      const runs = [];
      fac[i][2].forEach(([n, r], j, a) => {
        runs.push(T(n, { bold: true, fontSize: 15, color: C.accent1, breakLine: true }));
        runs.push(T(r, { fontSize: 12.5, color: C.text1, breakLine: j < a.length - 1, paraSpaceAfter: 6 }));
      });
      s.addText(runs, { x: x + 0.22, y: y + 0.98, w: cw - 0.4, h: ch - 1.1, margin: 0, valign: "top", isTextBox: true });
    }
    const ox = MX + gw + 0.3, ow = W - MX - ox;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ox, y: 1.75, w: ow, h: 4.95, rectRadius: 0.15, fill: { color: C.text2 }, line: { type: "none" }, objectName: "offices-card" });
    s.addText("University Offices", { x: ox + 0.3, y: 1.95, w: ow - 0.6, h: 0.45, fontSize: 17, bold: true, color: HEX.white, margin: 0, isTextBox: true });
    const offices = [
      ["FaEarthAmericas", "International Relations Office", "Assoc. Prof. Mohamed El-Sakka"],
      ["FaAward", "Quality Assurance Center", "Prof. Osama El-Baksawy"],
      ["FaChalkboardUser", "Training & Professional Development", "Assoc. Prof. Asmaa Rady"],
      ["FaBuildingColumns", "Secretary-General", "Ms. Ghada Shazly"],
    ];
    for (let i = 0; i < offices.length; i++) {
      const y = 2.6 + i * 1.0;
      await iconCircle(s, offices[i][0], ox + 0.3, y + 0.08, 0.5, i === 0 ? C.accent2 : C.accent1, i === 0 ? HEX.navy : HEX.white, "office-icon-" + i);
      s.addText([T(offices[i][1], { fontSize: 12, color: HEX.cyan, breakLine: true }), T(offices[i][2], { fontSize: 15, bold: true, color: HEX.white })],
        { x: ox + 0.95, y, w: ow - 1.15, h: 0.8, margin: 0, valign: "middle", isTextBox: true });
    }
  }

  // ===== 7. Divider: academics =====
  divider(sections.acad, "sec-acad", "02", "Academics", "Seven faculties with programs designed for tomorrow's labor market");

  // ===== 8. Faculties =====
  {
    const s = content(sections.acad, "faculties", "Seven Faculties, Future-Ready Programs", "13+ bachelor's programs, with artificial intelligence across engineering, computing and the arts");
    const fac = [
      ["FaGears", "Engineering", "Mechatronics & Automation\nAI & Machine Learning\nConstruction\nEnergy & Electrical Power"],
      ["FaUserDoctor", "Medicine", "Medicine and Surgery: 5 years plus a 2-year internship (ECTS)"],
      ["FaPrescriptionBottleMedical", "Pharmacy", "Clinical Pharmacy (PharmD)"],
      ["FaLaptopCode", "Computer Science & IT", "Computer Science\nArtificial Intelligence\nData Science\nBiomedical Informatics"],
      ["FaFlask", "Applied Health Sciences", "Medical Laboratory Technology\nMedical Imaging & Radiologic Technology"],
      ["FaChartLine", "Business Sciences", "Accounting & Software Applications\nDigital Marketing & E-Management"],
      ["FaLanguage", "Arts", "English Language & Translation\nGeomatics & Artificial Intelligence"],
    ];
    const cw = 2.85, ch = 2.35, gx = 0.25, gy = 0.25;
    for (let i = 0; i < fac.length; i++) {
      const row = i < 4 ? 0 : 1, col = i < 4 ? i : i - 4;
      const x = MX + col * (cw + gx) + (row === 1 ? (cw + gx) / 2 : 0);
      const y = 1.7 + row * (ch + gy);
      card(s, x, y, cw, ch, C.background1, "faculty-card-" + i);
      await iconCircle(s, fac[i][0], x + 0.25, y + 0.25, 0.6, C.accent1, HEX.white, "faculty-icon-" + i);
      s.addText(fac[i][1], { x: x + 0.95, y: y + 0.25, w: cw - 1.1, h: 0.6, fontSize: 16, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
      s.addText(fac[i][2], { x: x + 0.25, y: y + 1.0, w: cw - 0.45, h: ch - 1.15, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    }
  }

  // ===== 8b. One slide per faculty: quick facts + program director(s) =====
  for (const [id, f] of Object.entries(FACPAGES)) {
    const s = content(sections.acad, "fac_" + id, f.title, f.kicker);
    // left: photo + programs
    const lx = MX, lw = 4.25;
    await photo(s, "fac_" + id, lx, 1.75, lw, 2.55, 0.18, "fac-photo");
    card(s, lx, 4.5, lw, 2.2, C.background1, "fac-programs");
    s.addText("Programs", { x: lx + 0.25, y: 4.62, w: lw - 0.5, h: 0.4, fontSize: 14, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    s.addText(f.programs.map((t, j, a) => T(t, { bullet: true, breakLine: j < a.length - 1 })),
      { x: lx + 0.25, y: 5.05, w: lw - 0.5, h: 1.55, fontSize: f.programs.length > 4 ? 12 : 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 3, isTextBox: true });
    // middle: key facts
    const mx = lx + lw + 0.3, mw = 4.0;
    for (let i = 0; i < f.facts.length; i++) {
      const fh = (4.95 - (f.facts.length - 1) * 0.2) / f.facts.length, y = 1.75 + i * (fh + 0.2);
      card(s, mx, y, mw, fh, C.background1, "fac-fact-" + i);
      await iconCircle(s, f.facts[i][0], mx + 0.2, y + 0.2, 0.5, C.accent1, HEX.white, "fac-fact-icon-" + i);
      s.addText([T(f.facts[i][1], { bold: true, fontSize: 14, color: C.text2, breakLine: true }), T(f.facts[i][2], { fontSize: 12, color: C.text1 })],
        { x: mx + 0.85, y: y + 0.12, w: mw - 1.0, h: fh - 0.24, margin: 0, valign: "middle", isTextBox: true });
    }
    // right: director card(s), dark
    const rx = mx + mw + 0.3, rw = W - MX - rx;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx, y: 1.75, w: rw, h: 4.95, rectRadius: 0.15, fill: { color: C.text2 }, line: { type: "none" }, objectName: "director-card" });
    const ds = f.directors || [];
    if (!ds.length) {
      await iconCircle(s, f.icon || "FaBuildingColumns", rx + 0.3, 2.05, 0.7, C.accent1, HEX.white, "dir-icon");
      s.addText([T(f.side_title || "Faculty highlight", { fontSize: 13, color: HEX.cyan, breakLine: true }), T(f.side_text || "", { fontSize: 14, color: HEX.white })],
        { x: rx + 0.3, y: 2.95, w: rw - 0.6, h: 3.5, margin: 0, valign: "top", paraSpaceAfter: 6, isTextBox: true });
    }
    if (ds.length) {
      await iconCircle(s, "FaUserTie", rx + 0.3, 2.0, 0.55, C.accent2, HEX.navy, "dir-icon");
      s.addText(ds.length > 1 ? "Program Directors" : "Program Director", { x: rx + 1.0, y: 2.0, w: rw - 1.2, h: 0.55, fontSize: 15, bold: true, color: HEX.cyan, margin: 0, valign: "middle", isTextBox: true });
    }
    const top = 2.75, avail = 6.6 - top, dh = ds.length ? avail / ds.length : 0;
    for (let i = 0; i < ds.length; i++) {
      const y = top + i * dh, d = ds[i], two = ds.length > 1;
      if (i > 0) s.addShape(pres.shapes.LINE, { x: rx + 0.3, y: y - 0.08, w: rw - 0.6, h: 0, line: { color: "35537F", width: 1 }, objectName: "dir-sep" });
      s.addText([
        T(d.name, { fontSize: two ? 15 : 17, bold: true, color: HEX.white, breakLine: true }),
        T(d.label, { fontSize: two ? 11 : 12, color: HEX.cyan, breakLine: true, paraSpaceAfter: two ? 4 : 8 }),
        T(d.bio, { fontSize: two ? 11 : 13, color: "DCE6F2" }),
      ], { x: rx + 0.3, y, w: rw - 0.6, h: dh - 0.15, margin: 0, valign: "top", isTextBox: true, objectName: "director-" + i });
    }
  }

  // ===== 9. Standards =====
  {
    const s = content(sections.acad, "standards", "Built on International Academic Standards", "Rules and standards that support credit transfer, exchange and joint programs");
    await photo(s, "lab", W - MX - 4.1, 1.75, 4.1, 4.95, 0.2, "lab-photo");
    const items = [
      ["FaComments", "English-medium instruction", "English is the official language of teaching, textbooks and exams"],
      ["FaScaleBalanced", "Credit hours ≈ ECTS", "One semester's load ≈ 30 ECTS (750 hours of student workload)"],
      ["FaFileSignature", "Bilingual certificates", "Graduation certificates issued in Arabic and English"],
      ["FaRightLeft", "Mobility-ready rules", "Non-degree and auditor registration; transfer with course equivalency"],
      ["FaPassport", "International entry routes", "Admission rules accept IGCSE, the Nile International Certificate and foreign certificates"],
      ["FaNetworkWired", "Huawei ICT Academy", "Training in AI, cloud, networks, IoT and modern communications"],
    ];
    const cw = 3.45, ch = 1.5;
    for (let i = 0; i < items.length; i++) {
      const x = MX + (i % 2) * (cw + 0.25), y = 1.75 + Math.floor(i / 2) * (ch + 0.22);
      card(s, x, y, cw, ch, C.background1, "std-card-" + i);
      await iconCircle(s, items[i][0], x + 0.2, y + 0.22, 0.55, C.accent1, HEX.white, "std-icon-" + i);
      s.addText([T(items[i][1], { bold: true, fontSize: 15, color: C.text2, breakLine: true }), T(items[i][2], { fontSize: 12.5, color: C.text1 })],
        { x: x + 0.9, y: y + 0.18, w: cw - 1.05, h: ch - 0.3, margin: 0, valign: "top", isTextBox: true });
    }
  }

  // ===== 9b. Research output =====
  if (RESEARCH) {
    const R = RESEARCH;
    const s = content(sections.acad, "research", R.title, R.kicker);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 1.75, w: 4.1, h: 4.95, rectRadius: 0.15, fill: { color: C.text2 }, line: { type: "none" }, objectName: "research-hero" });
    s.addText(R.big, { x: MX + 0.4, y: 1.95, w: 3.4, h: 1.5, fontSize: 88, bold: true, color: HEX.cyan, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
    s.addText([T(R.big_label, { bold: true, fontSize: 19, color: HEX.white, breakLine: true }), T(R.big_sub, { fontSize: 14, color: "C9D8EC" })],
      { x: MX + 0.4, y: 3.5, w: 3.3, h: 1.6, margin: 0, valign: "top", isTextBox: true });
    if (R.photo !== false) await photo(s, "fac_research", MX + 0.4, 5.2, 3.3, 1.25, 0.1, "research-photo");
    const rx = MX + 4.35, rw = W - MX - rx;
    if (R.chart) {
      card(s, rx, 1.75, rw * 0.5 - 0.11, 4.95, C.background1, "research-chart-card");
      s.addText(R.chart.title, { x: rx + 0.25, y: 1.9, w: rw * 0.5 - 0.6, h: 0.4, fontSize: 14, bold: true, color: C.text2, margin: 0, isTextBox: true });
      s.addChart(pres.charts.BAR, [{ name: R.chart.title, labels: R.chart.labels, values: R.chart.values }], {
        x: rx + 0.15, y: 2.35, w: rw * 0.5 - 0.4, h: 4.2, barDir: R.chart.dir || "col", chartColors: [HEX.blue],
        showValue: true, dataLabelColor: HEX.ink, dataLabelFontSize: 11, catAxisLabelColor: HEX.slate, catAxisLabelFontSize: 11,
        valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, objectName: "research-chart",
      });
    }
    const fx = R.chart ? rx + rw * 0.5 + 0.11 : rx, fw = R.chart ? rw * 0.5 - 0.11 : rw;
    const n = R.facts.length, fh = (4.95 - (n - 1) * 0.2) / n;
    for (let i = 0; i < n; i++) {
      const y = 1.75 + i * (fh + 0.2);
      card(s, fx, y, fw, fh, C.background1, "research-fact-" + i);
      await iconCircle(s, R.facts[i][0], fx + 0.2, y + (fh - 0.5) / 2, 0.5, C.accent1, HEX.white, "research-icon-" + i);
      s.addText([T(R.facts[i][1], { bold: true, fontSize: 14, color: C.text2, breakLine: true }), T(R.facts[i][2], { fontSize: 12, color: C.text1 })],
        { x: fx + 0.85, y: y + 0.08, w: fw - 1.0, h: fh - 0.16, margin: 0, valign: "middle", isTextBox: true });
    }
  }

  // ===== 10. Divider: IR =====
  divider(sections.ir, "sec-ir", "03", "International Relations", "\"Global Openness\" is one of the university's four pillars");

  // ===== 11. Mandate timeline =====
  {
    const s = content(sections.ir, "mandate", "Internationalization Is a National Priority", "Egypt's national universities are being directed to go global, and EPNU is aligned");
    const steps = [
      ["MARCH 2026", "EPNU Academic Affairs Council", "Reviews the Minister's directives, including enhancing international cooperation and academic partnerships"],
      ["JUNE 2026", "National Universities Council", "Priorities: international rankings, global partnerships, international publishing, joint and dual degrees, international students. The Council approves MoUs between national universities and international institutions"],
      ["JULY 2026", "Port Said University Council", "EPNU's President takes part as the council reviews efforts to expand international partnerships and attract international students"],
    ];
    const y0 = 2.35, cw = 3.85, gap = 0.27;
    s.addShape(pres.shapes.LINE, { x: MX + 0.3, y: y0, w: 3 * cw + 2 * gap - 0.6, h: 0, line: { color: HEX.blue, width: 2.5 }, objectName: "timeline-line" });
    steps.forEach(([d, t, b], i) => {
      const x = MX + i * (cw + gap);
      s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y0 - 0.2, w: 0.4, h: 0.4, fill: { color: i === 2 ? C.accent2 : C.accent1 }, line: { color: HEX.white, width: 3 }, objectName: "timeline-dot-" + i });
      s.addText(d, { x, y: 1.6, w: cw, h: 0.4, fontSize: 14, bold: true, color: C.accent1, charSpacing: 1, margin: 0, isTextBox: true });
      card(s, x, 2.85, cw, 2.45, C.background1, "timeline-card-" + i);
      s.addText([T(t, { bold: true, fontSize: 16, color: C.text2, breakLine: true }), T(b, { fontSize: 13.5, color: C.text1 })],
        { x: x + 0.25, y: 3.05, w: cw - 0.5, h: 2.1, margin: 0, valign: "top", paraSpaceAfter: 4, isTextBox: true });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 5.6, w: W - 2 * MX, h: 1.1, rectRadius: 0.12, fill: { color: C.text2 }, line: { type: "none" }, objectName: "goal-banner" });
    await iconCircle(s, "FaQuoteLeft", MX + 0.3, 5.8, 0.7, C.accent1, HEX.white, "quote-icon");
    s.addText([T("Our partnership goal: ", { bold: true, color: HEX.cyan }), T("\"opening pathways for scientific and research exchange with regional and international universities and institutions\"", { color: HEX.white, italic: true })],
      { x: MX + 1.25, y: 5.7, w: W - 2 * MX - 1.5, h: 0.9, fontSize: 16, margin: 0, valign: "middle", isTextBox: true });
  }

  // ===== 12. UPM =====
  {
    const s = content(sections.ir, "upm", "First International Partner: Universiti Putra Malaysia", "A signed partnership agreement that includes a framework for faculty and student exchange");
    const steps = [
      ["FaFileSignature", "Agreement signed", "Partnership with UPM, including faculty and student exchange programs", "done"],
      ["FaVideo", "International webinar", "\"Global Horizons: Cross-Border Innovations in Science and Management\", hosted by UPM-Sarawak", "done"],
      ["FaFlask", "Joint research meeting", "First joint virtual scientific meeting: EPNU staff and students present research", "done"],
      ["FaCalendarCheck", "2nd edition announced", "EPNU announced it would host the next edition of the webinar", "next"],
      ["FaPlaneDeparture", "Next: exchange", "Activating faculty and student mobility under the agreement", "next"],
    ];
    const n = steps.length, gap = 0.2, cw = (W - 2 * MX - 4 * gap) / 5;
    for (let i = 0; i < n; i++) {
      const x = MX + i * (cw + gap), y = 1.75, done = steps[i][3] === "done";
      card(s, x, y, cw, 3.1, done ? C.background1 : C.accent4, "upm-step-" + i);
      await iconCircle(s, steps[i][0], x + 0.25, y + 0.25, 0.65, done ? C.accent1 : C.accent3, HEX.white, "upm-icon-" + i);
      s.addText(String(i + 1).padStart(2, "0"), { x: x + cw - 0.8, y: y + 0.3, w: 0.6, h: 0.5, fontSize: 18, bold: true, color: done ? C.accent1 : C.accent5, align: "right", margin: 0, isTextBox: true });
      s.addText([T(steps[i][1], { bold: true, fontSize: 15, color: C.text2, breakLine: true }), T(steps[i][2], { fontSize: 12.5, color: C.text1 })],
        { x: x + 0.25, y: y + 1.05, w: cw - 0.45, h: 1.95, margin: 0, valign: "top", paraSpaceAfter: 3, isTextBox: true });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 5.1, w: W - 2 * MX, h: 1.6, rectRadius: 0.12, fill: { color: C.text2 }, line: { type: "none" }, objectName: "upm-callout" });
    await iconCircle(s, "FaNewspaper", MX + 0.35, 5.5, 0.8, C.accent1, HEX.white, "press-icon");
    s.addText([T("Recognized abroad", { bold: true, fontSize: 18, color: HEX.cyan, breakLine: true }), T("EPNU pharmacy students' engagement in the UPM webinar was praised by Malaysian press outlets. The partnership is led by the Clinical Pharmacy program.", { fontSize: 14, color: HEX.white })],
      { x: MX + 1.45, y: 5.2, w: W - 2 * MX - 1.8, h: 1.4, margin: 0, valign: "middle", isTextBox: true });
  }

  // ===== 13. Network hub =====
  {
    const s = content(sections.ir, "network", "A Growing Global Network", "International links through partners, forums, initiatives and internationally trained faculty");
    const cx = W / 2, cy = 4.25;
    const nodes = [
      ["Malaysia", "Universiti Putra Malaysia: signed partnership agreement", -3.95, -1.65],
      ["Japan", "Egyptian-Japanese Education Initiative 2025", -4.25, 0.15],
      ["European Union", "Community invited to the Euro-Egyptian Innovation Partnership (Horizon Europe)", -3.95, 1.95],
      ["Russia & Arab World", "5th Forum of the Union of Russian & Arab University Rectors", 3.95, -1.65],
      ["United Kingdom & Italy", "Faculty expertise from Exeter (Newton-Mosharafa scholar) and Politecnico di Torino", 4.25, 0.15],
      ["China & MENA", "Huawei ICT Academy · Saudi German Hospital network", 3.95, 1.95],
    ];
    const nw = 3.5, nh = 1.15;
    nodes.forEach(([t, b, dx, dy], i) => {
      const nx = cx + dx - nw / 2, ny = cy + dy - nh / 2;
      const ex = dx < 0 ? nx + nw : nx;
      const lx = Math.min(cx, ex), lw = Math.abs(ex - cx);
      const ly = Math.min(cy, cy + dy), lh = Math.abs(dy);
      s.addShape(pres.shapes.LINE, { x: lx, y: ly, w: lw, h: lh || 0.001, flipH: (dx < 0) !== (dy < 0) && dy !== 0, line: { color: HEX.line, width: 2 }, objectName: "spoke-" + i });
      card(s, nx, ny, nw, nh, C.background1, "node-" + i);
      s.addText([T(t, { bold: true, fontSize: 15, color: C.accent1, breakLine: true }), T(b, { fontSize: 12, color: C.text1 })],
        { x: nx + 0.2, y: ny + 0.1, w: nw - 0.4, h: nh - 0.2, margin: 0, valign: "middle", isTextBox: true });
    });
    s.addShape(pres.shapes.OVAL, { x: cx - 1.05, y: cy - 1.05, w: 2.1, h: 2.1, fill: { color: C.text2 }, line: { color: HEX.cyan, width: 3 }, objectName: "hub" });
    if (LOGO) s.addImage({ data: LOGO, x: cx - 0.45, y: cy - 0.75, w: 0.9, h: 0.9, objectName: "hub-logo" });
    s.addText("EPNU", { x: cx - 1.0, y: cy + 0.2, w: 2.0, h: 0.45, fontSize: 18, bold: true, color: HEX.white, align: "center", margin: 0, isTextBox: true });
  }

  // ===== 14. Industry & health partners =====
  {
    const s = content(sections.ir, "industry", "Partners with Regional & Global Reach", "Industry and health partners that connect learning to international practice");
    const items = [
      ["FaNetworkWired", "Huawei ICT Academy", "Huawei ICT Academy at EPNU: training in AI, cloud computing, networks, IoT and modern communications", "tech"],
      ["FaHospital", "Saudi German Hospital, Cairo", "Clinical lab training (quality control, biosafety) with a hospital network operating in Saudi Arabia, the UAE, Egypt, Yemen and Morocco", "health"],
      ["FaPills", "Pharmaceutical industry", "Strategic partners include EIPICO, EVA Pharma, Lavender Life and Royal Cosmetics, supporting training, innovation and entrepreneurship", "pharma"],
      ["FaRibbon", "Children's Cancer Hospital 57357", "MoU signed with the 57357 Group: medical education, practical training, research and community service", "h57357"],
    ];
    const gap = 0.25, cw = (W - 2 * MX - 3 * gap) / 4;
    for (let i = 0; i < items.length; i++) {
      const x = MX + i * (cw + gap);
      await photo(s, items[i][3], x, 1.75, cw, 2.2, 0.15, "industry-photo-" + i);
      card(s, x, 4.2, cw, 2.5, C.background1, "industry-card-" + i);
      await iconCircle(s, items[i][0], x + 0.22, 4.4, 0.6, C.accent1, HEX.white, "industry-icon-" + i);
      s.addText(items[i][1], { x: x + 0.95, y: 4.37, w: cw - 1.1, h: 0.66, fontSize: 15, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
      s.addText(items[i][2], { x: x + 0.22, y: 5.15, w: cw - 0.44, h: 1.45, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    }
  }

  // ===== 15. Students =====
  {
    const s = content(sections.ir, "students", "Our Students at International Events", "Competing, publishing and presenting beyond the classroom");
    await photo(s, "students1", MX, 1.75, 4.0, 2.4, 0.15, "students-photo-1");
    await photo(s, "students2", MX, 4.35, 1.9, 2.4, 0.15, "students-photo-2");
    await photo(s, "students3", MX + 2.1, 4.35, 1.9, 2.4, 0.15, "students-photo-3");
    const ach = [
      ["FaRobot", "AI Hackathon, ICIAAI 2026", "Engineering team qualified and competed at the 1st International Conference on Innovations & Applications of AI"],
      ["FaTowerBroadcast", "ITC-Egypt 2026", "Engineering team passed the 2nd judging stage at the 6th International Conference on Communications"],
      ["FaStethoscope", "11th International Pediatric Conference", "Medical students attended, learning the latest advances in pediatric medicine"],
      ["FaTrophy", "Research & programming", "7 papers and 2nd place at Suez Canal University's student research conference; qualified for the Egyptian Collegiate Programming Contest (ECPC)"],
    ];
    for (let i = 0; i < ach.length; i++) {
      const y = 1.75 + i * 1.27;
      card(s, 4.85, y, 7.88, 1.12, C.background1, "ach-card-" + i);
      await iconCircle(s, ach[i][0], 5.05, y + 0.24, 0.64, C.accent1, HEX.white, "ach-icon-" + i);
      s.addText([T(ach[i][1], { bold: true, fontSize: 15, color: C.text2, breakLine: true }), T(ach[i][2], { fontSize: 12.5, color: C.text1 })],
        { x: 5.9, y: y + 0.1, w: 6.65, h: 0.92, margin: 0, valign: "middle", isTextBox: true });
    }
  }

  // ===== 16. National base =====
  {
    const s = content(sections.ir, "national", "A Strong National Partnership Base", "The foundation that supports international collaboration");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 1.75, w: 4.1, h: 4.95, rectRadius: 0.15, fill: { color: C.text2 }, line: { type: "none" }, objectName: "research-stat" });
    s.addText("23", { x: MX + 0.4, y: 2.0, w: 3.3, h: 1.4, fontSize: 80, bold: true, color: HEX.cyan, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
    s.addText([T("research papers in 2024–2025", { bold: true, fontSize: 18, color: HEX.white, breakLine: true }), T("from pharmacy research cooperation with Port Said University, published in international journals including Q1–Q3 titles", { fontSize: 14, color: "C9D8EC" })],
      { x: MX + 0.4, y: 3.4, w: 3.3, h: 1.75, margin: 0, valign: "top", isTextBox: true });
    await photo(s, "national", MX + 0.4, 5.2, 3.3, 1.25, 0.1, "national-photo");
    const partners = [
      ["FaBuildingColumns", "Port Said University", "Main academic partner; joint pharmacy research and shared academic events"],
      ["FaCode", "ITI & Creativa Port Said", "Cooperation talks on digital-skills training for CS students"],
      ["FaAnchor", "Suez Canal Authority Training Center", "Practical engineering workshops"],
      ["FaFileSignature", "Ahl Masr Hospital & BEDO", "Signed protocols: clinical training; industrial training & joint applied research"],
      ["FaKitMedical", "Armed Forces College of Medicine", "Visit to its anatomy hall and clinical simulation center"],
      ["FaCity", "Port Said Governorate & Healthcare Authority", "Community development and health training"],
    ];
    const cw = (W - MX - (MX + 4.35) - 0.22) / 2, ch = 1.5;
    for (let i = 0; i < partners.length; i++) {
      const x = MX + 4.35 + (i % 2) * (cw + 0.22), y = 1.75 + Math.floor(i / 2) * (ch + 0.22);
      card(s, x, y, cw, ch, C.background1, "partner-card-" + i);
      await iconCircle(s, partners[i][0], x + 0.22, y + 0.25, 0.58, C.accent1, HEX.white, "partner-icon-" + i);
      s.addText([T(partners[i][1], { bold: true, fontSize: 14.5, color: C.text2, breakLine: true }), T(partners[i][2], { fontSize: 12.5, color: C.text1 })],
        { x: x + 0.95, y: y + 0.15, w: cw - 1.1, h: ch - 0.3, margin: 0, valign: "middle", isTextBox: true });
    }
  }

  // ===== 17. Opportunities =====
  {
    const s = content(sections.fut, "opportunities", "Partnership Opportunities", "Where we invite international partners to build with us");
    const items = [
      ["FaPeopleArrows", "Student & staff exchange", "Semester exchange and short mobility on our credit-hour/ECTS framework"],
      ["FaUserGraduate", "Joint & dual degrees", "Co-designed programs in AI, health sciences, pharmacy and engineering"],
      ["FaMagnifyingGlassChart", "Joint research & EU funding", "Research consortia, including Horizon Europe calls"],
      ["FaChalkboardUser", "Visiting professors", "Guest teaching, joint supervision and international webinars"],
      ["FaEarthAfrica", "International student pathways", "Pathways for international students, building on foreign-certificate rules"],
      ["FaShip", "Sector hubs on the Suez Canal", "AI, health and maritime/logistics collaboration in a strategic region"],
    ];
    const cw = 3.85, ch = 2.35;
    for (let i = 0; i < items.length; i++) {
      const x = MX + (i % 3) * (cw + 0.27), y = 1.75 + Math.floor(i / 3) * (ch + 0.25);
      card(s, x, y, cw, ch, C.background1, "opp-card-" + i);
      await iconCircle(s, items[i][0], x + 0.3, y + 0.3, 0.75, C.accent1, HEX.white, "opp-icon-" + i);
      s.addText(items[i][1], { x: x + 0.3, y: y + 1.15, w: cw - 0.6, h: 0.45, fontSize: 17, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
      s.addText(items[i][2], { x: x + 0.3, y: y + 1.6, w: cw - 0.6, h: 0.65, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    }
  }

  // ===== 18. Closing =====
  {
    const s = newSlide("COVER_DARK", sections.fut, "closing");
    const img = await framed("closing", W, H, 0);
    if (img) s.addImage({ data: img, x: 0, y: 0, w: W, h: H, objectName: "closing-photo" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.text2, transparency: img ? 18 : 0 }, line: { type: "none" }, objectName: "closing-tint" });
    if (LOGO) s.addImage({ data: LOGO, x: W / 2 - 0.65, y: 1.0, w: 1.3, h: 1.3, objectName: "closing-logo" });
    s.addText("Let's Build the Future Together", { x: 1, y: 2.65, w: W - 2, h: 1.0, fontSize: 42, bold: true, color: HEX.white, fontFace: THEME.headFontFace, align: "center", margin: 0, isTextBox: true });
    s.addText("East Port Said National University welcomes partners for exchange, joint programs and research.", { x: 1.5, y: 3.7, w: W - 3, h: 0.8, fontSize: 18, color: "C9D8EC", align: "center", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: W / 2 - 3.2, y: 4.9, w: 6.4, h: 1.2, rectRadius: 0.15, fill: { color: C.accent3, transparency: 10 }, line: { color: HEX.cyan, width: 1.5 }, objectName: "contact-box" });
    s.addText([T("epnu.edu.eg", { bold: true, fontSize: 22, color: HEX.cyan, breakLine: true, hyperlink: { url: "https://epnu.edu.eg/en/" } }), T("International Relations Office  |  Director: Assoc. Prof. Mohamed El-Sakka", { fontSize: 14, color: HEX.white })],
      { x: W / 2 - 3.0, y: 4.98, w: 6.0, h: 1.05, align: "center", valign: "middle", margin: 0, isTextBox: true });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Wrote " + OUT);
})().catch((e) => { console.error(e); process.exit(1); });

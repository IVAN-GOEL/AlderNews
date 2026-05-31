const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "D4B896" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const SOIL = "6B3D14";
const BARK = "3D2B1F";
const AMBER = "C47C2B";
const GREEN = "3E6B3A";
const CANOPY = "1C3320";
const CREAM = "FBF6EC";
const MUTED = "9C8A74";
const LEAF = "E8C46A";

function h1(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 36, color: CANOPY, font: "Georgia" })],
    spacing: { before: 400, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: AMBER, space: 4 } }
  });
}

function h2(text, color = SOIL) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 26, color, font: "Georgia" })],
    spacing: { before: 320, after: 120 }
  });
}

function h3(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 22, color: AMBER, font: "Arial" })],
    spacing: { before: 200, after: 80 }
  });
}

function body(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, color: "2C1A0E", font: "Arial", ...options })],
    spacing: { before: 60, after: 60 },
    indent: options.indent ? { left: 360 } : {}
  });
}

function bullet(text, color = AMBER) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 22, color: "2C1A0E", font: "Arial" })],
    spacing: { before: 40, after: 40 }
  });
}

function kv(key, value, keyColor = SOIL, valColor = "1C1208") {
  return new Paragraph({
    children: [
      new TextRun({ text: key + ":  ", bold: true, size: 22, color: keyColor, font: "Arial" }),
      new TextRun({ text: value, size: 22, color: valColor, font: "Arial" })
    ],
    spacing: { before: 60, after: 60 },
    indent: { left: 360 }
  });
}

function divider() {
  return new Paragraph({
    children: [new TextRun({ text: "" })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D4B896", space: 2 } },
    spacing: { before: 200, after: 200 }
  });
}

function tag(text, color) {
  return new Paragraph({
    children: [
      new TextRun({ text: `  ${text}  `, size: 18, color: color, bold: true, font: "Arial",
        shading: { fill: color === GREEN ? "EAF2E9" : "FDF3E3", type: ShadingType.CLEAR } })
    ],
    spacing: { before: 40, after: 40 },
    indent: { left: 360 }
  });
}

function colorBox(label, hex, description) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 1200, type: WidthType.DXA },
        shading: { fill: hex, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: " ", size: 20 })] })]
      }),
      new TableCell({
        borders,
        width: { size: 2200, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 120, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, color: BARK, font: "Arial" })] })]
      }),
      new TableCell({
        borders,
        width: { size: 1600, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: `#${hex}`, size: 18, color: MUTED, font: "Courier New" })] })]
      }),
      new TableCell({
        borders,
        width: { size: 4360, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 120, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: description, size: 20, color: "444444", font: "Arial" })] })]
      })
    ]
  });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "–",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 560, hanging: 280 } },
                 run: { color: AMBER, font: "Arial" } }
      }]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [

      // TITLE BLOCK
      new Paragraph({
        children: [new TextRun({ text: "ALDERNEWS", bold: true, size: 64, color: CANOPY, font: "Georgia" })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 80 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Project Summary & Decision Log", size: 28, color: AMBER, font: "Georgia", italics: true })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Prepared: June 2026  ·  Both sides. Always.", size: 20, color: MUTED, font: "Arial" })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 }
      }),

      // ──────────────────────────────
      h1("1. The Concept"),
      h2("What Aldernews is"),
      body("Aldernews is an Indian news startup targeting all Indians as readers, with content created exclusively by verified college students studying journalism, mass media, and anchoring. It operates as a web platform."),
      body("The defining feature of Aldernews is its mandatory 50/50 political feed — every user's feed is algorithmically balanced so that exactly half the content is tagged pro-ruling party and half is tagged pro-opposition. This is non-negotiable and built into the platform's architecture."),

      h2("The core problem it solves"),
      bullet("Indian news media is heavily polarised — outlets align clearly with ruling or opposition positions"),
      bullet("Readers self-select into echo chambers, consuming only news that confirms their existing views"),
      bullet("There is no platform that structurally enforces balanced exposure"),
      bullet("Aldernews solves this by making balance a product feature, not an editorial aspiration"),

      h2("The creator model"),
      bullet("Only verified creators can publish — no anonymous or unverified content"),
      bullet("Creators are college students studying journalism, mass communication, anchoring, and related fields"),
      bullet("Verification is done via student ID and institutional email at the time of creator application"),
      bullet("Creators can publish written articles and/or video reports"),
      bullet("All content is tagged either pro-ruling party or pro-opposition before going live"),

      divider(),

      // ──────────────────────────────
      h1("2. Final Decisions"),

      h2("Platform"),
      kv("Format", "Web platform (not app-first)"),
      kv("Target audience", "All Indians"),
      kv("Content creators", "Verified college students (journalism, mass media, anchoring)"),
      kv("Feed logic", "50% pro-ruling / 50% pro-opposition — always enforced"),
      kv("Non-political content", "Different logic applied (not 50/50 balance)"),
      kv("Feed display", "Stacked feed with colour-coded tags"),
      kv("Layout style", "Clean and minimal — lots of whitespace, text-forward"),

      h2("Brand name"),
      kv("Final name", "Aldernews"),
      kv("Domain target", "aldernews.in"),
      kv("Tagline", "Both sides. Always."),
      kv("Name origin", "The alder tree — in Celtic tradition, the tree of truth and protection"),
      kv("Language", "English"),
      kv("Vibe", "Clean and trustworthy (like The Hindu, Mint)"),
      kv("Name emphasis", "Truth and credibility"),

      h2("Colour palette"),
      body("Final palette: Golden Hour Floor — warm amber, dried fern gold, rich soil. Energetic but earthy."),
      new Paragraph({ children: [new TextRun({ text: "" })], spacing: { before: 120, after: 80 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1200, 2200, 1600, 4360],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: "D4B896", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "Swatch", size: 18, bold: true, color: BARK, font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, shading: { fill: "F0E8D8", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "Name", size: 18, bold: true, color: BARK, font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 1600, type: WidthType.DXA }, shading: { fill: "F0E8D8", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "Hex", size: 18, bold: true, color: BARK, font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, shading: { fill: "F0E8D8", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "Usage", size: 18, bold: true, color: BARK, font: "Arial" })] })] }),
            ]
          }),
          colorBox("Cream", "FBF6EC", "Page backgrounds"),
          colorBox("Fallen leaf", "E8C46A", "Accents, highlights"),
          colorBox("Amber bark", "C47C2B", "Pro-ruling tag, buttons"),
          colorBox("Rich soil", "6B3D14", "Headlines, nav text"),
          colorBox("Evergreen", "3E6B3A", "Pro-opposition tag, CTAs"),
          colorBox("Night canopy", "1C3320", "Navigation bar, dark backgrounds"),
        ]
      }),
      new Paragraph({ children: [new TextRun({ text: "" })], spacing: { before: 80 } }),

      divider(),

      // ──────────────────────────────
      h1("3. Website Structure"),
      body("The Aldernews website is a five-page single-file HTML application. All pages share the Golden Hour colour palette and a consistent navigation bar."),

      h2("Page 1 — Login"),
      bullet("Split-screen layout: brand statement on the left, login form on the right"),
      bullet("Google OAuth (Continue with Google) as the primary sign-in method"),
      bullet("Email + password as a secondary option"),
      bullet("Brand tagline, platform features, and mission statement displayed on left panel"),
      bullet("Link to creator signup from the login footer"),

      h2("Page 2 — Articles"),
      bullet("Sticky navigation bar with night canopy background and fallen-leaf logo"),
      bullet("Live news ticker in amber below the nav"),
      bullet("Page header with today's date and story count"),
      bullet("50/50 balance bar — amber (pro-ruling) and green (pro-opposition) displayed prominently"),
      bullet("Featured article card at full width with image placeholder"),
      bullet("Two-column article grid below the featured story"),
      bullet("Stacked feed of latest stories with colour-coded tags"),
      bullet("Sidebar: balance explainer, trending topics, top creators of the week"),

      h2("Page 3 — Videos"),
      bullet("Same nav and ticker as Articles"),
      bullet("Featured video at full width with play button and duration badge"),
      bullet("Three-column video grid below"),
      bullet("Each video card shows: thumbnail, tag, title, creator name, college, and time ago"),
      bullet("Same pro-ruling / pro-opposition tag system as articles"),

      h2("Page 4 — Global"),
      bullet("Region filter pills: All, South Asia, US & Americas, Europe, Middle East, East Asia, Africa"),
      bullet("Three-column card grid — each card shows country flag, country name, tag, headline, excerpt"),
      bullet("Content is India-contextualised world news (how global events affect India)"),
      bullet("Same tag system — international tag in blue for global stories"),

      h2("Page 5 — Creator Signup"),
      bullet("Hero section with night canopy background, golden headline, three perk cards"),
      bullet("Perks: Publish freely, Verified badge, Real reach"),
      bullet("Application form divided into three sections: Personal, Academic, Content"),
      bullet("Personal: first name, last name, email, phone"),
      bullet("Academic: college name, course (dropdown), year of study (dropdown), student ID"),
      bullet("Content: content type checkboxes, beat/topic checkboxes, personal statement, portfolio link"),
      bullet("Submit button with note that applications are reviewed within 3–5 working days"),

      divider(),

      // ──────────────────────────────
      h1("4. Names Explored & Rejected"),
      body("The following names were considered and rejected before arriving at Aldernews:"),
      new Paragraph({ children: [new TextRun({ text: "" })], spacing: { before: 80 } }),

      h3("Hinglish names (first round)"),
      kv("Dono", "Rejected — user wanted English only"),
      kv("Saath Saath", "Rejected — user wanted English only"),
      kv("Pahlu", "Rejected — user wanted English only"),
      kv("Awaaz Ghar", "Rejected — user wanted English only"),
      kv("Barabar", "Rejected — user wanted English only"),
      kv("Khabar Do", "Rejected — user wanted English only"),

      h3("English names (credibility direction)"),
      kv("Veritas", "Considered but not chosen"),
      kv("The Lens", "Considered but not chosen"),
      kv("Candor", "Considered but not chosen"),
      kv("Plinth", "Considered but not chosen"),
      kv("Groundline", "Liked the feel — but .com taken by a power engineering firm"),
      kv("Meridian", "Considered but not chosen"),

      h3("Groundline-style compound names"),
      kv("Groundmark", "Strong option, not chosen"),
      kv("Earthpress", "Considered but not chosen"),
      kv("Rootmark", "Considered but not chosen"),
      kv("Groundrec", "Considered but not chosen"),
      kv("Bedrock", "Liked — but taken by Amazon AWS product and a US investment firm"),
      kv("Pressroot", "Considered but not chosen"),

      h3("Indian tree names"),
      kv("Banyan", "Strong option, not chosen"),
      kv("Peepal", "Liked — but Peepal Media (Kannada outlet) already exists"),
      kv("Teak / Sagwan", "Considered but not chosen"),
      kv("Neem", "Considered but not chosen"),
      kv("Ashoka", "Considered but not chosen"),

      h3("Global / Greek tree names"),
      kv("Laurel", "Strong option, not chosen"),
      kv("Olive", "Considered but not chosen"),
      kv("Cedar", "Considered but not chosen"),
      kv("Elm", "Considered but not chosen"),
      kv("Daphne", "Considered but not chosen"),
      kv("Arbor", "Considered but not chosen"),
      kv("Linden", "Considered but not chosen"),
      kv("Rowan", "Considered but not chosen"),
      kv("Alder → Aldernews", "CHOSEN — Celtic tree of truth, aldernews.in appears available"),

      divider(),

      // ──────────────────────────────
      h1("5. Palette Options Explored"),
      body("Three forest floor palette options were presented. Golden Hour Floor was selected."),

      h3("Alder Bark (not chosen)"),
      body("Moody and editorial — dark bark browns, soft moss, pale birch parchment. Feels like a broadsheet printed in a forest.", { color: MUTED }),

      h3("Golden Hour Floor (CHOSEN)"),
      body("Autumn leaf fall — warm amber, dried fern gold, rich soil. Energetic but earthy. Best for a warmer, more welcoming feel.", { color: GREEN }),

      h3("Damp Moss (not chosen)"),
      body("After rain — cool greens, grey stone, wet slate. The most trustworthy and serious of the three. Very clean, very credible.", { color: MUTED }),

      divider(),

      // ──────────────────────────────
      h1("6. Key Design Principles"),
      bullet("Balance is a product feature, not an editorial value — it is structurally enforced"),
      bullet("Student journalism is the supply side — quality controlled through a verification process"),
      bullet("The brand must feel credible and trustworthy, not edgy or provocative"),
      bullet("The colour system is semantic: amber always means pro-ruling, green always means pro-opposition"),
      bullet("The website is clean and minimal — whitespace-forward, text-led, no visual clutter"),
      bullet("The nature theme (forest, earth, roots) runs through name, palette, and identity"),
      bullet("The name Alder carries symbolic weight (Celtic tree of truth) without being obscure"),

      divider(),

      // ──────────────────────────────
      h1("7. Immediate Next Steps"),
      bullet("Register aldernews.in on GoDaddy.in or BigRock.in"),
      bullet("Design a logo — likely an alder leaf or stylised tree mark in the Golden Hour palette"),
      bullet("Build the backend: auth, creator verification pipeline, content tagging system"),
      bullet("Set up the 50/50 feed algorithm"),
      bullet("Recruit the first batch of verified student creators"),
      bullet("Plan a soft launch targeting journalism students at IIMC, Symbiosis, Xavier's, and Jadavpur"),

      new Paragraph({
        children: [new TextRun({ text: "aldernews.in  ·  Both sides. Always.", size: 20, color: MUTED, italics: true, font: "Georgia" })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 600, after: 200 }
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/aldernews-summary.docx', buf);
  console.log('Done');
});

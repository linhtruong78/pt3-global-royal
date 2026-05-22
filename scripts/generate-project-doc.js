'use strict';
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat, TabStopType,
} = require('docx');
const fs = require('fs');
const path = require('path');

// ── COLOURS ────────────────────────────────────────────────────
const GOLD       = 'C9A84C';
const BURGUNDY   = '6B1A2A';
const DARK_BG    = '1C0808';   // dark header fill
const IVORY_BG   = 'F9F7F4';   // alternating row fill
const BODY       = '1A1A1A';
const MUTED      = '666666';
const RULE       = 'D4C9B0';   // horizontal rule / border colour
const W          = 9360;       // content width (US Letter 8.5" − 2×1" margins)

// ── BORDER HELPERS ─────────────────────────────────────────────
const b  = (col = RULE) => ({ style: BorderStyle.SINGLE, size: 1, color: col });
const cb = (col = RULE) => ({ top: b(col), bottom: b(col), left: b(col), right: b(col) });

// ── PARAGRAPH HELPERS ──────────────────────────────────────────
const spacer = (after = 240) =>
  new Paragraph({ children: [], spacing: { after } });

const rule = () =>
  new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 4 } },
    spacing: { after: 360 },
  });

const thinRule = () =>
  new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: RULE, space: 2 } },
    spacing: { after: 240 },
  });

const bodyP = (text, opts = {}) =>
  new Paragraph({
    children: [new TextRun({ text, font: 'Arial', size: 22, color: BODY, ...opts })],
    spacing: { after: 160 },
  });

const muted = (text) =>
  new Paragraph({
    children: [new TextRun({ text, font: 'Arial', size: 20, color: MUTED })],
    spacing: { after: 120 },
  });

const bullet = (text) =>
  new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: [new TextRun({ text, font: 'Arial', size: 22, color: BODY })],
    spacing: { after: 100 },
  });

const subBullet = (text) =>
  new Paragraph({
    numbering: { reference: 'subbullets', level: 0 },
    children: [new TextRun({ text, font: 'Arial', size: 20, color: MUTED })],
    spacing: { after: 80 },
  });

// ── TABLE HELPERS ──────────────────────────────────────────────
const thCell = (text, w) =>
  new TableCell({
    borders: cb('A08030'),
    width: { size: w, type: WidthType.DXA },
    shading: { fill: DARK_BG, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: 'Arial', size: 20, bold: true, color: GOLD })],
    })],
  });

const tdCell = (text, w, fill = 'FFFFFF', color = BODY, bold = false) =>
  new TableCell({
    borders: cb(),
    width: { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: 'Arial', size: 20, color, bold })],
    })],
  });

const swatchCell = (hex, w) =>
  new TableCell({
    borders: cb(),
    width: { size: w, type: WidthType.DXA },
    shading: { fill: hex, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({ children: [] })],
  });

// ── STEP TABLE ROW (for setup guide) ───────────────────────────
const stepRow = (num, text, fill = 'FFFFFF') =>
  new TableRow({
    children: [
      tdCell(String(num), 480, fill, GOLD, true),
      tdCell(text, 8880, fill),
    ],
  });

const stepRowSub = (text) =>
  new TableRow({
    children: [
      tdCell('', 480, IVORY_BG),
      tdCell(text, 8880, IVORY_BG, MUTED),
    ],
  });

const stepHeaderRow = (label) =>
  new TableRow({
    children: [
      new TableCell({
        borders: cb('A08030'),
        columnSpan: 2,
        width: { size: W, type: WidthType.DXA },
        shading: { fill: DARK_BG, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [new Paragraph({
          children: [new TextRun({ text: label, font: 'Arial', size: 20, bold: true, color: GOLD })],
        })],
      }),
    ],
  });

// ── DOCUMENT ───────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: 'subbullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '–',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } },
        }],
      },
    ],
  },

  styles: {
    default: {
      document: { run: { font: 'Arial', size: 22, color: BODY } },
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1',
        basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 40, bold: true, font: 'Arial', color: GOLD },
        paragraph: {
          spacing: { before: 480, after: 200 },
          outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 4 } },
        },
      },
      {
        id: 'Heading2', name: 'Heading 2',
        basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: BURGUNDY },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 1 },
      },
      {
        id: 'Heading3', name: 'Heading 3',
        basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: BODY },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 },
      },
    ],
  },

  sections: [
    // ══════════════════════════════════════════════════════════
    // COVER PAGE (separate section — no header/footer)
    // ══════════════════════════════════════════════════════════
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: [
        spacer(2800),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: 'ARAK SAGI SHIRAZ RESERVE WINE & SHAMS BEER INC.',
            font: 'Arial', size: 18, color: 'A08840', characterSpacing: 180,
          })],
          spacing: { after: 480 },
        }),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: 'PT3 Global Royal', font: 'Arial', size: 80, bold: true, color: GOLD,
          })],
          spacing: { after: 320 },
        }),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: 'Website Project Document', font: 'Arial', size: 36, color: BODY,
          })],
          spacing: { after: 1200 },
        }),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [],
          border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: GOLD, space: 4 } },
          spacing: { after: 800 },
        }),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'Version 1.0   |   May 2026', font: 'Arial', size: 20, color: MUTED })],
          spacing: { after: 180 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'Platform: Squarespace 7.1', font: 'Arial', size: 20, color: MUTED })],
          spacing: { after: 180 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'Classification: Confidential', font: 'Arial', size: 20, color: MUTED })],
          spacing: { after: 0 },
        }),

        new Paragraph({ children: [new PageBreak()] }),
      ],
    },

    // ══════════════════════════════════════════════════════════
    // MAIN DOCUMENT
    // ══════════════════════════════════════════════════════════
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1200, right: 1440, bottom: 1200, left: 1440 },
        },
      },

      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [
              new TextRun({ text: 'PT3 Global Royal — Website Project Document', font: 'Arial', size: 18, color: '999999' }),
              new TextRun({ text: '\t', font: 'Arial', size: 18 }),
              new TextRun({ text: 'Confidential', font: 'Arial', size: 18, color: '999999' }),
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: W }],
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: RULE, space: 4 } },
          })],
        }),
      },

      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: 'PT3 Global Royal  |  Crafted for Royal Taste', font: 'Arial', size: 18, color: '999999' }),
              new TextRun({ text: '\t', font: 'Arial', size: 18 }),
              new TextRun({ text: 'Page ', font: 'Arial', size: 18, color: '999999' }),
              new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 18, color: '999999' }),
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: W }],
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: RULE, space: 4 } },
          })],
        }),
      },

      children: [

        // ── 1. PROJECT OVERVIEW ─────────────────────────────
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('1. Project Overview & Brand')] }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('1.1 Brand Summary')] }),
        bodyP('PT3 Global Royal is a luxury heritage beverage brand inspired by Persian royal culture, reserve winemaking traditions, and the art of collectible presentation. The brand encompasses two premium product lines:'),
        bullet('Arak Sagi Shiraz Reserve Wine — a prestige Persian-origin wine presented in collectible royal bottles'),
        bullet('Shams Beer Inc. — a premium craft beer brand drawing from ancient brewing heritage with a modern luxury aesthetic'),
        spacer(160),
        bodyP('The brand positions itself as an ultra-premium, culturally-rooted alternative in the global luxury beverage market, targeting 5-star hospitality, exclusive importers, and high-net-worth collectors across 12+ countries.'),
        spacer(200),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('1.2 Target Audience')] }),
        bullet('Luxury hospitality groups (5-star hotels, private clubs, exclusive resorts)'),
        bullet('Premium retail buyers and specialist wine & spirits merchants'),
        bullet('International importers and distribution partners seeking exclusive territory agreements'),
        bullet('High-net-worth collectors and executive gift buyers'),
        bullet('Press, media, and lifestyle editors covering luxury brands'),
        spacer(200),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('1.3 Project Scope')] }),
        bodyP('This project delivers a complete, production-ready luxury website built for the Squarespace platform, comprising:'),
        bullet('7 fully-designed pages with brand-consistent layout and content'),
        bullet('Custom CSS and code injection system for Squarespace 7.1 integration'),
        bullet('6 optimised photography assets sourced, resized, and compressed for the web'),
        bullet('GSAP ScrollTrigger animation system with hero entrance and scroll-triggered effects'),
        bullet('Fully responsive mobile layout including a full-screen mobile navigation menu'),
        bullet('Contact form with 6 inquiry types and animated success state'),
        spacer(320),

        // ── 2. BRAND IDENTITY ───────────────────────────────
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('2. Brand Identity')] }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('2.1 Color Palette')] }),
        spacer(80),

        // Color palette — swatch(720) name(1800) hex(1440) var(2400) usage(3000) = 9360
        new Table({
          width: { size: W, type: WidthType.DXA },
          columnWidths: [720, 1800, 1440, 2400, 3000],
          rows: [
            new TableRow({ children: [
              thCell('Swatch', 720), thCell('Name', 1800), thCell('Hex', 1440),
              thCell('CSS Variable', 2400), thCell('Usage', 3000),
            ]}),
            new TableRow({ children: [
              swatchCell('C9A84C', 720),
              tdCell('Gold', 1800, IVORY_BG, BODY, true), tdCell('#C9A84C', 1440, IVORY_BG),
              tdCell('--pt3-gold', 2400, IVORY_BG, MUTED), tdCell('Headings, borders, CTAs, accents', 3000, IVORY_BG),
            ]}),
            new TableRow({ children: [
              swatchCell('E2C97E', 720),
              tdCell('Gold Light', 1800, 'FFFFFF', BODY, true), tdCell('#E2C97E', 1440),
              tdCell('--pt3-gold-light', 2400, 'FFFFFF', MUTED), tdCell('Hover states, secondary highlights', 3000),
            ]}),
            new TableRow({ children: [
              swatchCell('6B1A2A', 720),
              tdCell('Burgundy', 1800, IVORY_BG, BODY, true), tdCell('#6B1A2A', 1440, IVORY_BG),
              tdCell('--pt3-burgundy', 2400, IVORY_BG, MUTED), tdCell('Reserve Wines accent, brand depth', 3000, IVORY_BG),
            ]}),
            new TableRow({ children: [
              swatchCell('8B2635', 720),
              tdCell('Wine Red', 1800, 'FFFFFF', BODY, true), tdCell('#8B2635', 1440),
              tdCell('--pt3-wine-red', 2400, 'FFFFFF', MUTED), tdCell('Wine product highlights', 3000),
            ]}),
            new TableRow({ children: [
              swatchCell('F0E8D8', 720),
              tdCell('Ivory', 1800, IVORY_BG, BODY, true), tdCell('#F0E8D8', 1440, IVORY_BG),
              tdCell('--pt3-ivory', 2400, IVORY_BG, MUTED), tdCell('Primary body text on dark backgrounds', 3000, IVORY_BG),
            ]}),
            new TableRow({ children: [
              swatchCell('0E0E0E', 720),
              tdCell('Black', 1800, 'FFFFFF', BODY, true), tdCell('#0E0E0E', 1440),
              tdCell('--pt3-black', 2400, 'FFFFFF', MUTED), tdCell('Page & hero section backgrounds', 3000),
            ]}),
            new TableRow({ children: [
              swatchCell('1A1A1A', 720),
              tdCell('Black Soft', 1800, IVORY_BG, BODY, true), tdCell('#1A1A1A', 1440, IVORY_BG),
              tdCell('--pt3-black-soft', 2400, IVORY_BG, MUTED), tdCell('Alternate section backgrounds', 3000, IVORY_BG),
            ]}),
          ],
        }),
        spacer(320),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('2.2 Typography')] }),
        spacer(80),

        // Typography — role(1800) font(2400) weight(1560) usage(3600) = 9360
        new Table({
          width: { size: W, type: WidthType.DXA },
          columnWidths: [1800, 2400, 1560, 3600],
          rows: [
            new TableRow({ children: [
              thCell('Role', 1800), thCell('Font Family', 2400),
              thCell('Weights Used', 1560), thCell('Usage', 3600),
            ]}),
            new TableRow({ children: [
              tdCell('Primary Serif', 1800, IVORY_BG, BODY, true),
              tdCell('Cormorant Garamond', 2400, IVORY_BG, BURGUNDY, true),
              tdCell('300, 400, 600', 1560, IVORY_BG),
              tdCell('Hero titles, headlines, pull quotes, brand name', 3600, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('Secondary Sans', 1800, 'FFFFFF', BODY, true),
              tdCell('Montserrat', 2400, 'FFFFFF', BODY, true),
              tdCell('300, 400, 500, 600', 1560),
              tdCell('Body text, labels, nav links, eyebrow caps', 3600),
            ]}),
            new TableRow({ children: [
              tdCell('Body Text', 1800, IVORY_BG),
              tdCell('Montserrat 0.75–0.85rem', 2400, IVORY_BG),
              tdCell('300 (Light)', 1560, IVORY_BG),
              tdCell('Paragraph text, card descriptions, footer', 3600, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('Eyebrow Caps', 1800),
              tdCell('Montserrat 0.55–0.65rem', 2400),
              tdCell('500, spacing 0.25em', 1560),
              tdCell('Section labels above headings — ALL CAPS in gold', 3600),
            ]}),
          ],
        }),
        spacer(320),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('2.3 Tone of Voice')] }),
        bodyP('The PT3 Global Royal brand voice is authoritative, refined, and heritage-focused:'),
        bullet('Elevated — Language mirrors the luxury positioning. Never casual or colloquial.'),
        bullet('Heritage-forward — References to Persian royal culture, ancient traditions, and legacy throughout.'),
        bullet('Confident, not boastful — Quiet authority rather than aggressive claims.'),
        bullet('Sensory — Descriptions evoke taste, aroma, colour, and craftsmanship.'),
        bullet('Exclusive — Language implies privilege: "Reserve", "Royal", "Collector\'s Edition".'),
        spacer(320),

        // ── 3. SITE ARCHITECTURE ────────────────────────────
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('3. Site Architecture & Pages')] }),
        bodyP('The website comprises 7 pages structured around the brand\'s narrative journey — from heritage through product collections to global distribution and contact.'),
        spacer(160),

        // Site map — #(480) page(1680) slug(1680) purpose(2520) key sections(3000) = 9360
        new Table({
          width: { size: W, type: WidthType.DXA },
          columnWidths: [480, 1680, 1680, 2520, 3000],
          rows: [
            new TableRow({ children: [
              thCell('#', 480), thCell('Page', 1680), thCell('URL Slug', 1680),
              thCell('Purpose', 2520), thCell('Key Sections', 3000),
            ]}),
            new TableRow({ children: [
              tdCell('1', 480, IVORY_BG, GOLD, true),
              tdCell('Home', 1680, IVORY_BG, BODY, true),
              tdCell('/', 1680, IVORY_BG, MUTED),
              tdCell('Brand gateway — first impression and navigation hub', 2520, IVORY_BG),
              tdCell('Hero, Brand Statement, Collections Grid, Heritage Teaser, Global Presence', 3000, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('2', 480, 'FFFFFF', GOLD, true),
              tdCell('Heritage', 1680, 'FFFFFF', BODY, true),
              tdCell('/heritage', 1680, 'FFFFFF', MUTED),
              tdCell('Tell the Persian royal heritage story and brand values', 2520),
              tdCell('Hero, Origin Story, Timeline (6 milestones), Values Grid, Pull Quote, CTA', 3000),
            ]}),
            new TableRow({ children: [
              tdCell('3', 480, IVORY_BG, GOLD, true),
              tdCell('Collections', 1680, IVORY_BG, BODY, true),
              tdCell('/collections', 1680, IVORY_BG, MUTED),
              tdCell('Overview of both product lines with links to detail pages', 2520, IVORY_BG),
              tdCell('Hero, Intro, Reserve Wines Feature, Shams Beer Feature, Packaging Grid, CTA', 3000, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('4', 480, 'FFFFFF', GOLD, true),
              tdCell('Reserve Wines', 1680, 'FFFFFF', BODY, true),
              tdCell('/reserve-wines', 1680, 'FFFFFF', MUTED),
              tdCell('Deep-dive into the Arak Sagi Shiraz Reserve Wine product line', 2520),
              tdCell('Split Hero, Philosophy, Wine Cards (3 SKUs), Tasting Profile, Food Pairings, CTA', 3000),
            ]}),
            new TableRow({ children: [
              tdCell('5', 480, IVORY_BG, GOLD, true),
              tdCell('Shams Beer', 1680, IVORY_BG, BODY, true),
              tdCell('/shams-beer', 1680, IVORY_BG, MUTED),
              tdCell('Deep-dive into the Shams Beer Inc. premium beer range', 2520, IVORY_BG),
              tdCell('Split Hero, Brand Story, Stats Band, Beer Cards (3 SKUs), Features, Occasions, CTA', 3000, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('6', 480, 'FFFFFF', GOLD, true),
              tdCell('Distribution', 1680, 'FFFFFF', BODY, true),
              tdCell('/distribution', 1680, 'FFFFFF', MUTED),
              tdCell('Global distribution network, regional coverage, partnership process', 2520),
              tdCell('Hero, Stats (4 KPIs), Region Cards (4), Partner Network (3), Process Steps, CTA', 3000),
            ]}),
            new TableRow({ children: [
              tdCell('7', 480, IVORY_BG, GOLD, true),
              tdCell('Contact', 1680, IVORY_BG, BODY, true),
              tdCell('/contact', 1680, IVORY_BG, MUTED),
              tdCell('Inquiry form and direct contact for all business types', 2520, IVORY_BG),
              tdCell('Hero, Contact Form (6 inquiry types), Info Panel (sticky), Social Links', 3000, IVORY_BG),
            ]}),
          ],
        }),
        spacer(320),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('3.1 Navigation Structure')] }),
        bodyP('All pages share a fixed navbar and footer injected globally via Squarespace Code Injection (Footer):'),
        bullet('Navbar links: Heritage · Collections · Reserve Wines · Shams Beer · Distribution'),
        bullet('Navbar CTA: "Inquire" → links to /contact'),
        bullet('Mobile: hamburger icon triggers full-screen overlay menu'),
        bullet('Footer columns: Navigate · Company · Contact'),
        spacer(320),

        // ── 4. SQUARESPACE SETUP ─────────────────────────────
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('4. Squarespace Setup Guide')] }),
        bodyP('Follow these steps in order to implement the PT3 Global Royal website on Squarespace 7.1.'),
        spacer(160),

        // Single table for all setup steps — step#(480) instruction(8880) = 9360
        new Table({
          width: { size: W, type: WidthType.DXA },
          columnWidths: [480, 8880],
          rows: [
            // 4.1
            stepHeaderRow('Phase 1 — Site Creation'),
            stepRow(1, 'Go to squarespace.com and click "Get Started" to create a new site.', IVORY_BG),
            stepRowSub('Choose any blank template — all visual design will be overridden by custom code.'),
            stepRow(2, 'In the Pages panel, create 7 pages with the exact URL slugs shown in Section 3.', IVORY_BG),
            stepRowSub('Home is created automatically at /'),
            stepRowSub('For each page: Pages → + → Blank Page → set the Page URL slug'),
            // 4.2
            stepHeaderRow('Phase 2 — Upload Image Assets'),
            stepRow(3, 'Go to the Squarespace Assets Library (Media panel).', IVORY_BG),
            stepRow(4, 'Upload all 7 image files listed in Section 5.2 of this document.', IVORY_BG),
            stepRow(5, 'After each upload, copy the Squarespace CDN URL (format: https://images.squarespace-cdn.com/content/…).', IVORY_BG),
            stepRowSub('Keep a note of each filename and its CDN URL — you will need these in Phase 5.'),
            // 4.3
            stepHeaderRow('Phase 3 — Apply Global Styles'),
            stepRow(6, 'Open 01-custom-css.css from the squarespace/ folder in a text editor.', IVORY_BG),
            stepRow(7, 'In Squarespace: Design → Custom CSS.', IVORY_BG),
            stepRow(8, 'Paste the entire file contents into the Custom CSS editor and click Save.', IVORY_BG),
            // 4.4
            stepHeaderRow('Phase 4 — Apply Code Injections'),
            stepRow(9, 'In Squarespace: Settings → Advanced → Code Injection.', IVORY_BG),
            stepRow(10, 'Header field: paste the contents of 02-code-injection-header.html.', IVORY_BG),
            stepRowSub('This loads Google Fonts, Tailwind CSS config, and GSAP animation libraries.'),
            stepRow(11, 'Open 03-code-injection-footer.html. Find the text [PASTE LOGO URL HERE].', IVORY_BG),
            stepRowSub('Replace [PASTE LOGO URL HERE] with the CDN URL of the uploaded logo-v1-transparent.png'),
            stepRow(12, 'Paste the updated footer code into the Footer field in Code Injection and click Save.', IVORY_BG),
            // 4.5
            stepHeaderRow('Phase 5 — Add Page Code Blocks'),
            stepRow(13, 'For each of the 7 pages, open its .html file from the squarespace/ folder.', IVORY_BG),
            stepRow(14, 'Find all [IMAGE: filename] placeholders (see Section 5.3) and replace with the CDN URLs from Phase 2.', IVORY_BG),
            stepRow(15, 'In Squarespace, open the page for editing. Click + to add a block → select Code → switch to HTML mode.', IVORY_BG),
            stepRow(16, 'Paste the updated file contents into the code block and click Apply, then Save.', IVORY_BG),
            stepRowSub('You will see a warning: "embedded scripts are disabled while editing." This is normal. Scripts work for site visitors and in incognito windows.'),
            // 4.6
            stepHeaderRow('Phase 6 — Section Width & Final Checks'),
            stepRow(17, 'If any section does not span the full page width: hover over the section → click pencil icon → set Content Width to Full, Left/Right Padding to 0.', IVORY_BG),
            stepRow(18, 'Open an incognito browser window and visit your Squarespace site URL.', IVORY_BG),
            stepRow(19, 'Verify GSAP animations, mobile menu, all navigation links, and the contact form success state.', IVORY_BG),
          ],
        }),
        spacer(320),

        // ── 5. FILE INVENTORY ───────────────────────────────
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('5. File & Asset Inventory')] }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('5.1 Squarespace Code Files')] }),
        bodyP('All code files are located in the squarespace/ folder of the project directory.'),
        spacer(80),

        // Code files — filename(2800) destination(3000) description(3560) = 9360
        new Table({
          width: { size: W, type: WidthType.DXA },
          columnWidths: [2800, 3000, 3560],
          rows: [
            new TableRow({ children: [
              thCell('Filename', 2800), thCell('Squarespace Destination', 3000), thCell('Description', 3560),
            ]}),
            new TableRow({ children: [
              tdCell('01-custom-css.css', 2800, IVORY_BG, BODY, true),
              tdCell('Design → Custom CSS', 3000, IVORY_BG, MUTED),
              tdCell('All global brand styles, Squarespace resets, typography, animations, responsive breakpoints', 3560, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('02-code-injection-header.html', 2800, 'FFFFFF', BODY, true),
              tdCell('Settings → Advanced → Code Injection → Header', 3000, 'FFFFFF', MUTED),
              tdCell('Google Fonts, Tailwind CSS config, GSAP + ScrollTrigger CDN loads', 3560),
            ]}),
            new TableRow({ children: [
              tdCell('03-code-injection-footer.html', 2800, IVORY_BG, BODY, true),
              tdCell('Settings → Advanced → Code Injection → Footer', 3000, IVORY_BG, MUTED),
              tdCell('Shared navbar HTML, mobile menu overlay, shared footer HTML, global scroll-animation JS', 3560, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('homepage.html', 2800, 'FFFFFF', BODY, true),
              tdCell('Home page → Edit → Code Block', 3000, 'FFFFFF', MUTED),
              tdCell('Hero, brand statement, collections grid, heritage teaser, global presence stats', 3560),
            ]}),
            new TableRow({ children: [
              tdCell('heritage.html', 2800, IVORY_BG, BODY, true),
              tdCell('Heritage page → Edit → Code Block', 3000, IVORY_BG, MUTED),
              tdCell('Hero, origin story, 6-milestone timeline, values grid, pull quote, CTA', 3560, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('collections.html', 2800, 'FFFFFF', BODY, true),
              tdCell('Collections page → Edit → Code Block', 3000, 'FFFFFF', MUTED),
              tdCell('Hero, wine feature, beer feature, packaging grid, CTA', 3560),
            ]}),
            new TableRow({ children: [
              tdCell('reserve-wines.html', 2800, IVORY_BG, BODY, true),
              tdCell('Reserve Wines page → Edit → Code Block', 3000, IVORY_BG, MUTED),
              tdCell('Split hero, 3 wine cards (Grand/Royal/Collector), tasting profile, food pairings, CTA', 3560, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('shams-beer.html', 2800, 'FFFFFF', BODY, true),
              tdCell('Shams Beer page → Edit → Code Block', 3000, 'FFFFFF', MUTED),
              tdCell('Split hero, 3 beer cards (Gold/Blanc/Noir), features grid, occasions grid, CTA', 3560),
            ]}),
            new TableRow({ children: [
              tdCell('distribution.html', 2800, IVORY_BG, BODY, true),
              tdCell('Distribution page → Edit → Code Block', 3000, IVORY_BG, MUTED),
              tdCell('Hero, 4 region cards (ME/EU/NA/APAC), partner network (3), 4-step process, CTA', 3560, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('contact.html', 2800, 'FFFFFF', BODY, true),
              tdCell('Contact page → Edit → Code Block', 3000, 'FFFFFF', MUTED),
              tdCell('Hero, inquiry form (6 types + success state), sticky info panel, social links, brand band', 3560),
            ]}),
          ],
        }),
        spacer(320),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('5.2 Image Assets')] }),
        bodyP('All image files are located in PT3/assets/images/. Upload all files to Squarespace Assets Library before implementing page code blocks.'),
        spacer(80),

        // Images — filename(2400) dimensions(1560) size(1200) format(800) used on(3400) = 9360
        new Table({
          width: { size: W, type: WidthType.DXA },
          columnWidths: [2400, 1560, 1200, 800, 3400],
          rows: [
            new TableRow({ children: [
              thCell('Filename', 2400), thCell('Dimensions', 1560), thCell('File Size', 1200),
              thCell('Format', 800), thCell('Used On', 3400),
            ]}),
            new TableRow({ children: [
              tdCell('logo-v1-transparent.png', 2400, IVORY_BG, BODY, true),
              tdCell('Original', 1560, IVORY_BG, MUTED), tdCell('—', 1200, IVORY_BG, MUTED),
              tdCell('PNG', 800, IVORY_BG, MUTED),
              tdCell('All pages via navbar (Code Injection Footer) + Contact info panel', 3400, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('hero-bg.jpg', 2400, 'FFFFFF', BODY, true),
              tdCell('1800 × 1000', 1560, 'FFFFFF', MUTED), tdCell('84 KB', 1200, 'FFFFFF', MUTED),
              tdCell('JPEG', 800, 'FFFFFF', MUTED),
              tdCell('Homepage hero background (luminosity blend at 18% opacity)', 3400),
            ]}),
            new TableRow({ children: [
              tdCell('wine-collection.jpg', 2400, IVORY_BG, BODY, true),
              tdCell('900 × 1200', 1560, IVORY_BG, MUTED), tdCell('55 KB', 1200, IVORY_BG, MUTED),
              tdCell('JPEG', 800, IVORY_BG, MUTED),
              tdCell('Homepage collections grid — Reserve Wines card', 3400, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('beer-collection.jpg', 2400, 'FFFFFF', BODY, true),
              tdCell('900 × 1200', 1560, 'FFFFFF', MUTED), tdCell('53 KB', 1200, 'FFFFFF', MUTED),
              tdCell('JPEG', 800, 'FFFFFF', MUTED),
              tdCell('Homepage collections grid — Shams Beer card', 3400),
            ]}),
            new TableRow({ children: [
              tdCell('heritage-visual.jpg', 2400, IVORY_BG, BODY, true),
              tdCell('900 × 1100', 1560, IVORY_BG, MUTED), tdCell('123 KB', 1200, IVORY_BG, MUTED),
              tdCell('JPEG', 800, IVORY_BG, MUTED),
              tdCell('Homepage heritage teaser + Heritage page origin story panel', 3400, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('wine-hero.jpg', 2400, 'FFFFFF', BODY, true),
              tdCell('1600 × 900', 1560, 'FFFFFF', MUTED), tdCell('76 KB', 1200, 'FFFFFF', MUTED),
              tdCell('JPEG', 800, 'FFFFFF', MUTED),
              tdCell('Reserve Wines page — hero background (luminosity blend)', 3400),
            ]}),
            new TableRow({ children: [
              tdCell('beer-hero.jpg', 2400, IVORY_BG, BODY, true),
              tdCell('1600 × 900', 1560, IVORY_BG, MUTED), tdCell('89 KB', 1200, IVORY_BG, MUTED),
              tdCell('JPEG', 800, IVORY_BG, MUTED),
              tdCell('Shams Beer page — hero background (luminosity blend)', 3400, IVORY_BG),
            ]}),
          ],
        }),
        spacer(320),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('5.3 Image Placeholder Reference')] }),
        bodyP('Each page code file uses placeholder text in the format [IMAGE: filename]. Replace each placeholder with the corresponding Squarespace CDN URL after uploading assets.'),
        spacer(80),

        // Placeholders — placeholder(3200) file(1760) replace in(2400) notes(2000) = 9360
        new Table({
          width: { size: W, type: WidthType.DXA },
          columnWidths: [3200, 1760, 2400, 2000],
          rows: [
            new TableRow({ children: [
              thCell('Placeholder Text', 3200), thCell('Image File', 1760),
              thCell('Appears In', 2400), thCell('Notes', 2000),
            ]}),
            new TableRow({ children: [
              tdCell('[PASTE LOGO URL HERE]', 3200, IVORY_BG),
              tdCell('logo-v1-transparent.png', 1760, IVORY_BG),
              tdCell('03-code-injection-footer.html', 2400, IVORY_BG, MUTED),
              tdCell('Navbar logo <img> src', 2000, IVORY_BG, MUTED),
            ]}),
            new TableRow({ children: [
              tdCell('[IMAGE: logo-v1-transparent.png]', 3200),
              tdCell('logo-v1-transparent.png', 1760),
              tdCell('homepage.html, heritage.html, contact.html', 2400, 'FFFFFF', MUTED),
              tdCell('Logo in content sections', 2000, 'FFFFFF', MUTED),
            ]}),
            new TableRow({ children: [
              tdCell('[IMAGE: hero-bg.jpg]', 3200, IVORY_BG),
              tdCell('hero-bg.jpg', 1760, IVORY_BG),
              tdCell('homepage.html', 2400, IVORY_BG, MUTED),
              tdCell('Hero <img> src', 2000, IVORY_BG, MUTED),
            ]}),
            new TableRow({ children: [
              tdCell('[IMAGE: wine-collection.jpg]', 3200),
              tdCell('wine-collection.jpg', 1760),
              tdCell('homepage.html', 2400, 'FFFFFF', MUTED),
              tdCell('Collections grid wine card', 2000, 'FFFFFF', MUTED),
            ]}),
            new TableRow({ children: [
              tdCell('[IMAGE: beer-collection.jpg]', 3200, IVORY_BG),
              tdCell('beer-collection.jpg', 1760, IVORY_BG),
              tdCell('homepage.html', 2400, IVORY_BG, MUTED),
              tdCell('Collections grid beer card', 2000, IVORY_BG, MUTED),
            ]}),
            new TableRow({ children: [
              tdCell('[IMAGE: heritage-visual.jpg]', 3200),
              tdCell('heritage-visual.jpg', 1760),
              tdCell('homepage.html, heritage.html', 2400, 'FFFFFF', MUTED),
              tdCell('Heritage teaser & origin story', 2000, 'FFFFFF', MUTED),
            ]}),
            new TableRow({ children: [
              tdCell('[IMAGE: wine-hero.jpg]', 3200, IVORY_BG),
              tdCell('wine-hero.jpg', 1760, IVORY_BG),
              tdCell('reserve-wines.html', 2400, IVORY_BG, MUTED),
              tdCell('Reserve Wines hero', 2000, IVORY_BG, MUTED),
            ]}),
            new TableRow({ children: [
              tdCell('[IMAGE: beer-hero.jpg]', 3200),
              tdCell('beer-hero.jpg', 1760),
              tdCell('shams-beer.html', 2400, 'FFFFFF', MUTED),
              tdCell('Shams Beer hero', 2000, 'FFFFFF', MUTED),
            ]}),
          ],
        }),
        spacer(320),

        // ── 6. TECHNICAL STACK ──────────────────────────────
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('6. Technical Stack')] }),
        spacer(80),

        // Tech — technology(2800) version(1360) role(5200) = 9360
        new Table({
          width: { size: W, type: WidthType.DXA },
          columnWidths: [2800, 1360, 5200],
          rows: [
            new TableRow({ children: [
              thCell('Technology', 2800), thCell('Version', 1360), thCell('Role', 5200),
            ]}),
            new TableRow({ children: [
              tdCell('Squarespace', 2800, IVORY_BG, BODY, true), tdCell('7.1', 1360, IVORY_BG, MUTED),
              tdCell('Hosting platform, CMS, DNS, SSL, global CDN', 5200, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('GSAP', 2800, 'FFFFFF', BODY, true), tdCell('3.12.5', 1360, 'FFFFFF', MUTED),
              tdCell('Hero entrance animations, scroll-triggered fade-up/fade-in, divider line draws, mobile menu', 5200),
            ]}),
            new TableRow({ children: [
              tdCell('GSAP ScrollTrigger', 2800, IVORY_BG, BODY, true), tdCell('3.12.5', 1360, IVORY_BG, MUTED),
              tdCell('Scroll-based animation triggers throughout all pages', 5200, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('Tailwind CSS', 2800, 'FFFFFF', BODY, true), tdCell('CDN', 1360, 'FFFFFF', MUTED),
              tdCell('Utility class support and responsive grid helpers', 5200),
            ]}),
            new TableRow({ children: [
              tdCell('Cormorant Garamond', 2800, IVORY_BG, BODY, true), tdCell('Google Fonts', 1360, IVORY_BG, MUTED),
              tdCell('Primary serif typeface — all headlines, hero titles, brand name', 5200, IVORY_BG),
            ]}),
            new TableRow({ children: [
              tdCell('Montserrat', 2800, 'FFFFFF', BODY, true), tdCell('Google Fonts', 1360, 'FFFFFF', MUTED),
              tdCell('Secondary sans-serif — body text, labels, navigation links, eyebrow caps', 5200),
            ]}),
            new TableRow({ children: [
              tdCell('sharp (Node.js)', 2800, IVORY_BG, BODY, true), tdCell('npm', 1360, IVORY_BG, MUTED),
              tdCell('Image processing pipeline: download, resize, and compress assets to mozjpeg for web optimisation', 5200, IVORY_BG),
            ]}),
          ],
        }),
        spacer(480),

        // ── CLOSING BAND ────────────────────────────────────
        thinRule(),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'PT3 Global Royal  —  Crafted for Royal Taste', font: 'Arial', size: 22, color: GOLD, bold: true })],
          spacing: { after: 120 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'This document is confidential and intended for authorised recipients only.', font: 'Arial', size: 18, color: MUTED })],
          spacing: { after: 0 },
        }),

      ], // end children
    }, // end main section
  ], // end sections
}); // end Document

// ── OUTPUT ─────────────────────────────────────────────────────
const OUT = path.join('C:\\Users\\truon\\OneDrive\\Desktop\\PT3', 'PT3-Global-Royal-Project-Document.docx');
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('Document saved to: ' + OUT);
}).catch(err => {
  console.error('Error:', err.message || err);
  process.exit(1);
});

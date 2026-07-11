/**
 * Generates templates/cv_template.docx
 * Run: npm run build:templates  (or  node scripts/build-cv-template.js)
 *
 * What is FIXED in the template (never changes per JD):
 *   - Name, contact info, links
 *   - Role titles, company names, locations, dates for all 4 roles
 *   - Education, Projects, Certifications, Languages
 *
 * What is DYNAMIC (filled at generation time):
 *   - {headline}            — subtitle line under name
 *   - {summary}             — professional summary paragraph
 *   - {#skills}{label}{value}{/skills}  — skills table rows (order & content vary per JD)
 *   - Per-role bullet loops (4 roles, each with its own loop tag):
 *       {#calvergy_bullets}{text}{/calvergy_bullets}
 *       {#senior_baris_bullets}{text}{/senior_baris_bullets}
 *       {#developer_baris_bullets}{text}{/developer_baris_bullets}
 *       {#junior_baris_bullets}{text}{/junior_baris_bullets}
 */
const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, BorderStyle, TabStopType, ExternalHyperlink,
  UnderlineType, LevelFormat, WidthType, Table, TableRow, TableCell
} = require('docx');
const fs   = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

// ── Design constants ─────────────────────────────────────────────
const BLUE_NAME = "2F5496";
const BLUE_RULE = "1A5276";
const BLUE_ROLE = "2E75B6";
const BLACK     = "000000";
const DARK      = "1A1A1A";
const MID       = "333333";
const FONT      = "Calibri";
const PAGE_W    = 11906;
const PAGE_H    = 16838;
const MAR_L     = 567;
const MAR_R     = 567;
const MAR_T     = 340;
const MAR_B     = 567;
const CONTENT_W = PAGE_W - MAR_L - MAR_R;

// ── Helpers ──────────────────────────────────────────────────────

function spacer(before = 80) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [new TextRun({ text: "", size: 20, font: FONT })] });
}

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 120, after: 0 }, keepNext: true,
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE_RULE, space: 1 } },
    children: [new TextRun({ text, bold: true, size: 26, font: FONT, color: BLUE_NAME, allCaps: true })]
  });
}

function roleHeader(title, dates) {
  return new Paragraph({
    spacing: { before: 140, after: 0 }, keepNext: true,
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    children: [
      new TextRun({ text: title, bold: true, size: 20, font: FONT, color: BLUE_ROLE }),
      new TextRun({ text: "\t" }),
      new TextRun({ text: dates, bold: true, size: 20, font: FONT, color: BLACK })
    ]
  });
}

function companyLine(company, location) {
  return new Paragraph({
    spacing: { before: 0, after: 60 }, keepNext: true,
    children: [
      new TextRun({ text: company, bold: true, size: 20, font: FONT, color: BLACK }),
      new TextRun({ text: "  –  " + location, size: 20, font: FONT, color: BLACK })
    ]
  });
}

// Paragraph whose sole content is a single tag (for loop boundaries + bullet items).
function tagPara(tagText, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before ?? 20, after: opts.after ?? 20 },
    ...(opts.numbering ? { numbering: opts.numbering } : {}),
    keepNext: opts.keepNext ?? false,
    children: [new TextRun({ text: tagText, size: 20, font: FONT, color: MID })]
  });
}

function bulletTagPara(tagText) {
  return tagPara(tagText, { numbering: { reference: "bullets", level: 0 } });
}

// Renders a bullet loop block for one role.
// openTag / closeTag e.g. "{#calvergy_bullets}" / "{/calvergy_bullets}"
function roleBulletLoop(openTag, closeTag) {
  return [
    tagPara(openTag,   { before: 20, after: 0 }),
    bulletTagPara("{text}"),
    tagPara(closeTag,  { before: 0, after: 20 }),
  ];
}

// Skills table: ONE template row with {#skills}{label} / {value}{/skills}
// docxtemplater repeats the row for each skills entry in the payload.
function skillsTable() {
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const noneI = { style: BorderStyle.NONE, size: 0, color: "FFFFFF", space: 0 };
  const labelCol = 2000;
  const valueCol = CONTENT_W - labelCol;
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [labelCol, valueCol],
    borders: { top: none, bottom: none, left: none, right: none, insideH: noneI, insideV: noneI },
    rows: [
      new TableRow({ children: [
        new TableCell({
          borders: { top: none, bottom: none, left: none, right: none },
          width: { size: labelCol, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 0, right: 120 },
          // {#skills} opens the loop; {label} is the skill category name
          children: [new Paragraph({ children: [new TextRun({ text: "{#skills}{label}", bold: true, size: 20, font: FONT, color: BLUE_NAME })] })]
        }),
        new TableCell({
          borders: { top: none, bottom: none, left: none, right: none },
          width: { size: valueCol, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 0, right: 0 },
          // {value} is the skill list; {/skills} closes the loop
          children: [new Paragraph({ children: [new TextRun({ text: "{value}{/skills}", size: 20, font: FONT, color: DARK })] })]
        })
      ]})
    ]
  });
}

function projectTitle(title, url) {
  const run = new TextRun({ text: title, bold: true, size: 20, font: FONT, color: BLUE_ROLE, underline: { type: UnderlineType.SINGLE } });
  return new Paragraph({
    spacing: { before: 100, after: 20 }, keepNext: true,
    children: url ? [new ExternalHyperlink({ link: url, children: [run] })] : [run]
  });
}

function projectBullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 }, spacing: { before: 20, after: 20 },
    children: [new TextRun({ text, size: 20, font: FONT, color: MID })]
  });
}

function eduLine(degree, field, university, dates) {
  return new Paragraph({
    spacing: { before: 60, after: 20 },
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    children: [
      new TextRun({ text: degree + " in ", bold: true, size: 20, font: FONT, color: BLUE_ROLE }),
      new TextRun({ text: field, bold: true, size: 20, font: FONT, color: BLUE_ROLE }),
      new TextRun({ text: " – " + university, size: 20, font: FONT, color: BLACK }),
      new TextRun({ text: "\t" }),
      new TextRun({ text: dates, bold: true, size: 20, font: FONT, color: BLACK })
    ]
  });
}

function certBullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 }, spacing: { before: 20, after: 20 },
    children: [new TextRun({ text, size: 20, font: FONT, color: MID })]
  });
}

function fixAndSave(buf, outputPath) {
  const zip = new AdmZip(buf);
  let xml = zip.readAsText('word/document.xml');
  xml = xml.replace(/<w:insideH w:val="single" w:color="auto" w:sz="4"\/>/g, '<w:insideH w:val="none" w:color="FFFFFF" w:sz="0"/>');
  xml = xml.replace(/<w:insideV w:val="single" w:color="auto" w:sz="4"\/>/g, '<w:insideV w:val="none" w:color="FFFFFF" w:sz="0"/>');
  zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, zip.toBuffer());
  console.log('Generated:', outputPath);
}

// ── Document ─────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 200 } } }
      }]
    }]
  },
  styles: { default: { document: { run: { font: FONT, size: 20, color: MID } } } },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MAR_T, right: MAR_R, bottom: MAR_B, left: MAR_L }
      }
    },
    children: [

      // ── HEADER ───────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 30 },
        children: [new TextRun({ text: "KHAWAJA BILAL UDDIN", bold: true, size: 48, font: FONT, color: BLUE_NAME })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "{headline}", size: 24, font: FONT, color: BLUE_NAME })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 20 },
        children: [new TextRun({ text: "Raunheim, Germany  |  +49 1521 5678891  |  mbilaluddin1994@gmail.com", size: 20, font: FONT, color: MID })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
        children: [
          new ExternalHyperlink({ link: "https://linkedin.com/in/khawaja-bilal-uddin", children: [new TextRun({ text: "linkedin.com/in/khawaja-bilal-uddin", size: 20, font: FONT, color: BLUE_ROLE, underline: { type: UnderlineType.SINGLE } })] }),
          new TextRun({ text: "  |  ", size: 20, font: FONT, color: MID }),
          new ExternalHyperlink({ link: "https://github.com/k-bilaluddin", children: [new TextRun({ text: "github.com/k-bilaluddin", size: 20, font: FONT, color: BLUE_ROLE, underline: { type: UnderlineType.SINGLE } })] })
        ]
      }),
      new Paragraph({
        spacing: { before: 80, after: 0 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE_NAME } },
        children: [new TextRun({ text: "" })]
      }),

      // ── SUMMARY ──────────────────────────────────────────────────
      sectionHeading("Professional Summary"),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: "{summary}", size: 20, font: FONT, color: MID })]
      }),
      spacer(100),

      // ── SKILLS ───────────────────────────────────────────────────
      // Fully dynamic: order and content vary per JD.
      // One template row is repeated for each {skills} entry.
      sectionHeading("Skills"),
      skillsTable(),
      spacer(100),

      // ── EXPERIENCE ───────────────────────────────────────────────
      // Role headers/company/dates are FIXED. Only bullets are dynamic.
      sectionHeading("Work Experience"),

      // Role 1: Calvergy
      roleHeader("Software Engineer", "10/2025 – 01/2026"),
      companyLine("Calvergy UG", "Aachen, Germany – Remote"),
      ...roleBulletLoop("{#calvergy_bullets}", "{/calvergy_bullets}"),
      spacer(80),

      // Role 2: Senior at Bari's
      roleHeader("Senior Software Developer", "01/2023 – 01/2025"),
      companyLine("Bari’s Technology Solutions", "Karachi, Pakistan"),
      ...roleBulletLoop("{#senior_baris_bullets}", "{/senior_baris_bullets}"),
      spacer(80),

      // Role 3: Developer at Bari's
      roleHeader("Software Developer", "01/2021 – 12/2022"),
      companyLine("Bari’s Technology Solutions", "Karachi, Pakistan"),
      ...roleBulletLoop("{#developer_baris_bullets}", "{/developer_baris_bullets}"),
      spacer(80),

      // Role 4: Junior at Bari's
      roleHeader("Junior Software Developer", "04/2019 – 12/2020"),
      companyLine("Bari’s Technology Solutions", "Karachi, Pakistan"),
      ...roleBulletLoop("{#junior_baris_bullets}", "{/junior_baris_bullets}"),
      spacer(100),

      // ── EDUCATION ─────────────────────────────────────────────────
      sectionHeading("Education"),
      eduLine("MSCS", "Computer Science", "Bahria University, Karachi", "03/2021 – 01/2023"),
      eduLine("BSCS", "Computer Science", "Bahria University, Karachi", "02/2015 – 01/2019"),
      spacer(100),

      // ── PROJECTS ──────────────────────────────────────────────────
      sectionHeading("Notable Projects"),

      projectTitle("GlobalPost Multi-Leg Shipping Consolidator", "https://dev.goglobalpost.com"),
      projectBullet("Architected and built the full REST API suite on .NET 8 microservices processing 3B+ annual orders across multi-carrier flows (USPS, DHL, DPD, UPS, FedEx), serving 3M+ active shippers"),

      projectTitle("GlobalPost Shipment Tracking System", "https://www.goglobalpost.com/tracking/"),
      projectBullet("High-throughput tracking engine handling millions of daily transactions, leveraging Redis caching, DynamoDB, and multi-source data aggregation"),

      projectTitle("GlobalPost Admin & Shipper Portals", "https://portal.goglobalpost.com/login"),
      projectBullet("Developed with React.js + .NET 8, serving 3M+ internal shippers with JWT-secured, role-based access control"),

      projectTitle("JobVault – AI-Driven Job Application Platform", "https://github.com/k-bilaluddin/jobvault"),
      projectBullet("Architected a self-hosted platform (.NET 9, Vue 3, MongoDB) where a Claude-powered agent evaluates job postings and selects real experience from a curated bullet library to generate tailored CVs and cover letters, removing manual tailoring and constraining generation against hallucinated experience"),
      projectBullet("Built an event-driven pipeline (RabbitMQ, Worker service) that generates DOCX/PDF documents and commits them atomically to a GitHub-based document vault via the Git Trees API, with dead-letter retry handling"),
      projectBullet("Deployed via a full GitHub Actions CI/CD pipeline and Cloudflare Tunnel, sustaining 11ms average response time at 0% error rate under Postman-based load testing (20 virtual users, 2,000 requests)"),
      spacer(100),

      // ── CERTIFICATIONS ────────────────────────────────────────────
      sectionHeading("Certifications & Achievements"),
      certBullet("AI-200: Azure AI Cloud Developer Associate – In Progress"),
      certBullet("Represented Bari’s Technology Solutions at GITEX GLOBAL 2023, Dubai – resulting in 5 qualified leads"),
      certBullet("Recognised as ‘Reliability Maestro’ for delivering 100% on time and resolving critical issues under pressure"),
      spacer(100),

      // ── LANGUAGES ─────────────────────────────────────────────────
      sectionHeading("Languages"),
      certBullet("English – C1 Advanced (IELTS 7.0)"),
      certBullet("German – B1 in Progress"),
    ]
  }]
});

const OUTPUT = path.resolve(__dirname, '../templates/cv_template.docx');
Packer.toBuffer(doc)
  .then(buf => fixAndSave(buf, OUTPUT))
  .catch(e => { console.error(e); process.exit(1); });

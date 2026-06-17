/**
 * Generates templates/cl_template.docx
 * Run once: node scripts/build-cl-template.js
 *
 * Visual design is identical to cl_template_locked.js.
 * Text content is replaced with docxtemplater tags so the output
 * .docx acts as a template — the generation service fills it at runtime.
 */
const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, BorderStyle, TabStopType, ExternalHyperlink,
  UnderlineType
} = require('docx');
const fs   = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

// ── Design constants (must match the locked script exactly) ──────
const BLUE_NAME = "2F5496";
const BLUE_ROLE = "2E75B6";
const MID       = "333333";
const FONT      = "Calibri";
const PAGE_W    = 11906;
const PAGE_H    = 16838;
const MAR_L     = 851;
const MAR_R     = 851;
const MAR_T     = 340;
const MAR_B     = 567;
const CONTENT_W = PAGE_W - MAR_L - MAR_R;

function spacer(before = 80) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [new TextRun({ text: "", size: 20, font: FONT })] });
}

// A body paragraph whose sole content is a docxtemplater tag string.
// paragraphLoop mode requires each loop tag to own its paragraph.
function tagPara(tagText) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 160, line: 276, lineRule: "auto" },
    children: [new TextRun({ text: tagText, size: 20, font: FONT, color: MID })]
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

const doc = new Document({
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
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 30 },
        children: [new TextRun({ text: "KHAWAJA BILAL UDDIN", bold: true, size: 48, font: FONT, color: BLUE_NAME })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "{headline}", size: 24, font: FONT, color: BLUE_NAME })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
        children: [new TextRun({ text: "Raunheim, Germany  |  +49 1521 5678891  |  mbilaluddin1994@gmail.com", size: 20, font: FONT, color: MID })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [
          new ExternalHyperlink({ link: "https://linkedin.com/in/khawaja-bilal-uddin", children: [
            new TextRun({ text: "linkedin.com/in/khawaja-bilal-uddin", size: 20, font: FONT, color: BLUE_ROLE, underline: { type: UnderlineType.SINGLE } })
          ]}),
          new TextRun({ text: "  |  ", size: 20, font: FONT, color: MID }),
          new ExternalHyperlink({ link: "https://github.com/k-bilaluddin", children: [
            new TextRun({ text: "github.com/k-bilaluddin", size: 20, font: FONT, color: BLUE_ROLE, underline: { type: UnderlineType.SINGLE } })
          ]})
        ]
      }),
      new Paragraph({
        spacing: { before: 80, after: 0 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE_NAME } },
        children: [new TextRun({ text: "" })]
      }),

      spacer(120),

      // ── DATE (right-aligned) ─────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [
          new TextRun({ text: "\t", size: 20, font: FONT }),
          new TextRun({ text: "{date}", size: 20, font: FONT, color: MID })
        ]
      }),

      spacer(160),

      // ── SUBJECT LINE ─────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 160 },
        children: [
          new TextRun({ text: "{company}  —  ", size: 20, font: FONT, color: MID }),
          new TextRun({ text: "Application for {role}", bold: true, size: 20, font: FONT, color: BLUE_NAME })
        ]
      }),

      // ── SALUTATION ───────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 160 },
        children: [new TextRun({ text: "{recipient}", size: 20, font: FONT, color: MID })]
      }),

      // ── BODY — paragraphLoop iterates coverLetterParagraphs ──────
      // Each tag must be the sole content of its paragraph for paragraphLoop to work.
      tagPara("{#paragraphs}"),
      tagPara("{text}"),
      tagPara("{/paragraphs}"),

      spacer(160),

      // ── SIGN OFF ─────────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Kind regards,", size: 20, font: FONT, color: MID })]
      }),
      spacer(120),
      new Paragraph({
        spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: "Khawaja Bilal Uddin", bold: true, size: 20, font: FONT, color: BLUE_NAME })]
      }),
    ]
  }]
});

const OUTPUT = path.resolve(__dirname, '../templates/cl_template.docx');
Packer.toBuffer(doc)
  .then(buf => fixAndSave(buf, OUTPUT))
  .catch(e => { console.error(e); process.exit(1); });

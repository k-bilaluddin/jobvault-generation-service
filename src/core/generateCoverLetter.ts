import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { GenerationPayload } from '../types/GenerationPayload';
import { processBoldMarkdown } from './boldMarkdown';

const TEMPLATE_PATH = path.resolve(__dirname, '../../templates/cl_template.docx');

export function generateCoverLetter(payload: GenerationPayload): Buffer {
  const templateBuf = fs.readFileSync(TEMPLATE_PATH);
  const zip = new PizZip(templateBuf);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  const today = new Date();
  const date = `${String(today.getDate()).padStart(2, '0')}.${String(today.getMonth() + 1).padStart(2, '0')}.${today.getFullYear()}`;

  doc.render({
    date,
    headline: payload.headline,
    company: payload.company,
    role: payload.role,
    recipient: payload.recipient,
    paragraphs: payload.coverLetterParagraphs.map((text) => ({ text })),
  });

  const renderedZip = doc.getZip();
  const docXml = renderedZip.file('word/document.xml')!.asText();
  renderedZip.file('word/document.xml', processBoldMarkdown(docXml));
  return renderedZip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

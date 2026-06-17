import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { GenerationPayload } from '../types/GenerationPayload';
import { processBoldMarkdown } from './boldMarkdown';

const TEMPLATE_PATH = path.resolve(__dirname, '../../templates/cv_template.docx');

export function generateCv(payload: GenerationPayload): Buffer {
  const templateBuf = fs.readFileSync(TEMPLATE_PATH);
  const zip = new PizZip(templateBuf);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  const roleMap: Record<string, { text: string }[]> = {};
  for (const r of payload.roles) {
    roleMap[r.id] = r.bullets.map((text) => ({ text }));
  }

  doc.render({
    headline: payload.headline,
    summary: payload.summary,
    skills: payload.skills,
    calvergy_bullets:        roleMap['calvergy']        ?? [],
    senior_baris_bullets:    roleMap['senior_baris']    ?? [],
    developer_baris_bullets: roleMap['developer_baris'] ?? [],
    junior_baris_bullets:    roleMap['junior_baris']    ?? [],
  });

  // Post-process: convert **markdown bold** to proper Word bold runs
  const renderedZip = doc.getZip();
  const docXml = renderedZip.file('word/document.xml')!.asText();
  renderedZip.file('word/document.xml', processBoldMarkdown(docXml));

  return renderedZip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer;
}

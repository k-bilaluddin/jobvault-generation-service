import { Router, Request, Response } from 'express';
import { generateCv, VALID_ROLE_IDS } from '../core/generateCv';
import { generateCoverLetter } from '../core/generateCoverLetter';
import { GenerationPayload } from '../types/GenerationPayload';

const router = Router();

function validateCvPayload(payload: GenerationPayload): string | null {
  if (!payload.headline?.trim()) return 'headline is required';
  if (!payload.summary?.trim())  return 'summary is required';
  if (!Array.isArray(payload.skills) || payload.skills.length === 0)
    return 'skills must be a non-empty array';
  if (!Array.isArray(payload.roles) || payload.roles.length === 0)
    return 'roles must be a non-empty array';

  for (const r of payload.roles) {
    if (!VALID_ROLE_IDS.includes(r.id as typeof VALID_ROLE_IDS[number])) {
      return `invalid role id "${r.id}". Must be one of: ${VALID_ROLE_IDS.join(', ')}`;
    }
    if (!Array.isArray(r.bullets) || r.bullets.length === 0) {
      return `role "${r.id}" must have at least one bullet`;
    }
  }

  return null;
}

function validateClPayload(payload: GenerationPayload): string | null {
  if (!payload.headline?.trim())   return 'headline is required';
  if (!payload.company?.trim())    return 'company is required';
  if (!payload.role?.trim())       return 'role is required';
  if (!payload.recipient?.trim())  return 'recipient is required';
  if (!Array.isArray(payload.coverLetterParagraphs) || payload.coverLetterParagraphs.length === 0)
    return 'coverLetterParagraphs must be a non-empty array';

  return null;
}

router.get('/health', (_req: Request, res: Response) => {
  res.json({ status: 'ok' });
});

router.post('/api/generate-cv', (req: Request, res: Response) => {
  try {
    const payload = req.body as GenerationPayload;

    const validationError = validateCvPayload(payload);
    if (validationError) {
      res.status(400).json({ error: 'Invalid payload', detail: validationError });
      return;
    }

    const docBuffer = generateCv(payload);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="CV_${payload.company}_${payload.role}.docx"`,
      'Content-Length': docBuffer.length,
    });
    res.send(docBuffer);
  } catch (err) {
    res.status(500).json({ error: 'CV generation failed', detail: String(err) });
  }
});

router.post('/api/generate-cover-letter', (req: Request, res: Response) => {
  try {
    const payload = req.body as GenerationPayload;

    const validationError = validateClPayload(payload);
    if (validationError) {
      res.status(400).json({ error: 'Invalid payload', detail: validationError });
      return;
    }

    const docBuffer = generateCoverLetter(payload);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="CoverLetter_${payload.company}_${payload.role}.docx"`,
      'Content-Length': docBuffer.length,
    });
    res.send(docBuffer);
  } catch (err) {
    res.status(500).json({ error: 'Cover letter generation failed', detail: String(err) });
  }
});

export default router;

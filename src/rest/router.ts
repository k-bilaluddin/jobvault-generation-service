import { Router, Request, Response } from 'express';
import { trace, SpanStatusCode } from '@opentelemetry/api';
import { generateCv } from '../core/generateCv';
import { generateCoverLetter } from '../core/generateCoverLetter';
import { GenerationPayload } from '../types/GenerationPayload';

const router = Router();
const tracer = trace.getTracer('jobvault-generation-service');

router.get('/health', (_req: Request, res: Response) => {
  res.json({ status: 'ok' });
});

router.post('/api/generate-cv', (req: Request, res: Response) => {
  const span = tracer.startSpan('generate-cv');
  try {
    const payload = req.body as GenerationPayload;
    const docBuffer = generateCv(payload);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="CV_${payload.company}_${payload.role}.docx"`,
      'Content-Length': docBuffer.length,
    });
    res.send(docBuffer);
  } catch (err) {
    span.setStatus({ code: SpanStatusCode.ERROR, message: String(err) });
    res.status(500).json({ error: 'CV generation failed', detail: String(err) });
  } finally {
    span.end();
  }
});

router.post('/api/generate-cover-letter', (req: Request, res: Response) => {
  const span = tracer.startSpan('generate-cover-letter');
  try {
    const payload = req.body as GenerationPayload;
    const docBuffer = generateCoverLetter(payload);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="CoverLetter_${payload.company}_${payload.role}.docx"`,
      'Content-Length': docBuffer.length,
    });
    res.send(docBuffer);
  } catch (err) {
    span.setStatus({ code: SpanStatusCode.ERROR, message: String(err) });
    res.status(500).json({ error: 'Cover letter generation failed', detail: String(err) });
  } finally {
    span.end();
  }
});

export default router;

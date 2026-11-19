/**
 * Velciti Bore Log Generator — Backend Server
 *
 * POST /api/export-pdf
 *   Body: { boreLog: BoreLog, svgMarkup: string }
 *   Response: application/pdf blob
 *
 * POST /api/export-excel
 *   Body: { boreLog: BoreLog }
 *   Response: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
 *
 * POST /api/export-svg
 *   Body: { boreLog: BoreLog, svgMarkup: string }
 *   Response: image/svg+xml
 *
 * PDF is generated with pdfkit + svg-to-pdfkit (pure Node.js, no browser).
 * Excel is generated with ExcelJS (pure Node.js, no browser).
 * Playwright / Chromium is NOT used and NOT required.
 */

import express from 'express';
import cors from 'cors';
import PDFDocument from 'pdfkit';
import SVGtoPDF from 'svg-to-pdfkit';
import type { Request, Response } from 'express';
import { buildExcelWorkbook } from './excelExport';
import type { BoreLog } from './types';

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors({ origin: ['https://borelog.velciti.com', 'http://localhost:5173'] }));
app.use(express.json({ limit: '5mb' }));

function buildFilename(boreLog: BoreLog, ext: string): string {
  const sanitize = (s: string) =>
    String(s || '').replace(/[\\/:*?"<>|]/g, '').trim().replace(/\s+/g, '_');
  const bhNo = sanitize(boreLog.boreholeInfo?.bhNo) || 'BH';
  const project = sanitize(boreLog.boreholeInfo?.project) || 'Project';
  return `BoreLog_${bhNo}_${project}.${ext}`;
}

app.get('/api/health', (_req: Request, res: Response) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.post('/api/export-pdf', async (req: Request, res: Response) => {
  const { boreLog, svgMarkup } = req.body as { boreLog: BoreLog; svgMarkup: string };

  if (!boreLog?.boreholeInfo) {
    res.status(400).json({ error: 'Invalid bore log data' });
    return;
  }
  if (typeof svgMarkup !== 'string' || !svgMarkup.trimStart().startsWith('<svg')) {
    res.status(400).json({ error: 'Missing or invalid SVG markup' });
    return;
  }
  if (Buffer.byteLength(svgMarkup, 'utf8') > 2 * 1024 * 1024) {
    res.status(400).json({ error: 'SVG payload exceeds 2 MB limit' });
    return;
  }

  try {
    const widthMatch = svgMarkup.match(/\bwidth="(\d+(?:\.\d+)?)"/);
    const heightMatch = svgMarkup.match(/\bheight="(\d+(?:\.\d+)?)"/);
    const svgW = widthMatch ? parseFloat(widthMatch[1]) : 795;
    const svgH = heightMatch ? parseFloat(heightMatch[1]) : 1123;
    const pdfW = svgW * 0.75;
    const pdfH = svgH * 0.75;

    const doc = new PDFDocument({ size: [pdfW, pdfH], margin: 0, autoFirstPage: true });
    const chunks: Buffer[] = [];
    doc.on('data', (chunk: Buffer) => chunks.push(chunk));

    await new Promise<void>((resolve, reject) => {
      doc.on('end', resolve);
      doc.on('error', reject);
      SVGtoPDF(doc, svgMarkup, 0, 0, {
        width: pdfW,
        height: pdfH,
        preserveAspectRatio: 'xMinYMin meet',
      });
      doc.end();
    });

    const pdfBuffer = Buffer.concat(chunks);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${buildFilename(boreLog, 'pdf')}"`);
    res.send(pdfBuffer);
  } catch (err) {
    console.error('PDF export error:', err);
    res.status(500).json({
      error: 'PDF generation failed',
      detail: err instanceof Error ? err.message : String(err),
    });
  }
});

app.post('/api/export-excel', async (req: Request, res: Response) => {
  const { boreLog } = req.body as { boreLog: BoreLog };

  if (!boreLog?.boreholeInfo) {
    res.status(400).json({ error: 'Invalid bore log data' });
    return;
  }

  try {
    const wb = await buildExcelWorkbook(boreLog);
    const xlsxBuffer = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${buildFilename(boreLog, 'xlsx')}"`);
    res.send(Buffer.from(xlsxBuffer));
  } catch (err) {
    console.error('Excel export error:', err);
    res.status(500).json({
      error: 'Excel generation failed',
      detail: err instanceof Error ? err.message : String(err),
    });
  }
});

app.post('/api/export-svg', async (req: Request, res: Response) => {
  const { boreLog, svgMarkup } = req.body as { boreLog: BoreLog; svgMarkup: string };

  if (!boreLog?.boreholeInfo) {
    res.status(400).json({ error: 'Invalid bore log data' });
    return;
  }
  if (typeof svgMarkup !== 'string' || !svgMarkup.trimStart().startsWith('<svg')) {
    res.status(400).json({ error: 'Missing or invalid SVG markup' });
    return;
  }
  if (Buffer.byteLength(svgMarkup, 'utf8') > 2 * 1024 * 1024) {
    res.status(400).json({ error: 'SVG payload exceeds 2 MB limit' });
    return;
  }

  res.setHeader('Content-Type', 'image/svg+xml');
  res.setHeader('Content-Disposition', `attachment; filename="${buildFilename(boreLog, 'svg')}"`);
  res.send(svgMarkup);
});

app.listen(PORT, () => {
  console.log(`\n✅  Velciti Bore Log backend running on http://localhost:${PORT}`);
  console.log(`    Health:  http://localhost:${PORT}/api/health`);
  console.log(`    PDF:     POST http://localhost:${PORT}/api/export-pdf`);
  console.log(`    Excel:   POST http://localhost:${PORT}/api/export-excel`);
  console.log(`    SVG:     POST http://localhost:${PORT}/api/export-svg\n`);
});

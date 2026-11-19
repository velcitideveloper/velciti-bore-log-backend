"use strict";
/**
 * Velciti Bore Log Generator — Backend Server
 *
 * POST /api/export-pdf
 *   Body: BoreLog JSON
 *   Response: application/pdf blob
 *
 * Uses Playwright (headless Chromium) to render the React frontend
 * bore log SVG and export it to PDF.
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const playwright_1 = require("playwright");
const excelExport_1 = require("./excelExport");
const types_1 = require("./types");
const app = (0, express_1.default)();
const PORT = process.env.PORT || 3001;
const DEFAULT_ALLOWED_ORIGINS = ['http://localhost:5173', 'https://borelog.velciti.com'];
const allowedOrigins = process.env.CORS_ALLOWED_ORIGINS
    ? process.env.CORS_ALLOWED_ORIGINS.split(',').map((o) => o.trim())
    : DEFAULT_ALLOWED_ORIGINS;
const corsOptions = {
    origin(origin, callback) {
        if (!origin || allowedOrigins.includes(origin)) {
            callback(null, true);
        }
        else {
            callback(new Error(`Origin ${origin} not allowed by CORS`));
        }
    },
    methods: ['GET', 'POST', 'PATCH', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
};
app.use((0, cors_1.default)(corsOptions));
app.options('*', (0, cors_1.default)(corsOptions));
app.use(express_1.default.json({ limit: '5mb' }));
/** `BoreLog_[BoreholeNo]_[ProjectName].ext`, invalid filename characters stripped. */
function buildFilename(boreLog, ext) {
    const sanitize = (s) => String(s || '').replace(/[\\/:*?"<>|]/g, '').trim().replace(/\s+/g, '_');
    const bhNo = sanitize(boreLog.boreholeInfo?.bhNo) || 'BH';
    const project = sanitize(boreLog.boreholeInfo?.project) || 'Project';
    return `BoreLog_${bhNo}_${project}.${ext}`;
}
// ─── Health check ────────────────────────────────────────────
app.get('/api/health', (_req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
});
// ─── PDF Export ──────────────────────────────────────────────
app.post('/api/export-pdf', async (req, res) => {
    const boreLog = req.body;
    if (!boreLog || !boreLog.boreholeInfo) {
        res.status(400).json({ error: 'Invalid bore log data' });
        return;
    }
    let browser;
    try {
        browser = await playwright_1.chromium.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox'],
        });
        const page = await browser.newPage();
        // Set viewport to A4 width at 96dpi
        await page.setViewportSize({ width: 794, height: 1123 });
        // Navigate to the frontend with bore log data injected via URL hash
        // The frontend reads the hash to load the bore log data for PDF rendering
        const frontendUrl = 'https://borelog.velciti.com/?mode=pdf';
        await page.goto(frontendUrl, { waitUntil: 'networkidle', timeout: 30000 });
        // Inject bore log data into the page and trigger render
        await page.evaluate((data) => {
            window.__BORE_LOG_DATA__ = data;
            window.dispatchEvent(new CustomEvent('bore-log-data', { detail: data }));
        }, boreLog);
        // Wait for SVG to render
        await page.waitForSelector('#bore-log-svg', { timeout: 10000 });
        await page.waitForTimeout(500); // let fonts settle
        // Get SVG dimensions
        const svgDimensions = await page.evaluate(() => {
            const svg = document.getElementById('bore-log-svg');
            if (!svg)
                return { width: 794, height: 1123 };
            const rect = svg.getBoundingClientRect();
            return { width: rect.width, height: rect.height };
        });
        // Export as PDF
        const pdfBuffer = await page.pdf({
            width: `${svgDimensions.width}px`,
            height: `${svgDimensions.height + 20}px`,
            printBackground: true,
            margin: { top: '0px', right: '0px', bottom: '0px', left: '0px' },
        });
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="${buildFilename(boreLog, 'pdf')}"`);
        res.send(Buffer.from(pdfBuffer));
    }
    catch (err) {
        console.error('PDF export error:', err);
        res.status(500).json({
            error: 'PDF generation failed',
            detail: err instanceof Error ? err.message : String(err),
        });
    }
    finally {
        if (browser)
            await browser.close();
    }
});
// ─── Excel Export ────────────────────────────────────────────
app.post('/api/export-excel', async (req, res) => {
    const boreLog = req.body;
    if (!boreLog || !boreLog.boreholeInfo) {
        res.status(400).json({ error: 'Invalid bore log data' });
        return;
    }
    let browser;
    try {
        browser = await playwright_1.chromium.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox'],
        });
        const page = await browser.newPage();
        await page.setViewportSize({ width: 794, height: 1123 });
        await page.goto('https://borelog.velciti.com/?mode=pdf', { waitUntil: 'networkidle', timeout: 30000 });
        await page.evaluate((data) => {
            window.__BORE_LOG_DATA__ = data;
            window.dispatchEvent(new CustomEvent('bore-log-data', { detail: data }));
        }, boreLog);
        await page.waitForSelector('#bore-log-svg', { timeout: 10000 });
        await page.waitForTimeout(500);
        // Clip just the graph column (header through last data row), matching the
        // layout constants in frontend/src/components/BoreLogReport.tsx: Y_COL_H1=181,
        // Y_DATA_START=265, ROW_H=44, COLS.graph={x:551,w:175}. ExcelJS can't
        // reproduce this custom plot as a native chart, so it's captured as an
        // image here and embedded over that one column — every other cell stays native.
        const totalRows = (0, types_1.buildReportRows)(boreLog).length;
        const clip = { x: 551, y: 181, width: 175, height: (265 - 181) + totalRows * 44 };
        const graphImage = await page.screenshot({ clip });
        const wb = await (0, excelExport_1.buildExcelWorkbook)(boreLog, graphImage);
        const xlsxBuffer = await wb.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${buildFilename(boreLog, 'xlsx')}"`);
        res.send(Buffer.from(xlsxBuffer));
    }
    catch (err) {
        console.error('Excel export error:', err);
        res.status(500).json({
            error: 'Excel generation failed',
            detail: err instanceof Error ? err.message : String(err),
        });
    }
    finally {
        if (browser)
            await browser.close();
    }
});
// ─── SVG Export (alternative lightweight export) ─────────────
app.post('/api/export-svg', async (req, res) => {
    const boreLog = req.body;
    if (!boreLog || !boreLog.boreholeInfo) {
        res.status(400).json({ error: 'Invalid bore log data' });
        return;
    }
    let browser;
    try {
        browser = await playwright_1.chromium.launch({ headless: true, args: ['--no-sandbox'] });
        const page = await browser.newPage();
        await page.goto('https://borelog.velciti.com/?mode=pdf', { waitUntil: 'networkidle' });
        await page.evaluate((data) => {
            window.__BORE_LOG_DATA__ = data;
            window.dispatchEvent(new CustomEvent('bore-log-data', { detail: data }));
        }, boreLog);
        await page.waitForSelector('#bore-log-svg', { timeout: 10000 });
        await page.waitForTimeout(500);
        const svgContent = await page.evaluate(() => {
            const svg = document.getElementById('bore-log-svg');
            return svg ? svg.outerHTML : '<svg/>';
        });
        const bhNo = boreLog.boreholeInfo.bhNo?.replace(/\s/g, '_') || 'BH';
        res.setHeader('Content-Type', 'image/svg+xml');
        res.setHeader('Content-Disposition', `attachment; filename="BoreLog_${bhNo}.svg"`);
        res.send(svgContent);
    }
    catch (err) {
        res.status(500).json({ error: 'SVG export failed' });
    }
    finally {
        if (browser)
            await browser.close();
    }
});
// ─── CORS error handler ──────────────────────────────────────
// Without this, a rejected origin bubbles up as a default Express 500
// with a full stack trace instead of a clean, header-less rejection.
app.use((err, _req, res, next) => {
    if (err.message.startsWith('Origin ') && err.message.endsWith('not allowed by CORS')) {
        res.status(403).json({ error: 'Origin not allowed by CORS' });
        return;
    }
    next(err);
});
app.listen(PORT, () => {
    console.log(`\n✅  Velciti Bore Log backend running on http://localhost:${PORT}`);
    console.log(`    Health: http://localhost:${PORT}/api/health`);
    console.log(`    PDF export: POST http://localhost:${PORT}/api/export-pdf`);
    console.log(`    Excel export: POST http://localhost:${PORT}/api/export-excel\n`);
});

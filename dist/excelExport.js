"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildExcelWorkbook = buildExcelWorkbook;
/**
 * Native, editable .xlsx generation for the Velciti Bore Log report.
 * Mirrors the SVG layout in frontend/src/components/BoreLogReport.tsx:
 * same header text, same column header merges, same per-layer cell
 * merging, same column order. Soil profile patterns use ExcelJS's
 * native pattern fills (not images) so cells stay fully editable.
 */
const exceljs_1 = __importDefault(require("exceljs"));
const types_1 = require("./types");
// ─── Soil pattern colours + closest built-in Excel fill pattern ─────
// Colours must match frontend/src/data/soilPatterns.ts exactly.
const SOIL_PATTERNS = {
    REDDISH_CLAYEY_SAND: { backgroundColor: '#7C3B06', patternColor: '#1a0a00', excelPattern: 'gray0625' },
    GREYISH_SILTY_SAND: { backgroundColor: '#8C8070', patternColor: '#1a1200', excelPattern: 'gray0625' },
    BROWNISH_SILTY_SAND: { backgroundColor: '#7B4829', patternColor: '#1a0a00', excelPattern: 'gray0625' },
    BROWNISH_POORLY_GRADED_SAND: { backgroundColor: '#8B6040', patternColor: '#1a0a00', excelPattern: 'gray0625' },
    GREYISH_POORLY_GRADED_SAND: { backgroundColor: '#907060', patternColor: '#1a1200', excelPattern: 'gray0625' },
    GREY_GRANITE_ROCK: { backgroundColor: '#B8B8B8', patternColor: '#606060', excelPattern: 'darkUp' },
    SANDY_CLAY: { backgroundColor: '#9B6040', patternColor: '#1a0a00', excelPattern: 'lightHorizontal' },
    CLAY: { backgroundColor: '#C4A882', patternColor: '#604020', excelPattern: 'lightGrid' },
    GRAVEL: { backgroundColor: '#B0A880', patternColor: '#504028', excelPattern: 'lightTrellis' },
    FILLED_SOIL: { backgroundColor: '#806040', patternColor: '#201000', excelPattern: 'darkTrellis' },
    CUSTOM: { backgroundColor: '#D0D0D0', patternColor: '#606060', excelPattern: 'lightGray' },
};
function argb(hex) {
    return 'FF' + hex.replace('#', '').toUpperCase();
}
function fmtDate(d) {
    if (!d)
        return '';
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(d))
        return d;
    const m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? `${m[3]}/${m[2]}/${m[1]}` : d;
}
const COL = { depth: 1, profile: 2, desc: 3, thick: 4, sptDepth: 5, nVal: 6, relDens: 7, graph: 8, remarks: 9 };
const LAST_COL = 9;
const DARK_GREEN = 'FF196128';
const MED_GREEN = 'FF28A042';
const LIGHT_GREEN = 'FF84E098';
const BLACK = 'FF000000';
const thin = {
    top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' },
};
const medium = {
    top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' },
};
function colLetter(n) {
    let s = '';
    while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}
/**
 * Builds the full workbook. `graphImage` is a PNG buffer of the
 * live report's graph column (header through last data row) —
 * ExcelJS's native chart API can't faithfully reproduce this custom
 * plot (connecting line + refusal-at-100 behaviour), so per the
 * agreed fallback it's embedded as a picture over that one column;
 * every other cell in the sheet stays native and editable.
 */
async function buildExcelWorkbook(boreLog, graphImage) {
    const wb = new exceljs_1.default.Workbook();
    wb.creator = 'Velciti Bore Log Generator';
    wb.created = new Date();
    const ws = wb.addWorksheet('Bore Log', {
        pageSetup: {
            paperSize: 9, // A4
            orientation: 'portrait',
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
            margins: { left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0, footer: 0 },
        },
    });
    // Column widths (Excel character units ≈ px / 7), matching COLS in BoreLogReport.tsx
    const widthsPx = { depth: 50, profile: 42, desc: 185, thick: 42, sptDepth: 62, nVal: 90, relDens: 70, graph: 175, remarks: 59 };
    ws.columns = [
        { width: widthsPx.depth / 7 }, { width: widthsPx.profile / 7 }, { width: widthsPx.desc / 7 },
        { width: widthsPx.thick / 7 }, { width: widthsPx.sptDepth / 7 }, { width: widthsPx.nVal / 7 },
        { width: widthsPx.relDens / 7 }, { width: widthsPx.graph / 7 }, { width: widthsPx.remarks / 7 },
    ];
    const bi = boreLog.boreholeInfo;
    let r = 1;
    const mergeAndFill = (row, fromCol, toCol, text, opts = {}) => {
        if (toCol > fromCol)
            ws.mergeCells(row, fromCol, row, toCol);
        const cell = ws.getCell(row, fromCol);
        cell.value = text;
        cell.font = { name: 'Times New Roman', size: opts.size ?? 11, bold: opts.bold ?? true };
        cell.alignment = opts.align ?? { horizontal: 'center', vertical: 'middle', wrapText: true };
        if (opts.fill)
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
        const border = opts.border ?? thin;
        for (let c = fromCol; c <= toCol; c++)
            ws.getCell(row, c).border = border;
        if (opts.heightPt)
            ws.getRow(row).height = opts.heightPt;
    };
    // ── Company header ──
    mergeAndFill(r, 1, LAST_COL, 'VELCITI CONSULTING ENGINEERS PVT. LTD.,', { fill: DARK_GREEN, size: 14, heightPt: 20, border: medium });
    r++;
    mergeAndFill(r, 1, LAST_COL, '(Soil Exploration and Foundation Design Division)', { fill: DARK_GREEN, size: 13, heightPt: 19, border: medium });
    r++;
    mergeAndFill(r, 1, LAST_COL, 'BORE LOG', { fill: MED_GREEN, size: 15, heightPt: 17, border: medium });
    r++;
    // ── Project info rows ──
    const leftAlign = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
    mergeAndFill(r, 1, 5, `Client : ${bi.client}`, { fill: MED_GREEN, align: leftAlign, heightPt: 19 });
    mergeAndFill(r, 6, LAST_COL, `Date of Started : ${fmtDate(bi.dateStarted)}`, { fill: MED_GREEN, align: leftAlign, heightPt: 19 });
    r++;
    mergeAndFill(r, 1, 5, `Project : ${bi.project}`, { fill: MED_GREEN, align: leftAlign, heightPt: 19 });
    mergeAndFill(r, 6, LAST_COL, `Date of Completed : ${fmtDate(bi.dateCompleted)}`, { fill: MED_GREEN, align: leftAlign, heightPt: 19 });
    r++;
    mergeAndFill(r, 1, 5, `Location : ${bi.location}`, { fill: MED_GREEN, align: leftAlign, heightPt: 19 });
    mergeAndFill(r, 6, 7, `BH No: ${bi.bhNo}`, { fill: MED_GREEN, align: leftAlign, heightPt: 19 });
    mergeAndFill(r, 8, LAST_COL, `GWL: ${bi.groundWaterLevel}`, { fill: MED_GREEN, align: leftAlign, heightPt: 19 });
    r++;
    mergeAndFill(r, 1, LAST_COL, `Co-ordinates : ${bi.coordinates}`, { fill: MED_GREEN, align: leftAlign, heightPt: 19 });
    r++;
    // ── Column headers (3 rows: group header / sub-header / scale) ──
    const hdrRow1 = r, hdrRow2 = r + 1, hdrRow3 = r + 2;
    mergeAndFill(hdrRow1, COL.depth, COL.depth, 'Depth Below\nG.L. (m)', { fill: LIGHT_GREEN, size: 9, heightPt: 24 });
    ws.mergeCells(hdrRow1, COL.depth, hdrRow3, COL.depth);
    mergeAndFill(hdrRow1, COL.profile, COL.profile, 'Soil\nProfile', { fill: LIGHT_GREEN, size: 9, heightPt: 24 });
    ws.mergeCells(hdrRow1, COL.profile, hdrRow3, COL.profile);
    mergeAndFill(hdrRow1, COL.desc, COL.desc, 'Description of Soil', { fill: LIGHT_GREEN, heightPt: 24 });
    ws.mergeCells(hdrRow1, COL.desc, hdrRow3, COL.desc);
    mergeAndFill(hdrRow1, COL.thick, COL.thick, 'Thickness\nof layer\n(m)', { fill: LIGHT_GREEN, size: 9, heightPt: 24 });
    ws.mergeCells(hdrRow1, COL.thick, hdrRow3, COL.thick);
    mergeAndFill(hdrRow1, COL.sptDepth, COL.relDens, 'Standard Penetration Test Data', { fill: LIGHT_GREEN, heightPt: 24 });
    mergeAndFill(hdrRow1, COL.graph, COL.graph, '', { fill: LIGHT_GREEN, heightPt: 24 }); // graph image covers header text
    ws.mergeCells(hdrRow1, COL.graph, hdrRow2, COL.graph);
    mergeAndFill(hdrRow1, COL.remarks, COL.remarks, 'Remarks', { fill: LIGHT_GREEN, heightPt: 24 });
    ws.mergeCells(hdrRow1, COL.remarks, hdrRow3, COL.remarks);
    mergeAndFill(hdrRow2, COL.sptDepth, COL.sptDepth, 'Depth at which\ntest is conducted\n(m)', { fill: LIGHT_GREEN, size: 9, heightPt: 24 });
    ws.mergeCells(hdrRow2, COL.sptDepth, hdrRow3, COL.sptDepth);
    mergeAndFill(hdrRow2, COL.nVal, COL.nVal, 'N-Value', { fill: LIGHT_GREEN, heightPt: 24 });
    ws.mergeCells(hdrRow2, COL.nVal, hdrRow3, COL.nVal);
    mergeAndFill(hdrRow2, COL.relDens, COL.relDens, 'Relative density /\nConsistency', { fill: LIGHT_GREEN, size: 9, heightPt: 24 });
    ws.mergeCells(hdrRow2, COL.relDens, hdrRow3, COL.relDens);
    mergeAndFill(hdrRow3, COL.graph, COL.graph, '', { fill: LIGHT_GREEN, heightPt: 15 }); // scale row — graph image covers 0..100 labels
    ws.getRow(hdrRow1).height = 24;
    ws.getRow(hdrRow2).height = 24;
    ws.getRow(hdrRow3).height = 15;
    r = hdrRow3 + 1;
    // ── Data rows ──
    const rows = (0, types_1.buildReportRows)(boreLog);
    const dataStartRow = r;
    // Group consecutive rows by layer (mirrors DataRows' layerGroups in BoreLogReport.tsx)
    const layerGroups = [];
    rows.forEach(row => {
        const last = layerGroups[layerGroups.length - 1];
        if (last && last.layer.id === row.layer.id)
            last.count++;
        else
            layerGroups.push({ layer: row.layer, startRow: row.rowIndex, count: 1 });
    });
    rows.forEach((row, i) => {
        const excelRow = dataStartRow + i;
        const spt = row.sptRecord;
        const core = row.coreRecord;
        const isCore = row.layer.layerType === 'ROCK';
        ws.getRow(excelRow).height = 33; // 44px * 0.75
        // Each row shows its own override; empty/null rows fall back to the
        // parent layer description (every row, not just the first).
        const descText = spt?.descriptionOverride || row.layer.description;
        const descCell = ws.getCell(excelRow, COL.desc);
        descCell.value = descText;
        descCell.font = { name: 'Times New Roman', size: 11 };
        descCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
        descCell.border = thin;
        const sptDepthCell = ws.getCell(excelRow, COL.sptDepth);
        sptDepthCell.value = isCore
            ? (row.isFirstRowOfLayer ? `${row.layer.fromDepth.toFixed(2)} to ${row.layer.toDepth.toFixed(2)}` : '')
            : (spt ? spt.testDepth.toFixed(1) : '');
        sptDepthCell.font = { name: 'Times New Roman', size: 11 };
        sptDepthCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        sptDepthCell.border = thin;
        const nValCell = ws.getCell(excelRow, COL.nVal);
        if (spt)
            nValCell.value = `${spt.nValue || '—'}\n[${spt.blowCounts}]`;
        else if (isCore && row.isFirstRowOfLayer)
            nValCell.value = '-';
        nValCell.font = { name: 'Times New Roman', size: 10 };
        nValCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        nValCell.border = thin;
        const relDensCell = ws.getCell(excelRow, COL.relDens);
        relDensCell.value = isCore ? '-' : (spt?.relativeDensity || '—');
        relDensCell.font = { name: 'Times New Roman', size: 10 };
        relDensCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        relDensCell.border = thin;
        const graphCell = ws.getCell(excelRow, COL.graph); // image overlays this column
        graphCell.border = thin;
        const remarksCell = ws.getCell(excelRow, COL.remarks);
        remarksCell.value = isCore ? (core?.remarks || 'CORE') : (spt?.remarks ? String(spt.remarks) : '');
        remarksCell.font = { name: 'Times New Roman', size: 11 };
        remarksCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        remarksCell.border = thin;
    });
    // Per-layer merged cells: depth, soil-profile pattern fill, thickness
    layerGroups.forEach(g => {
        const startExcelRow = dataStartRow + g.startRow;
        const endExcelRow = startExcelRow + g.count - 1;
        const isCore = g.layer.layerType === 'ROCK';
        const pattern = SOIL_PATTERNS[g.layer.patternId];
        const bgColor = g.layer.color ?? pattern.backgroundColor;
        if (g.count > 1)
            ws.mergeCells(startExcelRow, COL.depth, endExcelRow, COL.depth);
        const depthCell = ws.getCell(startExcelRow, COL.depth);
        depthCell.value = isCore ? `${g.layer.fromDepth.toFixed(2)} to ${g.layer.toDepth.toFixed(2)}` : g.layer.toDepth.toFixed(1);
        depthCell.font = { name: 'Times New Roman', size: 11 };
        depthCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        for (let rr = startExcelRow; rr <= endExcelRow; rr++)
            ws.getCell(rr, COL.depth).border = thin;
        if (g.count > 1)
            ws.mergeCells(startExcelRow, COL.profile, endExcelRow, COL.profile);
        const profileCell = ws.getCell(startExcelRow, COL.profile);
        profileCell.fill = { type: 'pattern', pattern: pattern.excelPattern, fgColor: { argb: argb(pattern.patternColor) }, bgColor: { argb: argb(bgColor) } };
        for (let rr = startExcelRow; rr <= endExcelRow; rr++)
            ws.getCell(rr, COL.profile).border = thin;
        if (g.count > 1)
            ws.mergeCells(startExcelRow, COL.thick, endExcelRow, COL.thick);
        const thickCell = ws.getCell(startExcelRow, COL.thick);
        thickCell.value = g.layer.thickness.toFixed(1);
        thickCell.font = { name: 'Times New Roman', size: 11 };
        thickCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        for (let rr = startExcelRow; rr <= endExcelRow; rr++)
            ws.getCell(rr, COL.thick).border = thin;
    });
    const lastDataRow = dataStartRow + rows.length - 1;
    // ── Termination row ──
    const termRow = lastDataRow + 1;
    mergeAndFill(termRow, 1, LAST_COL, `Borehole Terminated @ ${bi.finalDepth.toFixed(2)}m`, { fill: DARK_GREEN, size: 12, heightPt: 18, border: medium });
    // ── Embed the N-value graph image over the graph column, header through last data row ──
    if (graphImage) {
        const imageId = wb.addImage({ buffer: graphImage, extension: 'png' });
        ws.addImage(imageId, {
            tl: { col: COL.graph - 1, row: hdrRow1 - 1 },
            br: { col: COL.graph, row: lastDataRow },
        });
    }
    return wb;
}

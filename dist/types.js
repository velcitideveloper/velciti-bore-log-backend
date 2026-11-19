"use strict";
/**
 * Minimal local mirror of the frontend's BoreLog shape.
 * Backend and frontend are separate npm packages (no shared workspace),
 * so this intentionally duplicates just the fields the backend reads
 * from the JSON body — it must stay structurally compatible with
 * frontend/src/types/index.ts, not textually identical to it.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildReportRows = buildReportRows;
/** Mirrors frontend/src/utils/reportLayout.ts buildReportRows() exactly. */
function buildReportRows(boreLog) {
    const rows = [];
    const sortedLayers = [...boreLog.layers].sort((a, b) => a.fromDepth - b.fromDepth);
    sortedLayers.forEach(layer => {
        const layerSptRecords = boreLog.sptRecords
            .filter(r => r.layerId === layer.id)
            .sort((a, b) => a.testDepth - b.testDepth);
        const layerCoreRecords = boreLog.coreRecords.filter(r => r.layerId === layer.id);
        if (layer.layerType === 'ROCK') {
            const coreRecs = layerCoreRecords.length > 0 ? layerCoreRecords : [undefined];
            coreRecs.forEach((coreRec, idx) => {
                rows.push({
                    rowIndex: rows.length,
                    layer,
                    coreRecord: coreRec,
                    isFirstRowOfLayer: idx === 0,
                    isLastRowOfLayer: idx === coreRecs.length - 1,
                });
            });
        }
        else {
            if (layerSptRecords.length === 0) {
                rows.push({ rowIndex: rows.length, layer, isFirstRowOfLayer: true, isLastRowOfLayer: true });
            }
            else {
                layerSptRecords.forEach((spt, idx) => {
                    rows.push({
                        rowIndex: rows.length,
                        layer,
                        sptRecord: spt,
                        isFirstRowOfLayer: idx === 0,
                        isLastRowOfLayer: idx === layerSptRecords.length - 1,
                    });
                });
            }
        }
    });
    return rows;
}

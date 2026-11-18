/**
 * Minimal local mirror of the frontend's BoreLog shape.
 * Backend and frontend are separate npm packages (no shared workspace),
 * so this intentionally duplicates just the fields the backend reads
 * from the JSON body — it must stay structurally compatible with
 * frontend/src/types/index.ts, not textually identical to it.
 */

export type SoilPatternId =
  | 'REDDISH_CLAYEY_SAND'
  | 'GREYISH_SILTY_SAND'
  | 'BROWNISH_SILTY_SAND'
  | 'BROWNISH_POORLY_GRADED_SAND'
  | 'GREYISH_POORLY_GRADED_SAND'
  | 'GREY_GRANITE_ROCK'
  | 'SANDY_CLAY'
  | 'CLAY'
  | 'GRAVEL'
  | 'FILLED_SOIL'
  | 'CUSTOM';

export interface BoreholeInfo {
  client: string;
  project: string;
  location: string;
  coordinates: string;
  bhNo: string;
  dateStarted: string;
  dateCompleted: string;
  groundWaterLevel: string;
  finalDepth: number;
  elevation?: number;
}

export interface GeologicalLayer {
  id: string;
  layerType: 'SOIL' | 'ROCK';
  fromDepth: number;
  toDepth: number;
  thickness: number;
  description: string;
  patternId: SoilPatternId;
  color?: string;
  sortOrder?: number;
}

export interface SptRecord {
  id: string;
  layerId: string;
  testDepth: number;
  nValue: string;
  nValueNumeric: number | null;
  blowCounts: string;
  relativeDensity: string;
  remarks: string;
  descriptionOverride?: string | null;
}

export interface CoreRecord {
  id: string;
  layerId: string;
  fromDepth: number;
  toDepth: number;
  description?: string;
  tcr?: string;
  rqd?: string;
  remarks: string;
}

export interface BoreLog {
  id: string;
  boreholeInfo: BoreholeInfo;
  layers: GeologicalLayer[];
  sptRecords: SptRecord[];
  coreRecords: CoreRecord[];
  createdAt: string;
  updatedAt: string;
}

export interface ReportRow {
  rowIndex: number;
  layer: GeologicalLayer;
  sptRecord?: SptRecord;
  coreRecord?: CoreRecord;
  isFirstRowOfLayer: boolean;
  isLastRowOfLayer: boolean;
}

/** Mirrors frontend/src/utils/reportLayout.ts buildReportRows() exactly. */
export function buildReportRows(boreLog: BoreLog): ReportRow[] {
  const rows: ReportRow[] = [];
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
    } else {
      if (layerSptRecords.length === 0) {
        rows.push({ rowIndex: rows.length, layer, isFirstRowOfLayer: true, isLastRowOfLayer: true });
      } else {
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

import ExcelJS from "exceljs";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import { CJ_UPLOAD_HEADERS, ORIGINAL_ITEM_COL } from "@/lib/constants/excel";
import { makeDatedFileName } from "@/lib/utils/filename";
import { normalizeHeader, toText } from "@/lib/utils/normalize";

type TRow = Record<string, any>;

const sanitizeFileName = (name: string) =>
  name
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim();

const buildRowForCj = (row: TRow) => {
  // 원본 row에서 CJ 15컬럼을 채워야 함.
  // 컬럼명이 원본과 동일하다고 가정하지 말고, "동일 이름" 우선 매핑 + 없으면 빈값
  // (필요하면 나중에 더 정교한 매핑을 추가 가능)
  const out: Record<string, any> = {};
  for (const h of CJ_UPLOAD_HEADERS) {
    out[h] = row[h] ?? "";
  }
  return out;
};

export const buildCjGroupedRows = (originalHeaders: string[], originalRows: TRow[]) => {
  // 헤더 정규화 맵 (공백 차이 대비)
  const headerMap = new Map<string, string>(); // normalized -> originalHeader
  for (const h of originalHeaders) headerMap.set(normalizeHeader(h), h);

  const get = (row: TRow, header: string) => {
    const key = headerMap.get(normalizeHeader(header));
    if (!key) return "";
    return row[key];
  };

  const groups = new Map<string, TRow[]>();
  for (const row of originalRows) {
    const itemName = String(get(row, ORIGINAL_ITEM_COL) ?? "").trim();
    if (!itemName) continue;
    const arr = groups.get(itemName) ?? [];
    // CJ 양식 row로 변환(가능한 값만 채움)
    const normalizedRow: TRow = {};
    for (const h of CJ_UPLOAD_HEADERS) {
      normalizedRow[h] = get(row, h) ?? "";
    }
    arr.push(normalizedRow);
    groups.set(itemName, arr);
  }

  return groups; // key: 품목명, value: CJ양식 row들
};

const makeCjWorkbookBuffer = async (rows: TRow[]) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.columns = CJ_UPLOAD_HEADERS.map((h) => ({
    header: h,
    key: h,
    width: h.length > 8 ? 18 : 14,
  }));

  for (const r of rows) {
    const rowData: Record<string, any> = {};
    for (const h of CJ_UPLOAD_HEADERS) rowData[h] = toText(r[h]);
    ws.addRow(rowData);
  }

  // 헤더 bold
  ws.getRow(1).font = { bold: true };

  // AutoFilter
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: rows.length + 1, column: CJ_UPLOAD_HEADERS.length },
  };

  return wb.xlsx.writeBuffer();
};

export const downloadCjUploadsZip = async (groups: Map<string, TRow[]>, onProgress?: (done: number, total: number) => void) => {
  const zip = new JSZip();

  // 파일이 너무 많아져도 멈춘 것처럼 보이지 않게 호출부에서 스피너/진행률 표시 권장
  for (const [itemName, rows] of groups.entries()) {
    const buf = await makeCjWorkbookBuffer(rows);
    const safe = sanitizeFileName(itemName);
    zip.file(`${safe}.xlsx`, buf);
  }

  const zipBlob = await zip.generateAsync({ type: "blob" });
  saveAs(zipBlob, makeDatedFileName("한섬누리_CJ제출용_품목별엑셀.zip"));
};

import ExcelJS from "exceljs";
import { CJ_REPLY_KEY_COL, CJ_REPLY_TRACKING_COL } from "@/lib/constants/excel";
import { normalizeHeader, toText } from "@/lib/utils/normalize";

export type TReplyRecord = {
  key: string; // 고객주문번호
  tracking: string; // 운송장번호
  sourceFileName: string;
};

export type TReadReplyResult = {
  records: TReplyRecord[];
  skipped: Array<{ reason: string; sourceFileName: string }>;
};

export const readCjReplyFile = async (
  file: File,
): Promise<TReadReplyResult> => {
  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const ws = wb.worksheets[0];
  if (!ws) throw new Error(`회신 파일 시트를 찾을 수 없습니다: ${file.name}`);

  // 헤더 찾기
  const headers: string[] = [];
  ws.getRow(1).eachCell((cell, colNumber) => {
    headers[colNumber - 1] = String(cell.value ?? "").trim();
  });

  const headerMap = new Map<string, number>(); // normalized header -> col index (1-based)
  headers.forEach((h, idx) => headerMap.set(normalizeHeader(h), idx + 1));

  const keyCol = headerMap.get(normalizeHeader(CJ_REPLY_KEY_COL));
  const trackingCol = headerMap.get(normalizeHeader(CJ_REPLY_TRACKING_COL));

  if (!keyCol || !trackingCol) {
    return {
      records: [],
      skipped: [
        {
          reason: "필수 컬럼(고객주문번호/운송장번호)을 찾지 못함",
          sourceFileName: file.name,
        },
      ],
    };
  }

  const records: TReplyRecord[] = [];
  const skipped: Array<{ reason: string; sourceFileName: string }> = [];

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const key = toText(row.getCell(keyCol).value).trim();
    const tracking = toText(row.getCell(trackingCol).value).trim();

    if (!key) {
      skipped.push({
        reason: "고객주문번호 비어있음",
        sourceFileName: file.name,
      });
      return;
    }
    if (!tracking) {
      // 너가 "항상 채워져 있다" 했지만, 방어로직은 유지
      skipped.push({
        reason: `운송장번호 비어있음 (고객주문번호=${key})`,
        sourceFileName: file.name,
      });
      return;
    }

    records.push({ key, tracking, sourceFileName: file.name });
  });

  return { records, skipped };
};

export const readCjReplyFiles = async (files: File[]) => {
  const map = new Map<string, string>();
  const orderFileMap = new Map<string, Set<string>>();

  for (const file of files) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const sheet = workbook.worksheets[0];

    sheet.eachRow((row, rowIdx) => {
      if (rowIdx === 1) return;

      const orderNo = String(row.getCell("고객주문번호").value ?? "").trim();
      const tracking = String(row.getCell("운송장번호").value ?? "").trim();
      if (!orderNo || !tracking) return;

      // ✅ 운송장 매핑 (기존 로직)
      map.set(orderNo, tracking);

      // ✅ 파일별 주문번호 추적
      if (!orderFileMap.has(orderNo)) {
        orderFileMap.set(orderNo, new Set());
      }
      orderFileMap.get(orderNo)!.add(file.name);
    });
  }

  return { map, orderFileMap };
};

import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { ORIGINAL_KEY_COL, ORIGINAL_TRACKING_COL } from "@/lib/constants/excel";
import { makeDatedFileName } from "@/lib/utils/filename";
import { normalizeHeader, toText } from "@/lib/utils/normalize";

type TRow = Record<string, any>;

export type TApplyResult = {
  updatedRows: TRow[];
  unmatched: Array<{ customerOrderNo: string; tracking: string }>;
  duplicates: Array<{ key: string; count: number }>;
};

export const applyTracking = (
  originalHeaders: string[],
  originalRows: TRow[],
  replyMap: Map<string, string>, // 고객주문번호 -> 운송장번호
): TApplyResult => {
  // header map
  const headerMap = new Map<string, string>();
  for (const h of originalHeaders) headerMap.set(normalizeHeader(h), h);

  const keyHeader = headerMap.get(normalizeHeader(ORIGINAL_KEY_COL)) ?? ORIGINAL_KEY_COL;
  const trackingHeader = headerMap.get(normalizeHeader(ORIGINAL_TRACKING_COL)) ?? ORIGINAL_TRACKING_COL;

  // duplicate keys in original
  const counts = new Map<string, number>();
  for (const r of originalRows) {
    const k = String(r[keyHeader] ?? "").trim();
    if (!k) continue;
    counts.set(k, (counts.get(k) ?? 0) + 1);
  }
  const duplicates = Array.from(counts.entries())
    .filter(([, c]) => c > 1)
    .map(([key, count]) => ({ key, count }));

  // apply
  const updated = originalRows.map((r) => ({ ...r }));
  const unmatched: Array<{ customerOrderNo: string; tracking: string }> = [];

  // index for fast lookup: key -> list of row indexes
  const index = new Map<string, number[]>();
  updated.forEach((r, i) => {
    const k = String(r[keyHeader] ?? "").trim();
    if (!k) return;
    const arr = index.get(k) ?? [];
    arr.push(i);
    index.set(k, arr);
  });

  for (const [customerOrderNo, tracking] of replyMap.entries()) {
    const idxs = index.get(customerOrderNo);
    if (!idxs || idxs.length === 0) {
      unmatched.push({ customerOrderNo, tracking });
      continue;
    }
    for (const idx of idxs) {
      updated[idx][trackingHeader] = tracking;
    }
  }

  // tracking 컬럼이 원본에 없을 수도 있으니, 헤더에 없으면 추가
  const finalHeaders = [...originalHeaders];
  if (!finalHeaders.some((h) => normalizeHeader(h) === normalizeHeader(trackingHeader))) {
    finalHeaders.push(trackingHeader);
    // 새 컬럼을 추가한 경우, 빈 값 보장
    for (const r of updated) {
      if (!(trackingHeader in r)) r[trackingHeader] = "";
    }
  }

  return { updatedRows: updated, unmatched, duplicates };
};

export const downloadOriginalWithTracking = async (headers: string[], rows: TRow[]) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.columns = headers.map((h) => ({ header: h, key: h, width: h.length > 10 ? 22 : 16 }));

  for (const r of rows) {
    const rowData: Record<string, any> = {};
    for (const h of headers) rowData[h] = toText(r[h]);
    ws.addRow(rowData);
  }

  ws.getRow(1).font = { bold: true };
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: rows.length + 1, column: headers.length },
  };

  const buf = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buf]), makeDatedFileName("한섬누리_운송장번호_반영완료.xlsx"));
};

export const downloadUnmatchedExcel = async (unmatched: Array<{ customerOrderNo: string; tracking: string }>) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("미매칭");

  ws.columns = [
    { header: "고객주문번호", key: "customerOrderNo", width: 24 },
    { header: "운송장번호", key: "tracking", width: 20 },
  ];

  unmatched.forEach((u) => ws.addRow(u));

  ws.getRow(1).font = { bold: true };
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: unmatched.length + 1, column: 2 },
  };

  const buf = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buf]), makeDatedFileName("한섬누리_미매칭_주문목록.xlsx"));
};

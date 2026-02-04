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

export const applyTracking = (originalHeaders: string[], originalRows: Array<Record<string, any>>, replyMap: Map<string, string[]>) => {
  const updatedHeaders = [...originalHeaders];

  const extraColNames = (n: number) => `운송장번호(${n})`; // 2부터 사용

  let maxTrackingCount = 1; // 최소 1(기존 운송장번호)

  // 1) replyMap 전체에서 최대 운송장 개수 구하기
  for (const list of replyMap.values()) {
    if (list.length > maxTrackingCount) maxTrackingCount = list.length;
  }

  // 2) 필요한 추가 컬럼을 헤더에 미리 붙이기 (2..max)
  for (let i = 2; i <= maxTrackingCount; i++) {
    const col = extraColNames(i);
    if (!updatedHeaders.includes(col)) updatedHeaders.push(col);
  }

  // 3) 매핑 적용
  const unmatched: Array<{ customerOrderNo: string; tracking: string }> = [];
  const duplicates: Array<{ key: string; count: number }> = []; // 너 기존 로직 그대로 두면 됨

  // 원본 중복키 계산 로직(기존 그대로)...

  const updatedRows = originalRows.map((row) => {
    const key = String(row["상품주문번호"] ?? "").trim(); // ✅ 원본 키
    if (!key) return row;

    const list = replyMap.get(key); // ✅ CJ 고객주문번호 == 상품주문번호 값
    if (!list || list.length === 0) return row;

    // 1번째는 기존 운송장번호 컬럼
    row["운송장번호"] = list[0];

    // 2번째부터는 추가 컬럼
    for (let i = 2; i <= list.length; i++) {
      row[extraColNames(i)] = list[i - 1];
    }

    return row;
  });

  // unmatched 만드는 로직은 기존이 Map<string,string> 기준일텐데
  // 이제 Map<string,string[]>니까 아래처럼 바꾸면 됨:
  for (const [customerOrderNo, list] of replyMap.entries()) {
    const existsInOriginal = originalRows.some((r) => String(r["상품주문번호"] ?? "").trim() === customerOrderNo);
    if (!existsInOriginal) {
      // 운송장 여러 개면 각각 기록 (원하면 1개만 기록도 가능)
      for (const tracking of list) unmatched.push({ customerOrderNo, tracking });
    }
  }

  return { updatedHeaders, updatedRows, unmatched, duplicates };
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

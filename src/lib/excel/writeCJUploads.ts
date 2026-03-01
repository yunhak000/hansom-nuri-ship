import ExcelJS from "exceljs";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import {
  CJ_UPLOAD_HEADERS,
  ORIGINAL_ITEM_COL,
  ORIGINAL_KEY_COL,
  ORIGINAL_FALLBACK_KEY_COL,
} from "@/lib/constants/excel";
import { makeDatedFileName } from "@/lib/utils/filename";
import { normalizeHeader, toText } from "@/lib/utils/normalize";

type TRow = Record<string, unknown>;

const sanitizeFileName = (name: string) =>
  name
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim();

export const buildCjGroupedRows = (
  originalHeaders: string[],
  originalRows: TRow[],
) => {
  // 헤더 정규화 맵 (공백 차이 대비)
  const headerMap = new Map<string, string>(); // normalized -> originalHeader
  for (const h of originalHeaders) headerMap.set(normalizeHeader(h), h);

  const get = (row: TRow, header: string) => {
    const key = headerMap.get(normalizeHeader(header));
    if (!key) return "";
    return row[key];
  };

  // ✅ 원본 키: 상품주문번호 우선, 없으면 ★쇼핑몰 주문번호★
  const getOrderKey = (row: TRow) => {
    const primary = toText(get(row, ORIGINAL_KEY_COL)).trim();
    if (primary) return primary;

    const fallback = toText(get(row, ORIGINAL_FALLBACK_KEY_COL)).trim();
    if (fallback) return fallback;

    return "";
  };

  const groups = new Map<string, TRow[]>();

  for (const row of originalRows) {
    const itemName = String(get(row, ORIGINAL_ITEM_COL) ?? "").trim();
    if (!itemName) continue;

    // ✅ CJ 제출용 "고객주문번호"에 넣을 키 결정
    const orderKey = getOrderKey(row);
    if (!orderKey) {
      // 키가 아예 없는 행은 CJ 업로드 파일로 만들면 이후 회신 매칭이 불가능하므로 제외(권장)
      continue;
    }

    const arr = groups.get(itemName) ?? [];

    // CJ 양식 row로 변환(가능한 값만 채움)
    const normalizedRow: TRow = {};
    for (const h of CJ_UPLOAD_HEADERS) {
      // ✅ "고객주문번호" 컬럼은 원본의 상품주문번호/★쇼핑몰 주문번호★로 강제 주입
      if (normalizeHeader(h) === normalizeHeader("고객주문번호")) {
        normalizedRow[h] = orderKey;
        continue;
      }

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
    const rowData: Record<string, unknown> = {};
    for (const h of CJ_UPLOAD_HEADERS) rowData[h] = toText(r[h]);
    ws.addRow(rowData);
  }

  // 헤더 bold
  ws.getRow(1).font = { bold: true };

  // ✅ 필터 기능 제거(요청사항)
  // ws.autoFilter = { ... }  // <- 삭제

  return wb.xlsx.writeBuffer();
};

export const downloadCjUploadsZip = async (
  groups: Map<string, TRow[]>,
  onProgress?: (done: number, total: number) => void,
) => {
  const zip = new JSZip();

  const entries = Array.from(groups.entries());
  const total = entries.length;
  let done = 0;

  // 파일이 너무 많아져도 멈춘 것처럼 보이지 않게 진행률 콜백 지원
  for (const [itemName, rows] of entries) {
    const buf = await makeCjWorkbookBuffer(rows);
    const safe = sanitizeFileName(itemName);
    zip.file(`${safe}.xlsx`, buf);

    done += 1;
    onProgress?.(done, total);
  }

  const zipBlob = await zip.generateAsync({ type: "blob" });
  saveAs(zipBlob, makeDatedFileName("한섬누리_CJ제출용_품목별엑셀.zip"));
};

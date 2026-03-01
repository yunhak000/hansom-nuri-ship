import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { ORIGINAL_ITEM_COL, ORIGINAL_BOX_COL } from "@/lib/constants/excel";
import { extractKg } from "@/lib/utils/sort";
import { makeDatedFileName } from "@/lib/utils/filename";

type TRow = Record<string, unknown>;

export type TAggregateRow = {
  itemName: string;
  totalBox: number;
  kg: number | null;
  fruitKey: string;
};

const FRUIT_KEYWORDS = [
  "천혜향",
  "한라봉",
  "레드향",
  "감귤",
  "황금향",
  "카라향",
  "청견",
  "세토카",
  "데코폰",
];

const extractFruitKey = (itemName: string) => {
  for (const k of FRUIT_KEYWORDS) {
    if (itemName.includes(k)) return k;
  }
  // 못 찾으면 일단 첫 단어로 fallback
  return itemName.split(/\s+/)[0] ?? itemName;
};

const normalizeItemNameForAggregate = (itemName: string) => {
  return (
    itemName
      // 5kg -> 4.5kg (단, 앞이 숫자/점이면 제외: 2.5kg 같은 경우)
      .replace(/(?<![\d.])5(\s*kg\b)/gi, "4.5$1")
      // 10kg -> 9kg (단, 앞이 숫자/점이면 제외)
      .replace(/(?<![\d.])10(\s*kg\b)/gi, "9$1")
  );
};

export const buildAggregateRows = (originalRows: TRow[]): TAggregateRow[] => {
  const map = new Map<string, number>();

  for (const row of originalRows) {
    const itemName = String(row[ORIGINAL_ITEM_COL] ?? "").trim();
    if (!itemName) continue;

    const rawBox = row[ORIGINAL_BOX_COL];
    const box = Number(rawBox ?? 0);

    map.set(
      itemName,
      (map.get(itemName) ?? 0) + (Number.isFinite(box) ? box : 0),
    );
  }

  const normalizeKgForAggregate = (kg: number | null): number | null => {
    if (kg === null) return null;
    if (kg === 5) return 4.5;
    if (kg === 10) return 9;
    return kg;
  };

  const result: TAggregateRow[] = Array.from(map.entries()).map(
    ([itemName, totalBox]) => {
      const rawKg = extractKg(itemName);

      return {
        itemName,
        totalBox,
        kg: normalizeKgForAggregate(rawKg),
        fruitKey: extractFruitKey(itemName),
      };
    },
  );

  // ✅ 품목(과일키) → kg 오름차순 → 품목명
  result.sort((a, b) => {
    const fk = a.fruitKey.localeCompare(b.fruitKey, "ko");
    if (fk !== 0) return fk;

    const ak = a.kg ?? Number.POSITIVE_INFINITY;
    const bk = b.kg ?? Number.POSITIVE_INFINITY;
    if (ak !== bk) return ak - bk;

    return a.itemName.localeCompare(b.itemName, "ko");
  });

  return result;
};

export const downloadAggregateExcel = async (
  aggregateRows: TAggregateRow[],
) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("품목별 집계");

  worksheet.columns = [
    { header: "품목명", key: "itemName", width: 70 },
    { header: "총 박스수량", key: "totalBox", width: 16 },
  ];

  aggregateRows.forEach((row) => {
    worksheet.addRow({
      itemName: normalizeItemNameForAggregate(row.itemName),
      totalBox: row.totalBox,
    });
  });

  worksheet.getRow(1).font = { bold: true };

  // 🔧 너가 “필터 기능 없애겠다”고 했으면 아래 autoFilter 줄은 지워도 됨.
  // 남겨도 any랑은 무관하고 동작만(엑셀 필터) 달라져.
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: aggregateRows.length + 1, column: 2 },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), makeDatedFileName("한섬누리_품목별_집계.xlsx"));
};

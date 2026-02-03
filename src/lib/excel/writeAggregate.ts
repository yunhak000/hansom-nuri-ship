import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { ORIGINAL_ITEM_COL, ORIGINAL_BOX_COL } from "@/lib/constants/excel";
import { extractKg, sortByKgThenName } from "@/lib/utils/sort";
import { makeDatedFileName } from "@/lib/utils/filename";

export type TAggregateRow = {
  itemName: string;
  totalBox: number;
  kg: number | null;
};

const normalizeItemNameForAggregate = (itemName: string) => {
  return itemName.replace(/(\b)5(\s*kg\b)/gi, "$14.5$2").replace(/(\b)10(\s*kg\b)/gi, "$19$2");
};

export const buildAggregateRows = (originalRows: Array<Record<string, any>>): TAggregateRow[] => {
  const map = new Map<string, number>();

  for (const row of originalRows) {
    const itemName = String(row[ORIGINAL_ITEM_COL] ?? "").trim();
    if (!itemName) continue;

    const rawBox = row[ORIGINAL_BOX_COL];
    const box = Number(rawBox ?? 0);

    map.set(itemName, (map.get(itemName) ?? 0) + (Number.isFinite(box) ? box : 0));
  }

  const normalizeKgForAggregate = (kg: number | null): number | null => {
    if (kg === null) return null;
    if (kg === 5) return 4.5;
    if (kg === 10) return 9;
    return kg;
  };

  const result: TAggregateRow[] = Array.from(map.entries()).map(([itemName, totalBox]) => {
    const rawKg = extractKg(itemName);

    return {
      itemName,
      totalBox,
      kg: normalizeKgForAggregate(rawKg),
    };
  });

  result.sort(sortByKgThenName);
  return result;
};

export const downloadAggregateExcel = async (aggregateRows: TAggregateRow[]) => {
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

  // 헤더 스타일
  worksheet.getRow(1).font = { bold: true };

  // AutoFilter (정렬/필터 가능)
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: aggregateRows.length + 1, column: 2 },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), makeDatedFileName("한섬누리_품목별_집계.xlsx"));
};

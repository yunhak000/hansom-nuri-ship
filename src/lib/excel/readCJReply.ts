import ExcelJS from "exceljs";

export type TReadCjReplyResult = {
  map: any;
  orderFileMap: Map<string, Set<string>>;
};

const normalize = (v: unknown) => String(v ?? "").trim();

const normalizeHeader = (v: unknown) => normalize(v).replace(/\s/g, "");

const findHeaderCol = (headers: unknown[], target: string) => {
  const t = normalizeHeader(target);
  const exact = headers.findIndex((h) => normalizeHeader(h) === t);
  if (exact !== -1) return exact;

  const partial = headers.findIndex((h) => normalizeHeader(h).includes(t));
  return partial !== -1 ? partial : -1;
};

export const readCjReplyFiles = async (files: File[]): Promise<TReadCjReplyResult> => {
  const map = new Map<string, string>();
  const orderFileMap = new Map<string, Set<string>>();

  for (const file of files) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const sheet = workbook.worksheets[0];
    if (!sheet) continue;

    // 1행을 헤더로
    const headerRow = sheet.getRow(1);
    const headers = headerRow.values as unknown[]; // 1-based

    const orderCol = findHeaderCol(headers, "고객주문번호");
    const trackingCol = findHeaderCol(headers, "운송장번호");

    const fileOrderSet = new Set<string>();

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const orderNo = normalize(row.getCell(orderCol).value); // 고객주문번호
      const tracking = normalize(row.getCell(trackingCol).value); // 운송장번호
      if (!orderNo || !tracking) return;

      // ✅ 덮어쓰기 대신 배열에 누적
      const list = map.get(orderNo) ?? ([] as any);
      list.push(tracking);
      map.set(orderNo, list);

      fileOrderSet.add(orderNo);
    });

    // 파일 간 중복 검사용
    for (const orderNo of fileOrderSet) {
      const set = orderFileMap.get(orderNo) ?? new Set<string>();
      set.add(file.name);
      orderFileMap.set(orderNo, set);
    }
  }

  return { map, orderFileMap };
};

import ExcelJS from "exceljs";

export type TReadExcelResult = {
  headers: string[];
  rows: Array<Record<string, any>>;
};

export const readExcelFile = async (file: File): Promise<TReadExcelResult> => {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("엑셀 시트를 찾을 수 없습니다.");
  }

  const headers: string[] = [];
  worksheet.getRow(1).eachCell((cell, colNumber) => {
    headers[colNumber - 1] = String(cell.value ?? "").trim();
  });

  const rows: Array<Record<string, any>> = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const rowData: Record<string, any> = {};
    headers.forEach((header, index) => {
      rowData[header] = row.getCell(index + 1).value ?? "";
    });

    rows.push(rowData);
  });

  return { headers, rows };
};

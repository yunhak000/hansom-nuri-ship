export const normalizeHeader = (s: unknown) =>
  String(s ?? "")
    .trim()
    .replace(/\s+/g, "");

const pad2 = (n: number) => String(n).padStart(2, "0");

// yyyy-mm-dd hh:mm:ss
const formatDateTime = (d: Date) => {
  const yyyy = d.getFullYear();
  const mm = pad2(d.getMonth() + 1);
  const dd = pad2(d.getDate());
  const hh = pad2(d.getHours());
  const mi = pad2(d.getMinutes());
  const ss = pad2(d.getSeconds());
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
};

export const toText = (v: unknown) => {
  if (v == null) return "";

  // ✅ ExcelJS가 cell.value를 Date로 주는 경우
  if (v instanceof Date) return formatDateTime(v);

  // ✅ ExcelJS RichText / { text: ... } 형태 지원 (기존 유지)
  if (typeof v === "object" && v !== null && "text" in v) {
    const maybe = (v as { text?: unknown }).text;
    if (maybe instanceof Date) return formatDateTime(maybe);
    return String(maybe ?? "");
  }

  // ✅ ExcelJS Formula 결과/하이퍼링크 등에서 { result: ... } 로 오는 경우(방어)
  if (typeof v === "object" && v !== null && "result" in v) {
    const maybe = (v as { result?: unknown }).result;
    if (maybe instanceof Date) return formatDateTime(maybe);
    return String(maybe ?? "");
  }

  return String(v);
};

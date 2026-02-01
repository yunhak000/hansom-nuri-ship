export const normalizeHeader = (s: unknown) =>
  String(s ?? "")
    .trim()
    .replace(/\s+/g, "");

export const toText = (v: unknown) => {
  if (v == null) return "";
  if (typeof v === "object" && "text" in (v as any)) return String((v as any).text ?? "");
  return String(v);
};

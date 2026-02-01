export const extractKg = (itemName: string): number | null => {
  const match = String(itemName).match(/(\d+(?:\.\d+)?)\s*kg/i);
  return match ? Number(match[1]) : null;
};

export const sortByKgThenName = (a: { kg: number | null; itemName: string }, b: { kg: number | null; itemName: string }) => {
  const ak = a.kg ?? Number.POSITIVE_INFINITY;
  const bk = b.kg ?? Number.POSITIVE_INFINITY;

  if (ak !== bk) return ak - bk;
  return a.itemName.localeCompare(b.itemName, "ko");
};

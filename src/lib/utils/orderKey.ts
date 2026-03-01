import {
  ORIGINAL_KEY_COL,
  ORIGINAL_FALLBACK_KEY_COL,
} from "@/lib/constants/excel";
import { toText } from "@/lib/utils/normalize";

type TRow = Record<string, unknown>;

export const getOriginalOrderKey = (row: TRow): string => {
  const key1 = toText(row[ORIGINAL_KEY_COL]); // 상품주문번호
  if (key1) return key1;

  const key2 = toText(row[ORIGINAL_FALLBACK_KEY_COL]); // ★쇼핑몰 주문번호★
  if (key2) return key2;

  return "";
};

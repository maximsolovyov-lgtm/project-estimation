// Rate Card lookup — server-side only.
import rateCard from "./rateCard.json";
import type { Catalog, RateQuery } from "../src/shared/types";

interface RateRow {
  name: string;
  role: string;
  grade: string;
  rates: (number | null)[];
}

const ROWS = rateCard.rows as RateRow[];
const TIERS = rateCard.tiers as string[];
const norm = (s: string) => (s ?? "").trim().toLowerCase();

/** Project Estimation!L4: INDEX(Fix-Price, MATCH(role & grade), MATCH(tier)); 0 when missing (IFERROR → 0). */
export function lookupRate(q: RateQuery): number {
  if (!q.role || !q.grade || !q.tier) return 0;
  const row = ROWS.find((r) => norm(r.role) === norm(q.role) && norm(r.grade) === norm(q.grade));
  const ti = TIERS.findIndex((t) => norm(t) === norm(q.tier));
  if (!row || ti < 0) return 0;
  return row.rates[ti] ?? 0;
}

export function buildCatalog(): Catalog {
  const map = new Map<string, string[]>();
  for (const r of ROWS) {
    if (!r.role) continue;
    const g = map.get(r.role) ?? [];
    if (r.grade && !g.includes(r.grade)) g.push(r.grade);
    map.set(r.role, g);
  }
  return {
    roles: [...map.entries()].map(([role, grades]) => ({ role, grades })),
    tiers: TIERS,
    defaultHolidays: rateCard.defaultHolidays as string[],
  };
}


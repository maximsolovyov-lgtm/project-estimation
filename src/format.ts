const nf = (min: number, max: number) =>
  new Intl.NumberFormat("en-US", { minimumFractionDigits: min, maximumFractionDigits: max });
const f0 = nf(0, 0);
const f1 = nf(0, 1);
const f2 = nf(0, 2);
const m2 = nf(2, 2);

export const fmt = {
  num: (v: number) => f2.format(v || 0),
  num1: (v: number) => f1.format(v || 0),
  int: (v: number) => f0.format(v || 0),
  usd: (v: number) => `$${m2.format(v || 0)}`,
  usd0: (v: number) => `$${f0.format(v || 0)}`,
  pct: (v: number) => `${f1.format((v || 0) * 100)}%`,
};

export function safeFileName(s: string): string {
  return (s || "project").replace(/[\\/:*?"<>|]+/g, "_").trim().slice(0, 80) || "project";
}

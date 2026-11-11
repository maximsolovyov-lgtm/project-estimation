// Date helpers working on UTC day numbers so results never drift with time zones.

const MS_DAY = 86_400_000;

export function isoToDay(iso: string): number | null {
  if (!iso) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
  if (!m) return null;
  return Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / MS_DAY);
}

export function dayToIso(day: number): string {
  return new Date(day * MS_DAY).toISOString().slice(0, 10);
}

export function dayToDate(day: number): Date {
  return new Date(day * MS_DAY);
}

function isWeekend(day: number): boolean {
  const dow = new Date(day * MS_DAY).getUTCDay();
  return dow === 0 || dow === 6;
}

/** Excel WORKDAY(start, days, holidays) for days >= 0 (and negative for completeness). */
export function workday(start: number, days: number, holidays: Set<number>): number {
  let d = start;
  let left = Math.trunc(days);
  const step = left >= 0 ? 1 : -1;
  left = Math.abs(left);
  while (left > 0) {
    d += step;
    if (!isWeekend(d) && !holidays.has(d)) left--;
  }
  return d;
}

/** dd.mm.yyyy for UI/presentation. */
export function fmtDate(iso: string | null | undefined): string {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y}`;
}

// Phase management: every per-phase array (engagement % and package man-days) is kept
// in step with `phases`, so adding / removing / moving a phase never misaligns data.

import { DEFAULT_PHASES } from "./factory";
import { MAX_PHASES, MIN_PHASES, type ProjectData } from "./types";

const fit = (xs: number[] | undefined, n: number) =>
  Array.from({ length: n }, (_, i) => (typeof xs?.[i] === "number" && Number.isFinite(xs[i]) ? xs[i] : 0));

/** Pads/truncates per-phase arrays to the number of phases (used when loading saved projects). */
export function normalizePhases(d: ProjectData): ProjectData {
  const phases = d.phases?.length ? d.phases : DEFAULT_PHASES.map((name) => ({ name, workingDays: 0 }));
  const n = phases.length;
  return {
    ...d,
    phases,
    allocations: d.allocations.map((a) => ({ ...a, fte: fit(a.fte, n) })),
    packages: d.packages.map((p) => ({ ...p, days: fit(p.days, n) })),
  };
}

function remap(d: ProjectData, phases: ProjectData["phases"], order: (number | null)[]): ProjectData {
  const pick = (xs: number[]) => order.map((i) => (i === null ? 0 : xs[i] ?? 0));
  return {
    ...d,
    phases,
    allocations: d.allocations.map((a) => ({ ...a, fte: pick(a.fte) })),
    packages: d.packages.map((p) => ({ ...p, days: pick(p.days) })),
  };
}

const idx = (n: number) => Array.from({ length: n }, (_, i) => i as number | null);

/** Inserts a new empty phase at `at` (default: end). */
export function addPhase(d: ProjectData, at: number = d.phases.length): ProjectData {
  if (d.phases.length >= MAX_PHASES) return d;
  const phases = [...d.phases];
  phases.splice(at, 0, { name: `Phase ${d.phases.length + 1}`, workingDays: 0 });
  const order = idx(d.phases.length);
  order.splice(at, 0, null);
  return remap(d, phases, order);
}

export function removePhase(d: ProjectData, i: number): ProjectData {
  if (d.phases.length <= MIN_PHASES) return d;
  const order = idx(d.phases.length).filter((k) => k !== i);
  return remap(d, d.phases.filter((_, k) => k !== i), order);
}

export function movePhase(d: ProjectData, i: number, dir: -1 | 1): ProjectData {
  const j = i + dir;
  if (j < 0 || j >= d.phases.length) return d;
  const order = idx(d.phases.length);
  [order[i], order[j]] = [order[j], order[i]];
  return remap(d, order.map((k) => d.phases[k!]), order);
}

/** Back to the standard 5 phases. Data of phases at the same position is kept, the rest is dropped. */
export function resetPhases(d: ProjectData): ProjectData {
  const order = DEFAULT_PHASES.map((_, i) => (i < d.phases.length ? i : null));
  const phases = DEFAULT_PHASES.map((name, i) => ({ name, workingDays: d.phases[i]?.workingDays ?? 0 }));
  return remap(d, phases, order);
}

/** True when a phase holds any engagement or man-days (used to confirm deletion). */
export function phaseHasData(d: ProjectData, i: number): boolean {
  return d.allocations.some((a) => (a.fte[i] ?? 0) !== 0) || d.packages.some((p) => (p.days[i] ?? 0) !== 0);
}

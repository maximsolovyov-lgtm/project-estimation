// Calculation engine — a 1:1 port of the formulas on the "Project Estimation" sheet.
// Excel cell references are given next to every value.

import { dayToIso, isoToDay, workday } from "./dates";
import { ALL_BLOCK, type ProjectData } from "./types";

export type CheckState = "ok" | "bad" | "none";

export interface TeamResult {
  /** N4:N13 — Σ(engagement × working days) of allocation rows with the same Role+Grade+Tier. */
  work: number;
  /** M4:M13 */
  share: number;
}

export interface CalcResult {
  phaseStart: (string | null)[]; // I18:M18
  phaseEnd: (string | null)[]; // I19:M19
  projectStart: string | null; // N18
  projectEnd: string | null; // N19
  totalWorkingDays: number; // N20
  fteByPhase: number[]; // I21:M21
  fteTotal: number; // N21
  effortByRoles: number[]; // I22:M22
  effortByRolesTotal: number; // N22
  allocationAvgFte: number[]; // N23:N34
  allocationInTeam: boolean[];
  effortByBlock: number[]; // I36:M36
  effortByBlockTotal: number; // N36
  packageTotals: number[]; // N37:N…
  uplift: number[]; // I15:M15
  upliftTotal: number; // N15
  team: TeamResult[];
  teamWorkTotal: number; // N14
  teamShareTotal: number; // M14
  workRoles: number; // F4
  contingency: number; // F5
  contingencyPct: number; // F6
  totalWork: number; // F7
  projectCost: number; // F9
  discountPct: number; // F10
  discountAmount: number; // F11
  blendedRate: number; // F12
  durationMonths: number; // F13
  costAfterDiscount: number; // F14
  blockEstimation: { block: string; value: number }[]; // P3:Q13
  blockEstimationTotal: number; // Q14
  phaseEstimation: { phase: string; value: number }[]; // S3:V7
  phaseEstimationTotal: number; // V14
  checkTotal: CheckState; // CF on F4 / N14 / N36
  checkPhases: CheckState[]; // CF on I22:M22
  checkPhasesTotal: CheckState; // CF on N22
}

const n = (v: unknown): number => (typeof v === "number" && Number.isFinite(v) ? v : 0);
const eq = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();
const r = (v: number, d: number) => Math.round(v * 10 ** d) / 10 ** d;

/** CF rule: green when equal (rounded) and not zero, red when different, neutral when both 0. */
export function check(a: number, b: number, digits = 1, zeroSide: "b" | "a" = "b"): CheckState {
  const same = r(a, digits) === r(b, digits);
  if (same && (zeroSide === "b" ? b !== 0 : a !== 0)) return "ok";
  if (!(a === 0 && b === 0) && !same) return "bad";
  return "none";
}

export function calculate(p: ProjectData): CalcResult {
  const P = p.phases.length;
  const days = Array.from({ length: P }, (_, i) => n(p.phases[i]?.workingDays));
  const hol = new Set(p.holidays.map(isoToDay).filter((d): d is number => d !== null));

  // Rows 18-19: phase start / end dates.
  const start: (number | null)[] = [];
  const end: (number | null)[] = [];
  for (let i = 0; i < P; i++) {
    // I18 = input, J18 = IF(I19="","",WORKDAY(I19,1,Holidays))
    start[i] = i === 0 ? isoToDay(p.startDate) : end[i - 1] === null ? null : workday(end[i - 1]!, 1, hol);
    // I19 = IF(OR(I18="",I20=0),"",WORKDAY(I18-1,I20,Holidays))
    end[i] = start[i] === null || days[i] === 0 ? null : workday(start[i]! - 1, days[i], hol);
  }
  const startsSet = start.filter((d): d is number => d !== null);
  const endsSet = end.filter((d): d is number => d !== null);
  const projectStartDay = startsSet.length ? Math.min(...startsSet) : 0; // MIN of empty = 0
  const projectEndDay = endsSet.length ? Math.max(...endsSet) : 0;
  const totalWorkingDays = days.reduce((a, b) => a + b, 0); // N20

  // Rows 21-34: engagement grid.
  const fteByPhase = Array(P).fill(0);
  const allocationAvgFte: number[] = [];
  for (const a of p.allocations) {
    let sp = 0;
    for (let i = 0; i < P; i++) {
      const f = n(a.fte[i]);
      fteByPhase[i] += f;
      sp += f * days[i];
    }
    allocationAvgFte.push(totalWorkingDays === 0 ? 0 : sp / totalWorkingDays); // N23
  }
  const fteTotal = allocationAvgFte.reduce((a, b) => a + b, 0); // N21
  const effortByRoles = fteByPhase.map((f, i) => f * days[i]); // I22 = SUM(I23:I34)*I20
  const effortByRolesTotal = effortByRoles.reduce((a, b) => a + b, 0); // N22

  // Rows 36-…: entry table.
  const packageTotals = p.packages.map((pk) => Array.from({ length: P }, (_, i) => n(pk.days[i])).reduce((a, b) => a + b, 0));
  const effortByBlock = Array.from({ length: P }, (_, i) => p.packages.reduce((s, pk) => s + n(pk.days[i]), 0));
  const effortByBlockTotal = packageTotals.reduce((a, b) => a + b, 0);

  // Row 15: uplift for block "All" = total / (total - All), 1 when denominator is 0.
  const allPhase = Array.from({ length: P }, (_, i) =>
    p.packages.reduce((s, pk) => s + (eq(pk.block, ALL_BLOCK) ? n(pk.days[i]) : 0), 0),
  );
  const allTotal = p.packages.reduce((s, pk, k) => s + (eq(pk.block, ALL_BLOCK) ? packageTotals[k] : 0), 0);
  const up = (total: number, all: number) => (total - all === 0 ? 1 : total / (total - all));
  const uplift = effortByBlock.map((t, i) => up(t, allPhase[i]));
  const upliftTotal = up(effortByBlockTotal, allTotal);

  // Team N4:N13 = SUMPRODUCT((C=role)*(G=grade)*(H=tier)*I23:M34*I20:M20)
  const team: TeamResult[] = p.team.map((m) => {
    if (!m.role || !m.grade || !m.tier) return { work: 0, share: 0 };
    let w = 0;
    for (const a of p.allocations) {
      if (eq(a.role, m.role) && eq(a.grade, m.grade) && eq(a.tier, m.tier)) {
        for (let i = 0; i < P; i++) w += n(a.fte[i]) * days[i];
      }
    }
    return { work: w, share: 0 };
  });
  const teamWorkTotal = team.reduce((s, t) => s + t.work, 0); // N14
  team.forEach((t) => (t.share = teamWorkTotal === 0 ? 0 : t.work / teamWorkTotal)); // M4
  const teamShareTotal = team.reduce((s, t) => s + t.share, 0);
  const allocationInTeam = p.allocations.map((a) =>
    p.team.some((m) => eq(m.role, a.role) && eq(m.grade, a.grade) && eq(m.tier, a.tier)),
  );

  // Total for the project, F4:F14.
  const workRoles = teamWorkTotal; // F4 = N14
  const contingency = n(p.contingency); // F5
  const totalWork = workRoles + contingency; // F7 = F4+F5
  const contingencyPct = totalWork === 0 ? 0 : contingency / totalWork; // F6 = F5/F7
  const sumRate = p.team.reduce((s, m, k) => s + (m.role && m.grade && m.tier ? n(m.rate) : 0) * team[k].work, 0);
  const projectCost = sumRate * (1 + contingencyPct); // F9 = SUMPRODUCT(N,L)*(100%+F6)
  const discountPct = n(p.discount); // F10
  const discountAmount = projectCost * discountPct; // F11
  const costAfterDiscount = projectCost - discountAmount; // F14
  const blendedRate = totalWork === 0 ? 0 : costAfterDiscount / totalWork; // F12
  const durationMonths = (projectEndDay - projectStartDay) / 30; // F13 = (N19-N18)/30

  // Block Estimation Q3:Q13 = SUMIF(block, N37:N…) * N15
  const blockEstimation = p.blocks
    .filter((b) => b.trim() !== "")
    .map((b) => ({
      block: b,
      value: p.packages.reduce((s, pk, k) => s + (eq(pk.block, b) ? packageTotals[k] : 0), 0) * upliftTotal,
    }));
  const blockEstimationTotal = blockEstimation.reduce((s, b) => s + b.value, 0);

  // Phase Estimation S3:V7 = SUM of each phase column of the entry table.
  // (The template has V6/V7 cross-wired to M/L; here every phase gets its own column.)
  const phaseEstimation = p.phases.map((ph, i) => ({ phase: ph.name, value: effortByBlock[i] }));
  const phaseEstimationTotal = phaseEstimation.reduce((s, x) => s + x.value, 0);

  return {
    phaseStart: start.map((d) => (d === null ? null : dayToIso(d))),
    phaseEnd: end.map((d) => (d === null ? null : dayToIso(d))),
    projectStart: startsSet.length ? dayToIso(projectStartDay) : null,
    projectEnd: endsSet.length ? dayToIso(projectEndDay) : null,
    totalWorkingDays,
    fteByPhase,
    fteTotal,
    effortByRoles,
    effortByRolesTotal,
    allocationAvgFte,
    allocationInTeam,
    effortByBlock,
    effortByBlockTotal,
    packageTotals,
    uplift,
    upliftTotal,
    team,
    teamWorkTotal,
    teamShareTotal,
    workRoles,
    contingency,
    contingencyPct,
    totalWork,
    projectCost,
    discountPct,
    discountAmount,
    blendedRate,
    durationMonths,
    costAfterDiscount,
    blockEstimation,
    blockEstimationTotal,
    phaseEstimation,
    phaseEstimationTotal,
    checkTotal: check(teamWorkTotal, effortByBlockTotal, 1),
    checkPhases: effortByRoles.map((v, i) => check(v, effortByBlock[i], 1, "a")),
    checkPhasesTotal: check(effortByRolesTotal, effortByBlockTotal, 1, "a"),
  };
}

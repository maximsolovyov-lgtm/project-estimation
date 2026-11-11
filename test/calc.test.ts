// Verifies the calculation engine against the values Excel computed in Project_Estimation_Template.xlsx.
import { describe, expect, it } from "vitest";
import { calculate } from "../src/shared/calc";
import { sampleProject } from "../src/shared/factory";
import { workday } from "../src/shared/dates";
import { isoToDay } from "../src/shared/dates";
import { lookupRate } from "../worker/rates";

function sample() {
  const p = sampleProject();
  p.team = p.team.map((m) => ({ ...m, rate: lookupRate(m) }));
  return p;
}

describe("rate card lookup (Project Estimation!L4:L13)", () => {
  it("matches the template rates", () => {
    expect(sample().team.map((m) => m.rate)).toEqual([140.1, 100, 125.1, 66.3, 68.9, 51.4, 76.9, 42.4, 95.1, 92.7]);
  });
  it("returns 0 for missing combinations", () => {
    expect(lookupRate({ role: "Security Engineer", grade: "Lead", tier: "Tier 1" })).toBe(0);
    expect(lookupRate({ role: "Nope", grade: "Lead", tier: "Tier 1" })).toBe(0);
  });
});

describe("WORKDAY", () => {
  it("skips weekends", () => {
    const d = workday(isoToDay("2026-03-31")! - 1, 10, new Set());
    expect(d).toBe(isoToDay("2026-04-13"));
  });
});

describe("calculate() reproduces the template", () => {
  const c = calculate(sample());
  it("phase dates", () => {
    expect(c.phaseStart).toEqual(["2026-03-31", "2026-04-14", "2026-05-12", "2026-06-02", "2026-06-16"]);
    expect(c.phaseEnd).toEqual(["2026-04-13", "2026-05-11", "2026-06-01", "2026-06-15", "2026-07-06"]);
    expect(c.totalWorkingDays).toBe(70);
  });
  it("engagement grid", () => {
    expect(c.fteByPhase.map((v) => +v.toFixed(6))).toEqual([7.2, 8.4, 2, 1.4, 1.6]);
    expect(c.effortByRoles.map((v) => +v.toFixed(6))).toEqual([72, 168, 30, 14, 24]);
    expect(c.effortByRolesTotal).toBeCloseTo(308);
    expect(c.fteTotal).toBeCloseTo(4.4);
    expect(c.allocationAvgFte[0]).toBeCloseTo(0.4);
  });
  it("team workload", () => {
    expect(c.team.map((t) => t.work)).toEqual([28, 43, 22.5, 50, 0, 31, 27, 17, 23.5, 25].map((v) => expect.closeTo(v, 6)));
    expect(c.teamWorkTotal).toBeCloseTo(267);
    expect(c.team[0].share).toBeCloseTo(0.10486891385767791);
  });
  it("totals for the project", () => {
    expect(c.workRoles).toBeCloseTo(267);
    expect(c.contingencyPct).toBeCloseTo(0.011111111111111112);
    expect(c.totalWork).toBeCloseTo(270);
    expect(c.projectCost).toBeCloseTo(23554.237777777773, 6);
    expect(c.discountAmount).toBeCloseTo(2355.4237777777776, 6);
    expect(c.costAfterDiscount).toBeCloseTo(21198.813999999995, 6);
    expect(c.blendedRate).toBeCloseTo(78.51412592592591, 6);
    expect(c.durationMonths).toBeCloseTo(3.2333333333333334, 6);
  });
  it("block / phase estimation", () => {
    expect(c.effortByBlockTotal).toBe(308);
    expect(c.upliftTotal).toBe(1);
    expect(c.blockEstimation).toEqual([
      { block: "SPIFF", value: 145 },
      { block: "Maps", value: 0 },
      { block: "Sales Planning", value: 133 },
      { block: "Sales Program", value: 0 },
      { block: "Integration", value: 30 },
    ]);
    expect(c.phaseEstimation.map((p) => p.value)).toEqual([72, 168, 30, 14, 24]);
  });
  it("checks (conditional formatting)", () => {
    expect(c.checkTotal).toBe("bad"); // 267 by Team vs 308 by Block → red in the template too
    expect(c.checkPhases).toEqual(["ok", "ok", "ok", "ok", "ok"]);
  });
  it("uplift for block All", () => {
    const p = sample();
    p.packages[0].block = "All"; // 39 m/d of 308 move to All
    const c2 = calculate(p);
    expect(c2.upliftTotal).toBeCloseTo(308 / (308 - 39));
    expect(c2.blockEstimationTotal).toBeCloseTo(308);
  });
});

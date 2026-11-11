// Phase management: data stays aligned with phases; calculations and exports work for any phase count.
import { describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
import { join } from "node:path";
import { calculate } from "../src/shared/calc";
import { blankProject, sampleProject } from "../src/shared/factory";
import { addPhase, movePhase, normalizePhases, removePhase, resetPhases } from "../src/shared/phases";
import { buildWorkbook } from "../src/export/excel";
import { buildDeck } from "../src/export/pptx";
import { lookupRate } from "../worker/rates";
import { DEFAULT_PHASE_COUNT, MAX_PHASES, MIN_PHASES } from "../src/shared/types";

function sample() {
  const p = sampleProject();
  p.team = p.team.map((m) => ({ ...m, rate: lookupRate(m) }));
  return p;
}

describe("phase management", () => {
  it("new projects always start with 5 phases", () => {
    expect(blankProject().phases).toHaveLength(DEFAULT_PHASE_COUNT);
    expect(sampleProject().phases).toHaveLength(DEFAULT_PHASE_COUNT);
  });

  it("adding a phase inserts an empty column everywhere and keeps totals", () => {
    const p = sample();
    const q = addPhase(p, 2);
    expect(q.phases).toHaveLength(6);
    expect(q.phases[2].workingDays).toBe(0);
    expect(q.allocations[0].fte).toEqual([1, 0.5, 0, 0.2, 0.2, 0.2]);
    expect(q.packages[0].days).toEqual([10, 15, 0, 4, 5, 5]);
    const a = calculate(p);
    const b = calculate(q);
    expect(b.costAfterDiscount).toBeCloseTo(a.costAfterDiscount, 6);
    expect(b.effortByBlockTotal).toBe(a.effortByBlockTotal);
  });

  it("removing a phase drops its column", () => {
    const q = removePhase(sample(), 0);
    expect(q.phases.map((x) => x.name)[0]).toBe("Build");
    expect(q.allocations[0].fte).toEqual([0.5, 0.2, 0.2, 0.2]);
    expect(q.packages[0].days).toEqual([15, 4, 5, 5]);
  });

  it("moving a phase moves its data", () => {
    const q = movePhase(sample(), 0, 1);
    expect(q.phases.map((x) => x.name).slice(0, 2)).toEqual(["Build", "Discovery & Design"]);
    expect(q.allocations[0].fte.slice(0, 2)).toEqual([0.5, 1]);
    expect(q.packages[0].days.slice(0, 2)).toEqual([15, 10]);
  });

  it("respects min / max", () => {
    let p = blankProject();
    while (p.phases.length > MIN_PHASES) p = removePhase(p, 0);
    expect(removePhase(p, 0).phases).toHaveLength(MIN_PHASES);
    while (p.phases.length < MAX_PHASES) p = addPhase(p);
    expect(addPhase(p).phases).toHaveLength(MAX_PHASES);
  });

  it("reset returns to the 5 standard phases", () => {
    const q = resetPhases(addPhase(addPhase(sample())));
    expect(q.phases.map((x) => x.name)).toEqual(sample().phases.map((x) => x.name));
    expect(q.packages[0].days).toEqual([10, 15, 4, 5, 5]);
  });

  it("normalizes arrays of saved projects", () => {
    const p = sample();
    p.packages[0].days = [1, 2];
    expect(normalizePhases(p).packages[0].days).toEqual([1, 2, 0, 0, 0]);
  });
});

describe("exports with a non-default number of phases", () => {
  const out = process.env.EXPORT_DIR;
  for (const variant of ["7", "3"] as const) {
    it(`builds Excel and PowerPoint with ${variant} phases`, async () => {
      let p = sample();
      if (variant === "7") {
        p = addPhase(addPhase(p));
        p.phases[5] = { name: "Stabilization", workingDays: 5 };
        p.phases[6] = { name: "Handover", workingDays: 3 };
        p.allocations[0].fte[5] = 0.5;
        p.packages[0].days[5] = 2;
        p.packages[1].days[6] = 1;
      } else {
        p = removePhase(removePhase(p, 4), 3);
      }
      const c = calculate(p);
      const wb = buildWorkbook(p, c);
      const xbuf = await wb.xlsx.writeBuffer();
      const deck = buildDeck(p, c);
      const pbuf = (await deck.write({ outputType: "nodebuffer" })) as Buffer;
      expect(xbuf.byteLength).toBeGreaterThan(5000);
      expect(pbuf.byteLength).toBeGreaterThan(20000);
      if (out) {
        writeFileSync(join(out, `phases${variant}.xlsx`), Buffer.from(xbuf as ArrayBuffer));
        writeFileSync(join(out, `phases${variant}.pptx`), pbuf);
        writeFileSync(join(out, `phases${variant}.json`), JSON.stringify(c));
      }
    });
  }
});

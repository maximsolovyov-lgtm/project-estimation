// Import from Excel: the original template and the app's own exports load back into the same project.
import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { calculate } from "../src/shared/calc";
import { sampleProject } from "../src/shared/factory";
import { addPhase } from "../src/shared/phases";
import { uniqueProjectName } from "../src/shared/naming";
import { parseWorkbook } from "../src/import/excel";
import { buildWorkbook } from "../src/export/excel";
import { lookupRate } from "../worker/rates";
import type { ProjectData } from "../src/shared/types";

const withRates = (d: ProjectData): ProjectData => ({ ...d, team: d.team.map((m) => ({ ...m, rate: lookupRate(m) })) });
const strip = (d: ProjectData) => ({
  ...d,
  team: d.team.map(({ role, grade, tier }) => ({ role, grade, tier })),
  allocations: d.allocations.map(({ block, role, grade, tier, fte }) => ({ block, role, grade, tier, fte })),
  packages: d.packages.map(({ block, packageId, name, days }) => ({ block, packageId, name: name.trim(), days })),
  assumptions: d.assumptions.map((a) => a.text),
});

describe("import of the original Project_Estimation_Template.xlsx", async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("test/fixtures/Project_Estimation_Template.xlsx");
  const { data, warnings } = parseWorkbook(wb, "Project_Estimation_Template.xlsx");
  const ref = sampleProject();

  it("reads all tables of the template", () => {
    expect(data.name).toBe("Project Estimation Template"); // no name in the file -> file name
    expect(warnings.some((w) => w.includes("no project name"))).toBe(true);
    expect(data.startDate).toBe("2026-03-31");
    expect(data.phases).toEqual(ref.phases);
    expect(data.contingency).toBe(3);
    expect(data.discount).toBe(0.1);
    expect(data.blocks).toEqual(ref.blocks);
    expect(strip(data).team).toEqual(strip(ref).team);
    expect(strip(data).allocations).toEqual(strip(ref).allocations);
    expect(strip(data).packages).toEqual(strip(ref).packages);
  });

  it("gives the same results as the template", () => {
    const c = calculate(withRates(data));
    expect(c.teamWorkTotal).toBeCloseTo(267);
    expect(c.effortByBlockTotal).toBe(308);
    expect(c.costAfterDiscount).toBeCloseTo(21198.813999999995 * 8, 5); // daily rates (hourly x 8)
    expect(c.phaseEnd[4]).toBe("2026-07-06");
  });

  it("replaces outdated holidays (2013–2014 in the template) with the defaults", () => {
    expect(data.holidays[0]).toBe("2027-01-01");
    expect(warnings.some((w) => w.includes("before the project start"))).toBe(true);
  });
});

describe("round trip: export -> import", () => {
  for (const phases of [5, 7]) {
    it(`keeps all data with ${phases} phases`, async () => {
      let p = withRates(sampleProject(["2026-05-01"]));
      p.name = "Round trip";
      p.client = "ACME";
      p.assumptions = [{ id: "a", text: "Sandbox access in week 1" }];
      while (p.phases.length < phases) {
        p = addPhase(p);
        p.phases[p.phases.length - 1].workingDays = 4;
        p.packages[0].days[p.phases.length - 1] = 2;
        p.allocations[0].fte[p.phases.length - 1] = 0.5;
      }
      const buf = await buildWorkbook(p, calculate(p)).xlsx.writeBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf as ArrayBuffer);
      const { data, warnings } = parseWorkbook(wb, "x.xlsx");
      expect(warnings).toEqual([]);
      expect(strip(data)).toEqual({ ...strip(p), holidays: ["2026-05-01"] });
      expect(calculate(withRates(data)).costAfterDiscount).toBeCloseTo(calculate(p).costAfterDiscount, 6);
    });
  }
});

describe("unique name on import", () => {
  it("adds v2, v3 …", () => {
    expect(uniqueProjectName("KOF ICM", [])).toBe("KOF ICM");
    expect(uniqueProjectName("KOF ICM", ["KOF ICM"])).toBe("KOF ICM v2");
    expect(uniqueProjectName("KOF ICM", ["kof icm", "KOF ICM v2"])).toBe("KOF ICM v3");
    expect(uniqueProjectName("KOF ICM v2", ["KOF ICM", "KOF ICM v2"])).toBe("KOF ICM v3");
  });
});

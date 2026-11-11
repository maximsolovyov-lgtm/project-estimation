// Builds the Excel workbook and the PowerPoint deck from the template example.
// Set EXPORT_DIR to also write the files to disk for manual inspection.
import { describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
import { join } from "node:path";
import { calculate } from "../src/shared/calc";
import { sampleProject } from "../src/shared/factory";
import { buildWorkbook } from "../src/export/excel";
import { buildDeck } from "../src/export/pptx";
import { lookupRate } from "../worker/rates";

const p = sampleProject(["2026-05-01"]);
p.team = p.team.map((m) => ({ ...m, rate: lookupRate(m) }));
p.assumptions = [{ id: "a1", text: "Client provides sandbox access in week 1." }];
const c = calculate(p);
const out = process.env.EXPORT_DIR;

describe("exports", () => {
  it("builds the Excel workbook without the Rate Card", async () => {
    const wb = buildWorkbook(p, c);
    expect(wb.worksheets.map((w) => w.name)).toEqual(["Project Estimation", "Holidays", "Assumptions"]);
    const ws = wb.getWorksheet("Project Estimation")!;
    expect(ws.getCell("F14").formula).toBe("F9-F11");
    expect(ws.getCell("L4").value).toBe(1120.8); // 140.1 per hour x 8
    const buf = await wb.xlsx.writeBuffer();
    expect(buf.byteLength).toBeGreaterThan(5000);
    if (out) writeFileSync(join(out, "sample.xlsx"), Buffer.from(buf as ArrayBuffer));
  });

  it("builds the PowerPoint deck", async () => {
    const deck = buildDeck(p, c);
    const buf = (await deck.write({ outputType: "nodebuffer" })) as Buffer;
    expect(buf.byteLength).toBeGreaterThan(20000);
    if (out) writeFileSync(join(out, "sample.pptx"), buf);
  });
});

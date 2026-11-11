// "Export to Excel": rebuilds the Project Estimation sheet with the same layout and live
// formulas as Project_Estimation_Template.xlsx. The Rate Card is NOT exported: the Rate
// column holds the rates of the Team members only (as values instead of INDEX/MATCH lookups).

import ExcelJS from "exceljs";
import type { CalcResult } from "../shared/calc";
import { isoToDay, dayToDate } from "../shared/dates";
import { ALL_BLOCK, type ProjectData } from "../shared/types";
import { safeFileName } from "../format";
import { holidayName } from "../shared/holidays";

const BLUE = "FF0B4FB3"; // input cells
const HEAD_FILL = "FFDCE6F1";
const TITLE_FILL = "FF1F4E79";
const SUB_FILL = "FFF2F2F2";
const GREEN = "FF00B050";
const RED = "FFFF0000";

const thin = { style: "thin" as const, color: { argb: "FFBFBFBF" } };
const border = { top: thin, left: thin, bottom: thin, right: thin };

/** 1-based column index -> Excel letters (1 -> A, 27 -> AA). */
function colL(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function d(iso: string | null): Date | string {
  const n = iso ? isoToDay(iso) : null;
  return n === null ? "" : dayToDate(n);
}

export function buildWorkbook(p: ProjectData, c: CalcResult): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Project Estimation app";
  wb.created = new Date();
  wb.calcProperties = { fullCalcOnLoad: true };

  const ws = wb.addWorksheet("Project Estimation", {
    views: [{ showGridLines: false, zoomScale: 90 }],
    properties: { defaultRowHeight: 16 },
  });
  const hs = wb.addWorksheet("Holidays", { state: "hidden" });
  const as = wb.addWorksheet("Assumptions");

  // ---------------- layout (same anchors as the template, grows when needed) ----------------
  const n = p.phases.length;
  // Phase columns start at I (as in the template); the phase TOTAL column follows the last phase.
  // Team columns stay I:N; Block / Phase Estimation move right when there are more than 5 phases.
  const PH_COLS = Array.from({ length: n }, (_, i) => colL(9 + i));
  const FP = PH_COLS[0];
  const LP = PH_COLS[n - 1];
  const TC = colL(9 + n); // template: N
  const RB = Math.max(14, 9 + n);
  const O = colL(RB + 1), P = colL(RB + 2), Q = colL(RB + 3), R = colL(RB + 4), S = colL(RB + 5), V = colL(RB + 8);

  const T = Math.max(10, p.team.length, p.blocks.length - 1, n - 1); // team rows (template: 10, rows 4-13)
  const teamA = 4;
  const teamZ = teamA + T - 1;
  const tot = teamZ + 1; // template row 14
  const u = tot + 1; // uplift row (15)
  const h = u + 1; // phase header (16, merged 16:17)
  const rStart = h + 2; // 18
  const rEnd = h + 3; // 19
  const rDays = h + 4; // 20
  const rFte = h + 5; // 21
  const rEff = h + 6; // 22
  const A = Math.max(12, p.allocations.length);
  const a1 = h + 7; // 23
  const a2 = a1 + A - 1; // 34
  const pkH = a2 + 1; // 35
  const pkT = a2 + 2; // 36
  const p1 = a2 + 3; // 37
  const K = Math.max(16, p.packages.length);
  const p2 = p1 + K - 1;
  const END = Math.max(9934, p2); // formulas cover rows added later in Excel
  const holN = Math.max(1, p.holidays.length);
  const HOL = `Holidays!$A$1:$A$${holN}`;
  const blockListEnd = u; // P3:P{u}, P{u} = "All"

  const cell = (addr: string) => ws.getCell(addr);
  const f = (addr: string, formula: string, result?: unknown) => {
    cell(addr).value = { formula, result } as ExcelJS.CellFormulaValue;
  };
  const arr = (addr: string, formula: string, result?: unknown) => {
    cell(addr).value = { formula, result, shareType: "array", ref: addr } as unknown as ExcelJS.CellFormulaValue;
  };
  const listDV = (range: string, formula: string) => {
    const [from, to] = range.split(":");
    const col = from.replace(/\d+/g, "");
    const r0 = +from.replace(/\D+/g, "");
    const r1 = +(to ?? from).replace(/\D+/g, "");
    for (let r = r0; r <= r1; r++) {
      cell(`${col}${r}`).dataValidation = { type: "list", allowBlank: true, formulae: [formula] };
    }
  };
  const input = (addr: string, value: ExcelJS.CellValue) => {
    const x = cell(addr);
    x.value = value;
    x.font = { ...(x.font ?? {}), color: { argb: BLUE } };
  };
  const head = (addr: string, text: string) => {
    const x = cell(addr);
    x.value = text;
    x.font = { bold: true };
    x.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEAD_FILL } };
    x.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    x.border = border;
  };
  const vertical = (range: string, text: string) => {
    ws.mergeCells(range);
    const x = cell(range.split(":")[0]);
    x.value = text;
    x.font = { bold: true, color: { argb: "FFFFFFFF" } };
    x.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TITLE_FILL } };
    x.alignment = { textRotation: 90, vertical: "middle", horizontal: "center", wrapText: true };
  };

  // ---------------- columns ----------------
  const widths: Record<string, number> = {
    A: 12, B: 9, C: 12, D: 10, E: 10, F: 14, G: 11, H: 9, I: 38, J: 13, K: 13, L: 13, M: 13, N: 13,
  };
  for (let k = 15; k <= RB; k++) widths[colL(k)] = 13;
  Object.assign(widths, { [O]: 5, [P]: 18, [Q]: 12, [R]: 5, [S]: 26, [colL(RB + 6)]: 4, [colL(RB + 7)]: 4, [V]: 12 });
  Object.entries(widths).forEach(([k, w]) => (ws.getColumn(k).width = w));

  // ---------------- title ----------------
  cell("B1").value = "Project Cost and Labor Estimation";
  cell("B1").font = { bold: true, size: 16, color: { argb: "FF1F4E79" } };
  cell("I1").value = "Project Name";
  cell("I1").font = { bold: true };
  cell("I1").alignment = { horizontal: "right" };
  ws.mergeCells("J1:N1");
  input("J1", p.name);
  cell("J1").font = { bold: true, size: 13, color: { argb: BLUE } };
  if (p.client) {
    cell(`${P}1`).value = "Client";
    cell(`${P}1`).font = { bold: true };
    input(`${Q}1`, p.client);
  }
  cell("L2").value = "Calculated base of roles engagement estimation";
  cell(`${P}2`).value = "Calculated base of task estimation";
  cell(`${S}2`).value = "Calculated base of task estimation";
  ["L2", `${P}2`, `${S}2`].forEach((a) => (cell(a).font = { italic: true, size: 9, color: { argb: "FF7F7F7F" } }));

  // ---------------- Total for the project (B3:G14) ----------------
  vertical(`B3:B${tot}`, "Total for the project");
  ws.mergeCells("C3:G3");
  head("C3", "Estimation");
  const lbl = (r: number, text: string, bold = false) => {
    ws.mergeCells(`C${r}:E${r}`);
    cell(`C${r}`).value = text;
    cell(`C${r}`).font = { bold };
    for (const col of ["C", "F", "G"]) cell(`${col}${r}`).border = border;
  };
  lbl(4, "Work (roles)"); f("F4", `N${tot}`, c.workRoles); cell("G4").value = "m\\d";
  lbl(5, "Contingency Reserve m/h"); input("F5", p.contingency); cell("G5").value = "m\\d";
  lbl(6, "Contingency Reserve %"); f("F6", "F5/F7", c.contingencyPct); cell("G6").value = "%";
  lbl(7, "Total", true); f("F7", "F4+F5", c.totalWork); cell("G7").value = "m\\d";
  ws.mergeCells("C8:G8"); head("C8", "Finance");
  lbl(9, "Project Cost"); f("F9", `SUMPRODUCT($N$${teamA}:$N$${teamZ},$L$${teamA}:$L$${teamZ})*(100%+F6)`, c.projectCost); cell("G9").value = "USD";
  lbl(10, "Discount %"); input("F10", p.discount); cell("G10").value = "%";
  lbl(11, "Discount*"); f("F11", "F9*F10", c.discountAmount); cell("G11").value = "USD";
  lbl(12, "Blended Rate"); f("F12", "F14/F7", c.blendedRate); cell("G12").value = "USD";
  lbl(13, "Project Duration"); f("F13", `($${TC}$${rEnd}-$${TC}$${rStart})/30`, c.durationMonths); cell("G13").value = "Months";
  lbl(14, "Project Cost after Discount", true); f("F14", "F9-F11", c.costAfterDiscount); cell("G14").value = "USD";
  cell("F7").font = { bold: true };
  cell("F14").font = { bold: true };
  cell("F14").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
  ["F4", "F5", "F7"].forEach((a) => (cell(a).numFmt = "#,##0.0"));
  ["F6", "F10"].forEach((a) => (cell(a).numFmt = "0.0%"));
  ["F9", "F11", "F12", "F14"].forEach((a) => (cell(a).numFmt = '"$"#,##0.00'));
  cell("F13").numFmt = "0.00";

  // ---------------- Team (H3:N14) ----------------
  vertical(`H3:H${tot}`, "Team");
  ["Role", "Grade", "Tier", "Rate (USD)", "%", "Work"].forEach((t, i) => head(`${"IJKLMN"[i]}3`, t));
  const cAllocRole = `$C$${a1}:$C$${a2}`;
  const cAllocGrade = `$G$${a1}:$G$${a2}`;
  const cAllocTier = `$H$${a1}:$H$${a2}`;
  for (let k = 0; k < T; k++) {
    const r = teamA + k;
    const m = p.team[k];
    if (m) {
      input(`I${r}`, m.role);
      input(`J${r}`, m.grade);
      input(`K${r}`, m.tier);
      cell(`L${r}`).value = m.role && m.grade && m.tier ? m.rate : 0; // Rate Card is not exported
    } else cell(`L${r}`).value = 0;
    f(`M${r}`, `IF($N$${tot}=0,0,N${r}/$N$${tot})`, c.team[k]?.share ?? 0);
    arr(
      `N${r}`,
      `IF(OR($I${r}="",$J${r}="",$K${r}=""),0,SUMPRODUCT((${cAllocRole}=$I${r})*(${cAllocGrade}=$J${r})*(${cAllocTier}=$K${r})*$I$${a1}:$${LP}$${a2}*$I$${rDays}:$${LP}$${rDays}))`,
      c.team[k]?.work ?? 0,
    );
    for (const col of "IJKLMN") cell(`${col}${r}`).border = border;
    cell(`L${r}`).numFmt = "#,##0.0";
    cell(`M${r}`).numFmt = "0.0%";
    cell(`N${r}`).numFmt = "#,##0.0";
  }
  listDV(`K${teamA}:K${teamZ}`, '"Tier 1,Tier 2,Tier 3,Tier 4,Tier 5"');
  cell(`I${tot}`).value = "Total";
  f(`M${tot}`, `SUM(M${teamA}:M${teamZ})`, c.teamShareTotal);
  f(`N${tot}`, `SUM(N${teamA}:N${teamZ})`, c.teamWorkTotal);
  for (const col of "IJKLMN") {
    cell(`${col}${tot}`).font = { bold: true };
    cell(`${col}${tot}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUB_FILL } };
    cell(`${col}${tot}`).border = border;
  }
  cell(`M${tot}`).numFmt = "0.0%";
  cell(`N${tot}`).numFmt = "#,##0.0";

  // ---------------- Block Estimation (template: O3:Q14) ----------------
  vertical(`${O}3:${O}${tot}`, "Block Estimation");
  for (let r = 3; r < tot; r++) {
    const b = p.blocks[r - 3];
    if (b !== undefined) input(`${P}${r}`, b);
    const est = c.blockEstimation.find((x) => x.block === b)?.value ?? 0;
    f(`${Q}${r}`, `IF($${P}${r}="",0,SUMIF($A$${p1}:$A$${END},$${P}${r},$${TC}$${p1}:$${TC}$${END})*$${TC}$${u})`, b ? est : 0);
    cell(`${P}${r}`).border = border;
    cell(`${Q}${r}`).border = border;
    cell(`${Q}${r}`).numFmt = "#,##0.0";
  }
  f(`${Q}${tot}`, `SUM(${Q}3:${Q}${tot - 1})`, c.blockEstimationTotal);
  cell(`${Q}${tot}`).font = { bold: true };
  cell(`${Q}${tot}`).numFmt = "#,##0.0";
  cell(`${P}${u}`).value = ALL_BLOCK;
  cell(`${P}${u}`).font = { italic: true };

  // ---------------- Phase Estimation (template: R3:V14) ----------------
  vertical(`${R}3:${R}${tot}`, "Phase Estimation");
  for (let i = 0; i < n; i++) {
    const r = 3 + i;
    const col = PH_COLS[i];
    f(`${S}${r}`, `${col}$${h}`, p.phases[i].name);
    f(`${V}${r}`, `SUM(${col}$${p1}:${col}$${END})`, c.phaseEstimation[i]?.value ?? 0);
    cell(`${S}${r}`).border = border;
    cell(`${V}${r}`).border = border;
    cell(`${V}${r}`).numFmt = "#,##0.0";
  }
  f(`${V}${tot}`, `SUM(${V}3:${V}${tot - 1})`, c.phaseEstimationTotal);
  cell(`${V}${tot}`).font = { bold: true };
  cell(`${V}${tot}`).numFmt = "#,##0.0";

  // ---------------- Uplift row ----------------
  cell(`D${u}`).value = 'Uplift factor for block "All"';
  cell(`D${u}`).font = { italic: true, color: { argb: "FF7F7F7F" } };
  for (const [i, col] of [...PH_COLS, TC].entries()) {
    const rng = `${col}$${p1}:${col}$${END}`;
    const all = `SUMIF($A$${p1}:$A$${END},$${P}$${u},${rng})`;
    f(`${col}${u}`, `IF(SUM(${rng})-${all}=0,1,SUM(${rng})/(SUM(${rng})-${all}))`, i < n ? c.uplift[i] : c.upliftTotal);
    cell(`${col}${u}`).numFmt = "0.00";
    cell(`${col}${u}`).font = { italic: true, color: { argb: "FF7F7F7F" } };
  }

  // ---------------- Phases header ----------------
  ws.mergeCells(`A${h}:B${h + 1}`); head(`A${h}`, "Block");
  ws.mergeCells(`C${h}:F${h + 1}`); head(`C${h}`, "Role");
  ws.mergeCells(`G${h}:G${h + 1}`); head(`G${h}`, "Grade");
  ws.mergeCells(`H${h}:H${h + 1}`); head(`H${h}`, "Tier");
  PH_COLS.forEach((col, i) => {
    ws.mergeCells(`${col}${h}:${col}${h + 1}`);
    head(`${col}${h}`, p.phases[i].name);
    cell(`${col}${h}`).font = { bold: true, color: { argb: BLUE } };
  });
  ws.mergeCells(`${TC}${h}:${TC}${h + 1}`); head(`${TC}${h}`, "TOTAL");

  const rowLabel = (r: number, text: string, bold = false) => {
    ws.mergeCells(`A${r}:H${r}`);
    cell(`A${r}`).value = text;
    cell(`A${r}`).font = { bold };
    for (const col of ["A", ...PH_COLS, TC]) cell(`${col}${r}`).border = border;
  };
  rowLabel(rStart, "Phase Start Date");
  rowLabel(rEnd, "Phase End Date");
  rowLabel(rDays, "Amount of working days");
  rowLabel(rFte, "FTE, total");
  rowLabel(rEff, "Total efforts amount by Roles", true);
  PH_COLS.forEach((col, i) => {
    const prev = PH_COLS[i - 1];
    if (i === 0) input(`${col}${rStart}`, d(p.startDate));
    else f(`${col}${rStart}`, `IF(${prev}${rEnd}="","",WORKDAY(${prev}${rEnd},1,${HOL}))`, d(c.phaseStart[i]));
    f(`${col}${rEnd}`, `IF(OR(${col}${rStart}="",${col}${rDays}=0),"",WORKDAY(${col}${rStart}-1,${col}${rDays},${HOL}))`, d(c.phaseEnd[i]));
    input(`${col}${rDays}`, p.phases[i].workingDays);
    f(`${col}${rFte}`, `SUM(${col}${a1}:${col}${a2})`, c.fteByPhase[i]);
    f(`${col}${rEff}`, `SUM(${col}${a1}:${col}${a2})*${col}$${rDays}`, c.effortByRoles[i]);
  });
  f(`${TC}${rStart}`, `MIN(${FP}${rStart}:${LP}${rStart})`, d(c.projectStart));
  f(`${TC}${rEnd}`, `MAX(${FP}${rEnd}:${LP}${rEnd})`, d(c.projectEnd));
  f(`${TC}${rDays}`, `SUM(${FP}${rDays}:${LP}${rDays})`, c.totalWorkingDays);
  f(`${TC}${rFte}`, `SUM(${TC}${a1}:${TC}${a2})`, c.fteTotal);
  f(`${TC}${rEff}`, `SUM(${FP}${rEff}:${LP}${rEff})`, c.effortByRolesTotal);
  for (const col of [...PH_COLS, TC]) {
    cell(`${col}${rStart}`).numFmt = "dd.mm.yyyy";
    cell(`${col}${rEnd}`).numFmt = "dd.mm.yyyy";
    cell(`${col}${rFte}`).numFmt = "0.00";
    cell(`${col}${rEff}`).numFmt = "#,##0.0";
    cell(`${col}${rEff}`).font = { bold: true };
    for (const r of [rStart, rEnd, rDays, rFte, rEff]) cell(`${col}${r}`).alignment = { horizontal: "center" };
  }

  // ---------------- Allocation grid ----------------
  for (let k = 0; k < A; k++) {
    const r = a1 + k;
    const a = p.allocations[k];
    ws.mergeCells(`A${r}:B${r}`);
    ws.mergeCells(`C${r}:F${r}`);
    if (a) {
      input(`A${r}`, a.block);
      input(`C${r}`, a.role);
      input(`G${r}`, a.grade);
      input(`H${r}`, a.tier);
      PH_COLS.forEach((col, i) => a.fte[i] && input(`${col}${r}`, a.fte[i]));
    }
    f(`${TC}${r}`, `IF($${TC}$${rDays}=0,0,SUMPRODUCT($${FP}${r}:$${LP}${r},${FP}$${rDays}:${LP}$${rDays})/$${TC}$${rDays})`, c.allocationAvgFte[k] ?? 0);
    for (const col of ["A", "C", "G", "H", ...PH_COLS, TC]) cell(`${col}${r}`).border = border;
    for (const col of PH_COLS) {
      cell(`${col}${r}`).numFmt = "0%";
      cell(`${col}${r}`).font = { color: { argb: BLUE } };
      cell(`${col}${r}`).alignment = { horizontal: "center" };
    }
    cell(`${TC}${r}`).numFmt = "0.00";
  }
  listDV(`A${a1}:A${a2}`, `$${P}$3:$${P}$${blockListEnd}`);
  listDV(`C${a1}:C${a2}`, `$I$${teamA}:$I$${teamZ}`);
  listDV(`G${a1}:G${a2}`, `$J$${teamA}:$J$${teamZ}`);
  listDV(`H${a1}:H${a2}`, `$K$${teamA}:$K$${teamZ}`);

  // ---------------- Entry table ----------------
  ws.mergeCells(`A${pkH}:B${pkH}`); head(`A${pkH}`, "Block");
  head(`C${pkH}`, "Package ID");
  ws.mergeCells(`D${pkH}:H${pkH}`); head(`D${pkH}`, "Name of the service package");
  PH_COLS.forEach((col, i) => { head(`${col}${pkH}`, ""); f(`${col}${pkH}`, `${col}$${h}`, p.phases[i].name); });
  head(`${TC}${pkH}`, "Total, m/d");
  ws.mergeCells(`A${pkT}:H${pkT}`);
  cell(`A${pkT}`).value = "Total efforts amount by Block";
  cell(`A${pkT}`).font = { bold: true };
  PH_COLS.forEach((col, i) => f(`${col}${pkT}`, `SUM(${col}${p1}:${col}${END})`, c.effortByBlock[i]));
  f(`${TC}${pkT}`, `SUM(${TC}${p1}:${TC}${END})`, c.effortByBlockTotal);
  for (const col of ["A", ...PH_COLS, TC]) {
    cell(`${col}${pkT}`).border = border;
    cell(`${col}${pkT}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUB_FILL } };
    if (col !== "A") {
      cell(`${col}${pkT}`).numFmt = "#,##0.0";
      cell(`${col}${pkT}`).font = { bold: true };
      cell(`${col}${pkT}`).alignment = { horizontal: "center" };
    }
  }
  for (let k = 0; k < K; k++) {
    const r = p1 + k;
    const pk = p.packages[k];
    ws.mergeCells(`A${r}:B${r}`);
    ws.mergeCells(`D${r}:H${r}`);
    if (pk) {
      input(`A${r}`, pk.block);
      input(`C${r}`, pk.packageId);
      input(`D${r}`, pk.name);
    }
    PH_COLS.forEach((col, i) => input(`${col}${r}`, pk ? pk.days[i] || 0 : 0));
    f(`${TC}${r}`, `SUM(${FP}${r}:${LP}${r})`, c.packageTotals[k] ?? 0);
    for (const col of ["A", "C", "D", ...PH_COLS, TC]) cell(`${col}${r}`).border = border;
    for (const col of [...PH_COLS, TC]) cell(`${col}${r}`).alignment = { horizontal: "center" };
  }
  listDV(`A${p1}:A${p2}`, `$${P}$3:$${P}$${blockListEnd}`);

  // ---------------- Conditional formatting (checks) ----------------
  const okBad = (ref: string, okF: string, badF: string) =>
    ws.addConditionalFormatting({
      ref,
      rules: [
        { type: "expression", priority: 1, formulae: [okF], style: { font: { color: { argb: GREEN }, bold: true } } },
        { type: "expression", priority: 2, formulae: [badF], style: { font: { color: { argb: RED }, bold: true } } },
      ],
    });
  const N14 = `$N$${tot}`;
  const N36 = `$${TC}$${pkT}`;
  okBad("F4", `AND(${N36}<>0,ROUND(${N14},1)=ROUND(${N36},1))`, `AND(NOT(AND(${N14}=0,${N36}=0)),ROUND(${N14},1)<>ROUND(${N36},1))`);
  okBad(`N${tot}`, `AND(${N36}<>0,ROUND(${N14},2)=ROUND(${N36},2))`, `AND(NOT(AND(${N14}=0,${N36}=0)),ROUND(${N14},2)<>ROUND(${N36},2))`);
  okBad(`${TC}${pkT}`, `AND(${N36}<>0,ROUND(${N14},1)=ROUND(${N36},1))`, `AND(NOT(AND(${N14}=0,${N36}=0)),ROUND(${N14},1)<>ROUND(${N36},1))`);
  okBad(`${FP}${rEff}:${TC}${rEff}`, `AND(${FP}${rEff}<>0,ROUND(${FP}${rEff},1)=ROUND(${FP}${pkT},1))`, `AND(NOT(AND(${FP}${rEff}=0,${FP}${pkT}=0)),ROUND(${FP}${rEff},1)<>ROUND(${FP}${pkT},1))`);
  okBad(`${FP}${pkT}:${LP}${pkT}`, `AND(${FP}${pkT}<>0,ROUND(${FP}${rEff},1)=ROUND(${FP}${pkT},1))`, `AND(NOT(AND(${FP}${rEff}=0,${FP}${pkT}=0)),ROUND(${FP}${rEff},1)<>ROUND(${FP}${pkT},1))`);

  // ---------------- Instructions (same texts as the template) ----------------
  const note = (range: string, text: string, fill: string) => {
    ws.mergeCells(range);
    const x = cell(range.split(":")[0]);
    x.value = text;
    x.alignment = { wrapText: true, vertical: "top" };
    x.font = { size: 9 };
    x.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
  };
  note(`${P}${rStart}:${Q}${pkT - 1}`,
    "HOW TO USE\n\nBlue text = fill in manually.\nBlack text = calculated, do not overwrite.\n\nFILL IN:\n1. Team - Role, Grade and Tier for each team member (Rate was taken from the Rate Card, Fix-Price / A-la-Carte, when the estimate was exported).\n2. Contingency Reserve (m/h).\n3. Discount %.\n4. Phases - start date of the FIRST phase only, plus the number of working days for each phase.\n5. Allocation grid - pick Block, then Role, Grade and Tier (only Team members are offered), then enter the engagement % per phase (100% = 1 FTE, 200% = 2 FTE).\n6. Entry table (bottom) - pick Block, enter Package ID, package name and hours per phase.",
    "FFEAF1FB");
  note(`${S}${rStart}:${V}${pkT - 1}`,
    "CALCULATED AUTOMATICALLY\n\n- Phase End Date, and the start date of every phase after the first (weekends and the dates on the Holidays tab are skipped).\n- FTE, total - sum of the engagement entered for that phase.\n- Total efforts amount by Roles = engagement x working days.\n- Total efforts amount by Block = sum of the entry table.\n- Workload of specialists - effort per role and its % share.\n- Block Estimation and Phase Estimation.\n- Total for the project - Work, Total, Project Cost, Discount, Project Cost after Discount, Project Duration.",
    "FFF2F2F2");
  note(`${P}${p1}:${Q}${p1 + 15}`,
    "CHECKS\n\n- In the two \"Total efforts amount\" rows the font turns GREEN when the by-Roles and by-Block figures agree and RED when they differ. Black means both sides are empty.\n- Block \"All\" - hours booked to the block \"All\" are not reported on their own; they are spread across the other blocks in proportion to each block's own hours (uplift factor row above the phase headers).",
    "FFFFF2CC");
  note(`${S}${p1}:${V}${p1 + 15}`,
    "HOW TO FILL IN - FOLLOW THIS ORDER\n\n1. Team - Role, Grade and Tier for each team member.\n2. Project start date - start date of the FIRST phase only.\n3. Efforts by block - in the bottom table pick a Block and enter hours per phase.\n4. Duration - number of working days for each phase.\n5. Roles and engagement - pick Block, then Role, Grade and Tier from the Team, then enter the engagement % per phase (100% = 1 FTE, 200% = 2 FTE).\n6. Check - \"Total efforts amount by Roles\" must equal \"by Block\": the font turns green when they match, red when they differ.\n7. Contingency Reserve in m/h.\n8. Discount %.\n\nBlue text = you fill in. Black text = calculated, do not overwrite.",
    "FFEAF1FB");

  ws.views = [{ state: "frozen", xSplit: 0, ySplit: 0, showGridLines: false, zoomScale: 90 }];
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };

  // ---------------- Holidays (hidden) ----------------
  p.holidays.forEach((iso, i) => {
    const c2 = hs.getCell(`A${i + 1}`);
    c2.value = d(iso);
    c2.numFmt = "dd.mm.yyyy";
    hs.getCell(`B${i + 1}`).value = holidayName(iso);
  });
  hs.getColumn("A").width = 14;
  hs.getColumn("B").width = 34;

  // ---------------- Assumptions ----------------
  as.getColumn("A").width = 6;
  as.getColumn("B").width = 110;
  as.getCell("A1").value = "Assumptions";
  as.getCell("A1").font = { bold: true, size: 14, color: { argb: "FF1F4E79" } };
  as.getCell("A3").value = "#";
  as.getCell("B3").value = "Assumption";
  ["A3", "B3"].forEach((a) => {
    as.getCell(a).font = { bold: true };
    as.getCell(a).fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEAD_FILL } };
  });
  p.assumptions.forEach((x, i) => {
    as.getCell(`A${i + 4}`).value = i + 1;
    as.getCell(`B${i + 4}`).value = x.text;
    as.getCell(`B${i + 4}`).alignment = { wrapText: true, vertical: "top" };
  });

  return wb;
}

export async function exportToExcel(p: ProjectData, c: CalcResult): Promise<void> {
  const wb = buildWorkbook(p, c);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${safeFileName(p.name)} - Estimation.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

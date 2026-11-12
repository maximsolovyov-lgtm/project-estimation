// "Import from Excel": reads a filled-in Project_Estimation_Template.xlsx (or a file exported
// by this app, with any number of phases) back into a project.
// Tables are located by their labels, not by fixed addresses, so grown/shifted layouts work.
// Rates in the file are ignored: the server resolves them from the hidden Rate Card.

import type ExcelJS from "exceljs";
import { uid } from "../shared/factory";
import { DEFAULT_HOLIDAYS } from "../shared/holidays";
import { ALL_BLOCK, MAX_PHASES, type Allocation, type ProjectData, type TeamMember, type WorkPackage } from "../shared/types";

export interface ImportResult {
  data: ProjectData;
  warnings: string[];
}

type V = ExcelJS.CellValue;

/** Unwraps formula results, rich text and hyperlinks. */
function raw(v: V): unknown {
  if (v && typeof v === "object" && !(v instanceof Date)) {
    const o = v as unknown as Record<string, unknown>;
    if ("result" in o) return raw(o.result as V);
    if ("formula" in o || "sharedFormula" in o) return null;
    if (Array.isArray(o.richText)) return (o.richText as { text: string }[]).map((t) => t.text).join("");
    if ("text" in o) return o.text;
    if ("error" in o) return null;
  }
  return v;
}

function text(v: V): string {
  const r = raw(v);
  if (r === null || r === undefined) return "";
  if (r instanceof Date) return r.toISOString().slice(0, 10);
  return String(r).trim();
}

function num(v: V): number {
  const r = raw(v);
  if (typeof r === "number" && Number.isFinite(r)) return r;
  if (typeof r === "string") {
    const s = r.replace(/[$,\s]/g, "");
    const pct = s.endsWith("%");
    const n = Number(pct ? s.slice(0, -1) : s);
    if (Number.isFinite(n)) return pct ? n / 100 : n;
  }
  return 0;
}

function isoDate(v: V): string {
  const r = raw(v);
  if (r instanceof Date && !Number.isNaN(r.getTime())) return r.toISOString().slice(0, 10);
  if (typeof r === "number" && r > 20000 && r < 80000) return new Date(Math.round((r - 25569) * 86_400_000)).toISOString().slice(0, 10);
  if (typeof r === "string" && /^\d{4}-\d{2}-\d{2}/.test(r)) return r.slice(0, 10);
  if (typeof r === "string") {
    const m = /^(\d{1,2})[./](\d{1,2})[./](\d{4})$/.exec(r.trim());
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }
  return "";
}

const norm = (s: string) => s.trim().toLowerCase();

export function parseWorkbook(wb: ExcelJS.Workbook, fileName = ""): ImportResult {
  const warnings: string[] = [];
  const ws = wb.getWorksheet("Project Estimation") ?? wb.worksheets[0];
  if (!ws) throw new Error("The file has no worksheets");
  const lastRow = Math.max(ws.rowCount, ws.actualRowCount);
  const lastCol = Math.max(ws.columnCount, 26);
  const at = (r: number, c: number) => ws.getRow(r).getCell(c).value;

  /** First cell (row, col) whose text matches. */
  const find = (pred: (t: string) => boolean, rows = lastRow, cols = lastCol): [number, number] | null => {
    for (let r = 1; r <= rows; r++) {
      for (let c = 1; c <= cols; c++) if (pred(norm(text(at(r, c))))) return [r, c];
    }
    return null;
  };
  const findInCol = (col: number, pred: (t: string) => boolean, from = 1): number | null => {
    for (let r = from; r <= lastRow; r++) if (pred(norm(text(at(r, col))))) return r;
    return null;
  };
  const rightOf = (r: number, c: number) => {
    for (let k = c + 1; k <= c + 8; k++) {
      const t = text(at(r, k));
      if (t) return t;
    }
    return "";
  };

  // ---------- header: project name / client ----------
  const nameCell = find((t) => t === "project name", 3);
  let name = nameCell ? rightOf(nameCell[0], nameCell[1]) : "";
  if (!name) {
    name = fileName.replace(/\.(xlsx|xlsm|xls)$/i, "").replace(/ - Estimation$/i, "").replace(/_/g, " ").trim() || "Imported project";
    warnings.push(`The file has no project name; used "${name}".`);
  }
  const clientCell = find((t) => t === "client", 3);
  const client = clientCell ? rightOf(clientCell[0], clientCell[1]) : "";

  // ---------- Total for the project: contingency & discount (labels in C, values in F) ----------
  const labelValue = (pred: (t: string) => boolean): number | null => {
    const hit = find(pred, 40, 8);
    if (!hit) return null;
    for (let c = hit[1] + 1; c <= hit[1] + 6; c++) {
      const v = raw(at(hit[0], c));
      if (typeof v === "number") return v;
      if (typeof v === "string" && /^[\s$\d.,%-]+$/.test(v) && /\d/.test(v)) return num(v as V);
    }
    return null;
  };
  const contingency = labelValue((t) => t.startsWith("contingency reserve m")) ?? 0;
  let discount = labelValue((t) => t === "discount %") ?? 0;
  if (discount > 1) discount = discount / 100; // typed as 10 instead of 10%

  // ---------- phases ----------
  const startRow = findInCol(1, (t) => t === "phase start date");
  if (!startRow) throw new Error('Not a project estimation file: "Phase Start Date" row not found');
  const hdrRow = (() => {
    for (let r = startRow - 1; r >= 1; r--) if (norm(text(at(r, 1))) === "block") return r;
    return startRow - 2;
  })();
  const endRow = findInCol(1, (t) => t === "phase end date", startRow) ?? startRow + 1;
  const daysRow = findInCol(1, (t) => t === "amount of working days", startRow) ?? startRow + 2;
  const effRow = findInCol(1, (t) => t === "total efforts amount by roles", startRow) ?? startRow + 4;
  const firstPh = 9; // column I, as in the template
  const phaseCols: number[] = [];
  for (let c = firstPh; c <= lastCol && phaseCols.length < MAX_PHASES; c++) {
    const t = text(at(hdrRow, c));
    if (norm(t) === "total" || (!t && !text(at(daysRow, c)))) break;
    phaseCols.push(c);
  }
  if (!phaseCols.length) throw new Error("No phases found in the phase header row");
  const phases = phaseCols.map((c, i) => ({ name: text(at(hdrRow, c)) || `Phase ${i + 1}`, workingDays: Math.round(num(at(daysRow, c))) }));
  const startDate = isoDate(at(startRow, firstPh));
  if (!startDate) warnings.push("Project start date not found; please set it in the Phases block.");
  void endRow;

  // ---------- team (header "Role / Grade / Tier" above the summary) ----------
  const teamHdr = find((t) => t === "role", hdrRow - 1);
  const team: TeamMember[] = [];
  if (teamHdr) {
    const [hr, rc] = teamHdr;
    for (let r = hr + 1; r < hdrRow; r++) {
      const role = text(at(r, rc));
      if (norm(role) === "total") break;
      const grade = text(at(r, rc + 1));
      const tier = text(at(r, rc + 2));
      if (!role && !grade && !tier) continue;
      team.push({ id: uid(), role, grade, tier, rate: 0 });
    }
  }
  if (!team.length) warnings.push("No Team members found.");

  // ---------- blocks (list ending with "All" on the uplift row) ----------
  const blocks: string[] = [];
  const upl = find((t) => t.startsWith("uplift factor for block"), hdrRow);
  if (upl) {
    let blockCol = 0;
    for (let c = 1; c <= lastCol; c++) if (norm(text(at(upl[0], c))) === norm(ALL_BLOCK)) blockCol = c;
    if (blockCol) {
      for (let r = 3; r < upl[0]; r++) {
        const b = text(at(r, blockCol));
        if (b && norm(b) !== norm(ALL_BLOCK) && !blocks.includes(b)) blocks.push(b);
      }
    }
  }

  // ---------- allocation grid (between "by Roles" and the package table header) ----------
  const pkHdr = (() => {
    for (let r = effRow + 1; r <= lastRow; r++) if (norm(text(at(r, 1))) === "block" && norm(text(at(r, 3))) === "package id") return r;
    return null;
  })();
  if (!pkHdr) throw new Error('Service package table ("Block / Package ID") not found');
  const allocations: Allocation[] = [];
  for (let r = effRow + 1; r < pkHdr; r++) {
    const block = text(at(r, 1));
    const role = text(at(r, 3));
    const grade = text(at(r, 7));
    const tier = text(at(r, 8));
    const fte = phaseCols.map((c) => num(at(r, c)));
    if (!block && !role && fte.every((x) => !x)) continue;
    allocations.push({ id: uid(), block, role, grade, tier, fte });
  }

  // ---------- service packages ----------
  const pkTotal = findInCol(1, (t) => t === "total efforts amount by block", pkHdr) ?? pkHdr + 1;
  const packages: WorkPackage[] = [];
  for (let r = pkTotal + 1; r <= lastRow; r++) {
    const block = text(at(r, 1));
    const packageId = text(at(r, 3));
    const name4 = text(at(r, 4));
    const days = phaseCols.map((c) => num(at(r, c)));
    if (!block && !packageId && !name4 && days.every((x) => !x)) continue;
    packages.push({ id: uid(), block, packageId, name: name4, days });
  }

  // blocks referenced in the tables but missing from the list
  for (const b of [...allocations.map((a) => a.block), ...packages.map((p) => p.block)]) {
    if (b && norm(b) !== norm(ALL_BLOCK) && !blocks.some((x) => norm(x) === norm(b))) blocks.push(b);
  }

  // link engagement rows to team members
  for (const a of allocations) {
    const m = team.find((t) => norm(t.role) === norm(a.role) && norm(t.grade) === norm(a.grade) && norm(t.tier) === norm(a.tier));
    if (m) a.memberId = m.id;
  }

  // ---------- holidays ----------
  let holidays: string[] = [];
  const hs = wb.getWorksheet("Holidays");
  if (hs) {
    hs.eachRow((row) => {
      const iso = isoDate(row.getCell(1).value);
      if (iso && !holidays.includes(iso)) holidays.push(iso);
    });
  }
  const relevant = holidays.filter((h) => !startDate || h >= startDate);
  if (!relevant.length) {
    if (holidays.length) warnings.push(`The file's ${holidays.length} holidays are all before the project start; the default US 2027–2028 holidays are used instead.`);
    holidays = DEFAULT_HOLIDAYS.map((h) => h.date);
  }
  holidays.sort();

  // ---------- assumptions ----------
  const assumptions: ProjectData["assumptions"] = [];
  const as = wb.worksheets.find((w) => /^assum/i.test(w.name));
  if (as) {
    const ours = norm(text(as.getCell("B3").value)) === "assumption";
    as.eachRow((row, r) => {
      if (ours) {
        if (r > 3 && text(row.getCell(2).value)) assumptions.push({ id: uid(), text: text(row.getCell(2).value) });
        return;
      }
      const parts: string[] = [];
      row.eachCell((cell) => {
        const t = text(cell.value);
        if (t) parts.push(t);
      });
      if (parts.length && norm(parts.join(" ")) !== "assumptions") assumptions.push({ id: uid(), text: parts.join(" · ") });
    });
  }

  return {
    data: {
      name,
      client,
      startDate: startDate || new Date().toISOString().slice(0, 10),
      phases,
      contingency,
      discount,
      blocks: blocks.length ? blocks : [""],
      team: team.length ? team : [{ id: uid(), role: "", grade: "", tier: "", rate: 0 }],
      allocations,
      packages,
      holidays,
      assumptions,
    },
    warnings,
  };
}

/** Browser entry point: reads the File with ExcelJS (lazy-loaded with the import code). */
export async function importExcelFile(file: File): Promise<ImportResult> {
  const { default: Excel } = await import("exceljs");
  const wb = new Excel.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  return parseWorkbook(wb, file.name);
}

// "Сформировать презентацию": PowerPoint deck where every element of the estimate gets its own slide.

import PptxGenJS from "pptxgenjs";
import type { CalcResult, CheckState } from "../shared/calc";
import { fmtDate, isoToDay, dayToIso } from "../shared/dates";
import { PHASE_COUNT, type ProjectData } from "../shared/types";
import { fmt, safeFileName } from "../format";

const BRAND = "1F4E79";
const ACCENT = "2E75B6";
const INK = "1D2733";
const MUTED = "6B7785";
const LINE = "D9E1EA";
const OK = "138A36";
const BAD = "C62828";
const PHASE_COLORS = ["2E75B6", "F0A030", "4CAF7A", "8E6FC8", "E0607E"];
const SERIES = ["2E75B6", "F0A030", "4CAF7A", "8E6FC8", "E0607E", "26A6A6", "A3794C", "5C6BC0", "D4A017", "7D8C99"];
const W = 13.333;
const H = 7.5;
const FONT = "Segoe UI";

type Slide = PptxGenJS.Slide;
type Row = PptxGenJS.TableRow;

const hdr = (text: string, align: "left" | "right" | "center" = "left"): PptxGenJS.TableCell => ({
  text,
  options: { bold: true, color: "FFFFFF", fill: { color: BRAND }, align, fontSize: 11 },
});
const cellT = (text: string, opts: PptxGenJS.TableCellProps = {}): PptxGenJS.TableCell => ({
  text,
  options: { fontSize: 11, color: INK, ...opts },
});
const num = (text: string, opts: PptxGenJS.TableCellProps = {}) => cellT(text, { align: "right", ...opts });
const totalRow = (opts: PptxGenJS.TableCellProps = {}): PptxGenJS.TableCellProps => ({ bold: true, fill: { color: "EEF3F9" }, ...opts });
const checkColor = (s: CheckState) => (s === "ok" ? OK : s === "bad" ? BAD : INK);

export function buildDeck(p: ProjectData, c: CalcResult): PptxGenJS {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = `${p.name} — Estimation`;
  if (p.client) pptx.company = p.client;
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };

  pptx.defineSlideMaster({
    title: "CONTENT",
    background: { color: "FFFFFF" },
    objects: [
      { rect: { x: 0, y: 0, w: W, h: 0.12, fill: { color: BRAND } } },
      { line: { x: 0.5, y: 7.0, w: W - 1, h: 0, line: { color: LINE, width: 1 } } },
      { text: { text: p.name, options: { x: 0.5, y: 7.05, w: 9, h: 0.3, fontSize: 9, color: MUTED, fontFace: FONT } } },
    ],
    slideNumber: { x: W - 1.0, y: 7.05, w: 0.5, h: 0.3, fontSize: 9, color: MUTED, fontFace: FONT },
  });

  const content = (title: string, subtitle?: string): Slide => {
    const s = pptx.addSlide({ masterName: "CONTENT" });
    s.addText(title, { x: 0.5, y: 0.3, w: W - 1, h: 0.6, fontSize: 26, bold: true, color: BRAND, fontFace: FONT });
    if (subtitle) s.addText(subtitle, { x: 0.5, y: 0.88, w: W - 1, h: 0.35, fontSize: 13, color: MUTED, fontFace: FONT });
    return s;
  };
  const table = (s: Slide, rows: Row[], o: PptxGenJS.TableProps) =>
    s.addTable(rows, { fontFace: FONT, border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", ...o });

  const phases = p.phases.slice(0, PHASE_COUNT);
  const teamOk = p.team.map((m, i) => ({ m, r: c.team[i] })).filter((x) => x.m.role && x.m.grade && x.m.tier);

  // ---------- 1. Title ----------
  {
    const s = pptx.addSlide();
    s.background = { color: BRAND };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.2, w: W, h: 2.3, fill: { color: "173B5C" } });
    s.addText("Project Cost and Labor Estimation", { x: 0.8, y: 1.3, w: 11.5, h: 0.5, fontSize: 18, color: "BFD4EA", fontFace: FONT });
    s.addText(p.name || "Project", { x: 0.8, y: 1.9, w: 11.5, h: 1.6, fontSize: 40, bold: true, color: "FFFFFF", fontFace: FONT, valign: "top" });
    if (p.client) s.addText(p.client, { x: 0.8, y: 3.6, w: 11.5, h: 0.5, fontSize: 20, color: "FFFFFF", fontFace: FONT });
    const kpis = [
      ["Project Cost after Discount", fmt.usd0(c.costAfterDiscount)],
      ["Total effort", `${fmt.num1(c.totalWork)} m/d`],
      ["Duration", `${fmt.num(c.durationMonths)} months`],
      ["Schedule", `${fmtDate(c.projectStart)} – ${fmtDate(c.projectEnd)}`],
    ];
    kpis.forEach(([k, v], i) => {
      const x = 0.8 + i * 3.0;
      s.addText(v, { x, y: 5.55, w: 2.9, h: 0.6, fontSize: 20, bold: true, color: "FFFFFF", fontFace: FONT });
      s.addText(k, { x, y: 6.15, w: 2.9, h: 0.4, fontSize: 11, color: "BFD4EA", fontFace: FONT });
    });
    s.addText(new Date().toLocaleDateString("ru-RU"), { x: W - 3, y: 0.4, w: 2.5, h: 0.3, fontSize: 11, color: "BFD4EA", align: "right", fontFace: FONT });
  }

  // ---------- 2. Total for the project ----------
  {
    const s = content("Total for the project", "Estimation and finance summary");
    const cards: [string, string, string?][] = [
      ["Project Cost after Discount", fmt.usd(c.costAfterDiscount), "FFF7E0"],
      ["Project Cost", fmt.usd(c.projectCost)],
      ["Total effort", `${fmt.num1(c.totalWork)} m/d`],
      ["Project Duration", `${fmt.num(c.durationMonths)} months`],
    ];
    cards.forEach(([k, v, fill], i) => {
      const x = 0.5 + i * 3.1;
      s.addShape(pptx.ShapeType.roundRect, { x, y: 1.45, w: 2.95, h: 1.25, fill: { color: fill ?? "F3F7FB" }, line: { color: LINE }, rectRadius: 0.1 });
      s.addText(v, { x: x + 0.15, y: 1.55, w: 2.7, h: 0.65, fontSize: 22, bold: true, color: BRAND, fontFace: FONT });
      s.addText(k, { x: x + 0.15, y: 2.18, w: 2.7, h: 0.4, fontSize: 11, color: MUTED, fontFace: FONT });
    });
    const est: Row[] = [
      [hdr("Estimation"), hdr("Value", "right"), hdr("Unit")],
      [cellT("Work (roles)"), num(fmt.num(c.workRoles), { color: checkColor(c.checkTotal), bold: true }), cellT("m/d")],
      [cellT("Contingency Reserve"), num(fmt.num(c.contingency)), cellT("m/d")],
      [cellT("Contingency Reserve %"), num(fmt.pct(c.contingencyPct)), cellT("%")],
      [cellT("Total", totalRow()), num(fmt.num(c.totalWork), totalRow()), cellT("m/d", totalRow())],
    ];
    const fin: Row[] = [
      [hdr("Finance"), hdr("Value", "right"), hdr("Unit")],
      [cellT("Project Cost"), num(fmt.usd(c.projectCost)), cellT("USD")],
      [cellT("Discount %"), num(fmt.pct(c.discountPct)), cellT("%")],
      [cellT("Discount"), num(fmt.usd(c.discountAmount)), cellT("USD")],
      [cellT("Blended Rate"), num(fmt.usd(c.blendedRate)), cellT("USD")],
      [cellT("Project Duration"), num(fmt.num(c.durationMonths)), cellT("Months")],
      [cellT("Project Cost after Discount", totalRow()), num(fmt.usd(c.costAfterDiscount), totalRow()), cellT("USD", totalRow())],
    ];
    table(s, est, { x: 0.5, y: 3.05, w: 5.9, colW: [3.3, 1.7, 0.9], rowH: 0.36 });
    table(s, fin, { x: 6.9, y: 3.05, w: 5.9, colW: [3.3, 1.7, 0.9], rowH: 0.36 });
  }

  // ---------- 3. Team ----------
  {
    const s = content("Team", "Roles, grades, tiers and rates (Fix-Price / A-la-Carte)");
    const rows: Row[] = [[hdr("Role"), hdr("Grade"), hdr("Tier"), hdr("Rate (USD)", "right"), hdr("%", "right"), hdr("Work, m/d", "right")]];
    teamOk.forEach(({ m, r }) => rows.push([cellT(m.role), cellT(m.grade), cellT(m.tier), num(fmt.num(m.rate)), num(fmt.pct(r.share)), num(fmt.num(r.work))]));
    rows.push([cellT("Total", totalRow({ colspan: 4 }) as PptxGenJS.TableCellProps), num(fmt.pct(c.teamShareTotal), totalRow()), num(fmt.num(c.teamWorkTotal), totalRow({ color: checkColor(c.checkTotal) }))]);
    table(s, rows, { x: 0.5, y: 1.4, w: W - 1, colW: [4.8, 1.6, 1.3, 1.6, 1.3, 1.73], rowH: Math.min(0.42, 5.3 / rows.length), fontSize: rows.length > 14 ? 9 : 11 });
  }

  // ---------- 4. Workload of specialists ----------
  {
    const s = content("Workload of specialists", "Effort per role (m/d) and share of the total");
    const labels = teamOk.map(({ m }) => `${m.role} (${m.grade})`);
    s.addChart(pptx.ChartType.bar, [{ name: "Work, m/d", labels, values: teamOk.map(({ r }) => +r.work.toFixed(2)) }], {
      x: 0.4, y: 1.35, w: 7.6, h: 5.5, barDir: "bar", chartColors: [ACCENT],
      catAxisLabelFontSize: 10, valAxisLabelFontSize: 9, catAxisOrientation: "maxMin" as unknown as "minMax",
      showValue: true, dataLabelFontSize: 9, dataLabelFormatCode: "0.0", valGridLine: { color: "EEEEEE", size: 0.5 },
      catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
    });
    const shares = teamOk.filter(({ r }) => r.work > 0);
    if (shares.length) {
      s.addChart(pptx.ChartType.doughnut, [{ name: "Share", labels: shares.map(({ m }) => m.role), values: shares.map(({ r }) => +r.share.toFixed(4)) }], {
        x: 8.2, y: 1.35, w: 4.7, h: 5.5, holeSize: 55, chartColors: SERIES, showPercent: true, showValue: false,
        dataLabelColor: "FFFFFF", dataLabelFontSize: 9, showLegend: true, legendPos: "b", legendFontSize: 9, legendFontFace: FONT,
      });
    }
  }

  // ---------- 5. Project schedule (Gantt) ----------
  {
    const s = content("Project schedule", `${fmtDate(c.projectStart)} – ${fmtDate(c.projectEnd)} · ${c.totalWorkingDays} working days`);
    const d0 = c.projectStart ? isoToDay(c.projectStart)! : null;
    const d1 = c.projectEnd ? isoToDay(c.projectEnd)! : null;
    if (d0 === null || d1 === null || d1 < d0) {
      s.addText("Set the start date and the working days of the phases to build the schedule.", { x: 0.5, y: 3, w: W - 1, h: 0.6, fontSize: 16, color: MUTED, align: "center" });
    } else {
      const labelW = 3.2;
      const gx = 0.5 + labelW;
      const gw = W - 0.5 - gx;
      const gy = 1.95;
      const rowH = 0.78;
      const span = d1 - d0 + 1;
      const xOf = (day: number) => gx + ((day - d0) / span) * gw;
      // month grid
      const first = new Date(dayToIso(d0) + "T00:00:00Z");
      let m = new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth(), 1));
      const gh = rowH * PHASE_COUNT;
      while (m.getTime() / 86_400_000 <= d1) {
        const md = Math.round(m.getTime() / 86_400_000);
        const next = new Date(Date.UTC(m.getUTCFullYear(), m.getUTCMonth() + 1, 1));
        const nd = Math.round(next.getTime() / 86_400_000);
        const a = Math.max(md, d0);
        const b = Math.min(nd, d1 + 1);
        const x = xOf(a);
        const w = xOf(b) - x;
        s.addShape(pptx.ShapeType.rect, { x, y: gy - 0.45, w, h: 0.4, fill: { color: "EEF3F9" }, line: { color: "FFFFFF", width: 1 } });
        if (w > 0.45) {
          s.addText(m.toLocaleDateString("en-US", { month: "short", year: "2-digit", timeZone: "UTC" }), {
            x, y: gy - 0.45, w, h: 0.4, fontSize: 10, color: BRAND, align: "center", bold: true, fontFace: FONT,
          });
        }
        s.addShape(pptx.ShapeType.line, { x, y: gy, w: 0, h: gh, line: { color: LINE, width: 0.75, dashType: "dash" } });
        m = next;
      }
      for (let i = 0; i < PHASE_COUNT; i++) {
        const y = gy + i * rowH;
        if (i % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: W - 1, h: rowH, fill: { color: "F8FAFC" }, line: { color: "F8FAFC" } });
        s.addText([
          { text: phases[i].name, options: { bold: true, color: INK, fontSize: 12, breakLine: true } },
          { text: `${phases[i].workingDays} working days · ${fmt.num(c.effortByRoles[i])} m/d`, options: { color: MUTED, fontSize: 9 } },
        ], { x: 0.55, y, w: labelW - 0.1, h: rowH, valign: "middle", fontFace: FONT });
        const a = c.phaseStart[i] ? isoToDay(c.phaseStart[i]!) : null;
        const b = c.phaseEnd[i] ? isoToDay(c.phaseEnd[i]!) : null;
        if (a !== null && b !== null) {
          const x = xOf(a);
          const w = Math.max(0.08, xOf(b + 1) - x);
          s.addShape(pptx.ShapeType.roundRect, { x, y: y + 0.17, w, h: rowH - 0.34, fill: { color: PHASE_COLORS[i] }, line: { color: PHASE_COLORS[i] }, rectRadius: 0.06 });
          const label = `${fmtDate(c.phaseStart[i])} – ${fmtDate(c.phaseEnd[i])}`;
          if (w > 2.0) s.addText(label, { x, y: y + 0.17, w, h: rowH - 0.34, fontSize: 10, color: "FFFFFF", bold: true, align: "center", fontFace: FONT });
          else {
            const outsideX = x + w + 0.05;
            const fitsRight = outsideX + 2.0 <= W - 0.4;
            s.addText(label, { x: fitsRight ? outsideX : x - 2.05, y: y + 0.17, w: 2.0, h: rowH - 0.34, fontSize: 9, color: INK, align: fitsRight ? "left" : "right", fontFace: FONT });
          }
        }
      }
      // go-live milestone
      const ex = xOf(d1 + 1);
      s.addShape(pptx.ShapeType.diamond, { x: ex - 0.13, y: gy + gh + 0.08, w: 0.26, h: 0.26, fill: { color: BAD }, line: { color: BAD } });
      s.addText(`Project end ${fmtDate(c.projectEnd)}`, { x: ex - 3.2, y: gy + gh + 0.05, w: 3.0, h: 0.32, fontSize: 10, color: BAD, align: "right", bold: true, fontFace: FONT });
    }
  }

  // ---------- 6. Phases ----------
  {
    const s = content("Phases", "Dates, duration, FTE and efforts per phase");
    const rows: Row[] = [[hdr("Phase"), hdr("Start"), hdr("End"), hdr("Working days", "right"), hdr("FTE, total", "right"), hdr("Efforts by Roles, m/d", "right"), hdr("Efforts by Block, m/d", "right")]];
    phases.forEach((ph, i) =>
      rows.push([
        cellT(ph.name, { bold: true, color: PHASE_COLORS[i] }), cellT(fmtDate(c.phaseStart[i])), cellT(fmtDate(c.phaseEnd[i])),
        num(String(ph.workingDays)), num(fmt.num(c.fteByPhase[i])),
        num(fmt.num(c.effortByRoles[i]), { color: checkColor(c.checkPhases[i]), bold: true }), num(fmt.num(c.effortByBlock[i])),
      ]),
    );
    rows.push([
      cellT("TOTAL", totalRow()), cellT(fmtDate(c.projectStart), totalRow()), cellT(fmtDate(c.projectEnd), totalRow()),
      num(String(c.totalWorkingDays), totalRow()), num(fmt.num(c.fteTotal), totalRow()),
      num(fmt.num(c.effortByRolesTotal), totalRow({ color: checkColor(c.checkPhasesTotal) })), num(fmt.num(c.effortByBlockTotal), totalRow()),
    ]);
    table(s, rows, { x: 0.5, y: 1.4, w: W - 1, colW: [2.73, 1.3, 1.3, 1.4, 1.3, 2.15, 2.15], rowH: 0.5 });
    s.addText("Green = efforts by Roles match efforts by Block, red = they differ.", { x: 0.5, y: 5.4, w: W - 1, h: 0.3, fontSize: 10, color: MUTED, italic: true });
  }

  // ---------- 7. Roles engagement ----------
  {
    const allocs = p.allocations.map((a, i) => ({ a, avg: c.allocationAvgFte[i] })).filter(({ a }) => a.role);
    const chunks = chunk(allocs, 13);
    chunks.forEach((part, pi) => {
      const s = content(`Roles engagement${chunks.length > 1 ? ` (${pi + 1}/${chunks.length})` : ""}`, "Engagement per phase, 100% = 1 FTE");
      const rows: Row[] = [[hdr("Block"), hdr("Role · Grade · Tier"), ...phases.map((ph) => hdr(ph.name, "center")), hdr("Avg FTE", "right")]];
      part.forEach(({ a, avg }) =>
        rows.push([
          cellT(a.block, { fontSize: 10 }), cellT(`${a.role} · ${a.grade} · ${a.tier}`, { fontSize: 10 }),
          ...a.fte.slice(0, PHASE_COUNT).map((v) => cellT(v ? `${Math.round(v * 100)}%` : "", { align: "center", fontSize: 10, fill: { color: heat(v) }, color: v >= 1 ? "FFFFFF" : INK })),
          num(fmt.num(avg), { fontSize: 10 }),
        ]),
      );
      if (pi === chunks.length - 1) {
        rows.push([
          cellT("FTE, total", totalRow({ colspan: 2 }) as PptxGenJS.TableCellProps),
          ...c.fteByPhase.map((v) => cellT(fmt.num(v), totalRow({ align: "center" }))),
          num(fmt.num(c.fteTotal), totalRow()),
        ]);
      }
      table(s, rows, { x: 0.5, y: 1.4, w: W - 1, colW: [1.6, 3.93, 1.2, 1.2, 1.2, 1.2, 1.2, 0.8], rowH: 0.36, fontSize: 10 });
    });
  }

  // ---------- 8. Block Estimation ----------
  {
    const s = content("Block Estimation", c.upliftTotal !== 1 ? `Block "All" spread proportionally (uplift ×${fmt.num(c.upliftTotal)})` : "Effort per block, m/d");
    const be = c.blockEstimation;
    if (be.some((b) => b.value > 0)) {
      s.addChart(pptx.ChartType.bar, [{ name: "m/d", labels: be.map((b) => b.block), values: be.map((b) => +b.value.toFixed(2)) }], {
        x: 0.4, y: 1.4, w: 7.4, h: 5.4, barDir: "col", chartColors: SERIES.slice(0, Math.max(1, be.length)), showValue: true,
        dataLabelFormatCode: "0.0", dataLabelFontSize: 10, catAxisLabelFontSize: 10, valGridLine: { color: "EEEEEE", size: 0.5 },
      });
    }
    const rows: Row[] = [[hdr("Block"), hdr("m/d", "right"), hdr("Share", "right")]];
    be.forEach((b) => rows.push([cellT(b.block), num(fmt.num(b.value)), num(fmt.pct(c.blockEstimationTotal ? b.value / c.blockEstimationTotal : 0))]));
    rows.push([cellT("Total", totalRow()), num(fmt.num(c.blockEstimationTotal), totalRow()), num("100%", totalRow())]);
    table(s, rows, { x: 8.1, y: 1.5, w: 4.73, colW: [2.33, 1.3, 1.1], rowH: 0.4 });
  }

  // ---------- 9. Phase Estimation ----------
  {
    const s = content("Phase Estimation", "Effort per phase from the service packages, m/d");
    const pe = c.phaseEstimation;
    s.addChart(pptx.ChartType.bar, [{ name: "m/d", labels: pe.map((x) => x.phase), values: pe.map((x) => +x.value.toFixed(2)) }], {
      x: 0.4, y: 1.4, w: 7.4, h: 5.4, barDir: "col", chartColors: PHASE_COLORS, showValue: true,
      dataLabelFormatCode: "0.0", dataLabelFontSize: 10, catAxisLabelFontSize: 10, valGridLine: { color: "EEEEEE", size: 0.5 },
    });
    const rows: Row[] = [[hdr("Phase"), hdr("m/d", "right"), hdr("Share", "right")]];
    pe.forEach((x, i) => rows.push([cellT(x.phase, { color: PHASE_COLORS[i], bold: true }), num(fmt.num(x.value)), num(fmt.pct(c.phaseEstimationTotal ? x.value / c.phaseEstimationTotal : 0))]));
    rows.push([cellT("Total", totalRow()), num(fmt.num(c.phaseEstimationTotal), totalRow()), num("100%", totalRow())]);
    table(s, rows, { x: 8.1, y: 1.5, w: 4.73, colW: [2.53, 1.1, 1.1], rowH: 0.42 });
  }

  // ---------- 10. Service packages ----------
  {
    const pk = p.packages.map((x, i) => ({ x, t: c.packageTotals[i] })).filter(({ x, t }) => x.packageId || x.name || t);
    const chunks = chunk(pk, 12);
    if (!chunks.length) chunks.push([]);
    chunks.forEach((part, pi) => {
      const s = content(`Service packages${chunks.length > 1 ? ` (${pi + 1}/${chunks.length})` : ""}`, "Efforts by block and phase, m/d");
      const rows: Row[] = [[hdr("Block"), hdr("ID"), hdr("Service package"), ...phases.map((ph) => hdr(ph.name, "right")), hdr("Total", "right")]];
      part.forEach(({ x, t }) =>
        rows.push([cellT(x.block, { fontSize: 9 }), cellT(x.packageId, { fontSize: 9 }), cellT(x.name, { fontSize: 9 }), ...x.days.slice(0, PHASE_COUNT).map((v) => num(v ? fmt.num(v) : "", { fontSize: 9 })), num(fmt.num(t), { fontSize: 9, bold: true })]),
      );
      if (pi === chunks.length - 1) {
        rows.push([
          cellT("Total efforts amount by Block", totalRow({ colspan: 3, fontSize: 9 }) as PptxGenJS.TableCellProps),
          ...c.effortByBlock.map((v, i) => num(fmt.num(v), totalRow({ fontSize: 9, color: checkColor(c.checkPhases[i]) }))),
          num(fmt.num(c.effortByBlockTotal), totalRow({ fontSize: 9 })),
        ]);
      }
      table(s, rows, { x: 0.5, y: 1.4, w: W - 1, colW: [1.3, 0.9, 3.73, 1.1, 1.1, 1.1, 1.1, 1.1, 0.9], rowH: 0.36, fontSize: 9 });
    });
  }

  // ---------- 11. Cost structure ----------
  {
    const s = content("Cost structure", "Project cost by role: rate × work, incl. contingency reserve, before discount");
    const items = teamOk.map(({ m, r }) => ({ label: `${m.role} (${m.grade}, ${m.tier})`, cost: m.rate * r.work * (1 + c.contingencyPct) })).filter((x) => x.cost > 0);
    if (items.length) {
      s.addChart(pptx.ChartType.bar, [{ name: "USD", labels: items.map((x) => x.label), values: items.map((x) => +x.cost.toFixed(2)) }], {
        x: 0.4, y: 1.35, w: 8.0, h: 5.5, barDir: "bar", chartColors: [BRAND], catAxisOrientation: "maxMin" as unknown as "minMax", showValue: true,
        dataLabelFormatCode: "$#,##0", dataLabelFontSize: 9, catAxisLabelFontSize: 9, valAxisLabelFormatCode: "$#,##0", valGridLine: { color: "EEEEEE", size: 0.5 },
      });
    }
    const rows: Row[] = [
      [hdr("Item"), hdr("USD", "right")],
      [cellT("Project Cost"), num(fmt.usd(c.projectCost))],
      [cellT(`Discount (${fmt.pct(c.discountPct)})`), num(`−${fmt.usd(c.discountAmount)}`, { color: BAD })],
      [cellT("Project Cost after Discount", totalRow()), num(fmt.usd(c.costAfterDiscount), totalRow())],
      [cellT("Blended Rate"), num(fmt.usd(c.blendedRate))],
    ];
    table(s, rows, { x: 8.7, y: 1.5, w: 4.13, colW: [2.53, 1.6], rowH: 0.45 });
  }

  // ---------- 12. Checks ----------
  {
    const s = content("Consistency check", "Total efforts amount by Roles vs by Block");
    const st = (x: CheckState) => (x === "ok" ? "✔ match" : x === "bad" ? "✖ differ" : "—");
    const rows: Row[] = [[hdr("Phase"), hdr("By Roles, m/d", "right"), hdr("By Block, m/d", "right"), hdr("Difference", "right"), hdr("Status", "center")]];
    phases.forEach((ph, i) =>
      rows.push([cellT(ph.name), num(fmt.num(c.effortByRoles[i])), num(fmt.num(c.effortByBlock[i])), num(fmt.num(c.effortByRoles[i] - c.effortByBlock[i])), cellT(st(c.checkPhases[i]), { align: "center", bold: true, color: checkColor(c.checkPhases[i]) })]),
    );
    rows.push([cellT("TOTAL", totalRow()), num(fmt.num(c.effortByRolesTotal), totalRow()), num(fmt.num(c.effortByBlockTotal), totalRow()), num(fmt.num(c.effortByRolesTotal - c.effortByBlockTotal), totalRow()), cellT(st(c.checkPhasesTotal), totalRow({ align: "center", color: checkColor(c.checkPhasesTotal) }))]);
    rows.push([cellT("Work (Team) vs Block", { italic: true }), num(fmt.num(c.teamWorkTotal)), num(fmt.num(c.effortByBlockTotal)), num(fmt.num(c.teamWorkTotal - c.effortByBlockTotal)), cellT(st(c.checkTotal), { align: "center", bold: true, color: checkColor(c.checkTotal) })]);
    table(s, rows, { x: 0.5, y: 1.4, w: W - 1, colW: [4.33, 2, 2, 2, 2], rowH: 0.45 });
    const orphan = p.allocations.filter((a, i) => a.role && !c.allocationInTeam[i]);
    if (orphan.length) {
      s.addText(
        `Engagement rows not matching any Team member (not counted in Work): ${orphan.map((a) => `${a.role} · ${a.grade} · ${a.tier}`).join("; ")}`,
        { x: 0.5, y: 5.4, w: W - 1, h: 0.8, fontSize: 11, color: BAD, fontFace: FONT },
      );
    }
  }

  // ---------- 13. Assumptions ----------
  const assumptions = p.assumptions.filter((a) => a.text.trim());
  if (assumptions.length) {
    chunk(assumptions, 10).forEach((part, pi, all) => {
      const s = content(`Assumptions${all.length > 1 ? ` (${pi + 1}/${all.length})` : ""}`);
      s.addText(
        part.map((a) => ({ text: a.text, options: { bullet: { type: "number" as const }, breakLine: true } })),
        { x: 0.6, y: 1.4, w: W - 1.2, h: 5.4, fontSize: 15, color: INK, valign: "top", paraSpaceAfter: 8, fontFace: FONT },
      );
    });
  }

  return pptx;
}

function chunk<T>(xs: T[], n: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < xs.length; i += n) out.push(xs.slice(i, i + n));
  return out;
}

/** Engagement heat colour: 0 → white, 1 FTE → brand blue. */
function heat(v: number): string {
  if (!v) return "FFFFFF";
  const t = Math.min(1, v / 1.5);
  const from = [0xea, 0xf1, 0xfb];
  const to = [0x1f, 0x4e, 0x79];
  return from.map((f, i) => Math.round(f + (to[i] - f) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
}

export async function exportToPptx(p: ProjectData, c: CalcResult): Promise<void> {
  const deck = buildDeck(p, c);
  await deck.writeFile({ fileName: `${safeFileName(p.name)} - Estimation.pptx` });
}

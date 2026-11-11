import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { api } from "./api";
import { NumInput } from "./components/NumInput";
import { OpenDialog } from "./components/OpenDialog";
import { calculate, type CheckState } from "./shared/calc";
import { fmtDate } from "./shared/dates";
import { blankProject, newAllocation, newPackage, newTeamMember, sampleProject, uid } from "./shared/factory";
import { fmt } from "./format";
import { holidayName } from "./shared/holidays";
import { ALL_BLOCK, PHASE_COUNT, type Catalog, type ProjectData, type TeamMember } from "./shared/types";

type Toast = { kind: "ok" | "bad" | "info"; text: string } | null;
const PH = Array.from({ length: PHASE_COUNT }, (_, i) => i);
const memberLabel = (m: { role: string; grade: string; tier: string }) => `${m.role} · ${m.grade} · ${m.tier}`;
const sameCombo = (a: { role: string; grade: string; tier: string }, b: { role: string; grade: string; tier: string }) =>
  a.role.trim().toLowerCase() === b.role.trim().toLowerCase() &&
  a.grade.trim().toLowerCase() === b.grade.trim().toLowerCase() &&
  a.tier.trim().toLowerCase() === b.tier.trim().toLowerCase();
const chk = (s: CheckState) => (s === "ok" ? "chk-ok" : s === "bad" ? "chk-bad" : "");

export default function App() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [data, setData] = useState<ProjectData | null>(null);
  const [projectId, setProjectId] = useState<string | null>(null);
  const [dirty, setDirty] = useState(false);
  const [busy, setBusy] = useState<string | null>(null);
  const [toast, setToast] = useState<Toast>(null);
  const [showOpen, setShowOpen] = useState(false);
  const [showCreate, setShowCreate] = useState(false);
  const toastTimer = useRef<number | undefined>(undefined);

  const notify = useCallback((t: Toast) => {
    setToast(t);
    window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast(null), 4000);
  }, []);

  useEffect(() => {
    api
      .catalog()
      .then((c) => {
        setCatalog(c);
        const last = localStorage.getItem("pe:lastProject");
        if (last) {
          api.get(last).then((p) => {
            setData(p.data);
            setProjectId(p.id);
          }).catch(() => setData(blankProject(c.defaultHolidays)));
        } else setData(blankProject(c.defaultHolidays));
      })
      .catch((e: Error) => notify({ kind: "bad", text: `Не удалось загрузить справочник ролей: ${e.message}` }));
  }, [notify]);

  useEffect(() => {
    const h = (e: BeforeUnloadEvent) => {
      if (dirty) e.preventDefault();
    };
    window.addEventListener("beforeunload", h);
    return () => window.removeEventListener("beforeunload", h);
  }, [dirty]);

  const calc = useMemo(() => (data ? calculate(data) : null), [data]);

  const update = useCallback((fn: (d: ProjectData) => ProjectData) => {
    setData((d) => (d ? fn(d) : d));
    setDirty(true);
  }, []);

  /** Ask the Worker for the rates of the given team members only. */
  const resolveRates = useCallback(async (members: TeamMember[]) => {
    const need = members.filter((m) => m.role && m.grade && m.tier);
    if (!need.length) return;
    try {
      const rates = await api.rates(need.map(({ role, grade, tier }) => ({ role, grade, tier })));
      const byId = new Map(need.map((m, i) => [m.id, rates[i]]));
      setData((d) =>
        d
          ? {
              ...d,
              team: d.team.map((m) => {
                const r = byId.get(m.id);
                // ignore stale answers if the member changed meanwhile
                const src = need.find((x) => x.id === m.id);
                return r !== undefined && src && sameCombo(src, m) ? { ...m, rate: r } : m;
              }),
            }
          : d,
      );
    } catch (e) {
      notify({ kind: "bad", text: `Не удалось получить рейты: ${(e as Error).message}` });
    }
  }, [notify]);

  function confirmDiscard() {
    return !dirty || confirm("Есть несохранённые изменения. Продолжить без сохранения?");
  }

  // ---------- toolbar actions ----------
  function createProject(kind: "blank" | "sample") {
    if (!catalog) return;
    const d = kind === "blank" ? blankProject(catalog.defaultHolidays) : sampleProject(catalog.defaultHolidays);
    setData(d);
    setProjectId(null);
    setDirty(true);
    setShowCreate(false);
    localStorage.removeItem("pe:lastProject");
    if (kind === "sample") resolveRates(d.team);
    notify({ kind: "info", text: "Новый проект создан. Не забудьте сохранить." });
  }

  async function save() {
    if (!data) return;
    if (!data.name.trim()) return notify({ kind: "bad", text: "Укажите название проекта" });
    setBusy("save");
    try {
      const saved = projectId ? await api.update(projectId, data) : await api.create(data);
      setProjectId(saved.id);
      setData(saved.data);
      setDirty(false);
      localStorage.setItem("pe:lastProject", saved.id);
      notify({ kind: "ok", text: "Проект сохранён" });
    } catch (e) {
      notify({ kind: "bad", text: `Ошибка сохранения: ${(e as Error).message}` });
    } finally {
      setBusy(null);
    }
  }

  async function openProject(id: string) {
    if (!confirmDiscard()) return;
    setBusy("open");
    try {
      const p = await api.get(id);
      setData(p.data);
      setProjectId(p.id);
      setDirty(false);
      setShowOpen(false);
      localStorage.setItem("pe:lastProject", p.id);
    } catch (e) {
      notify({ kind: "bad", text: `Не удалось открыть: ${(e as Error).message}` });
    } finally {
      setBusy(null);
    }
  }

  async function exportExcel() {
    if (!data || !calc) return;
    setBusy("xlsx");
    try {
      const { exportToExcel } = await import("./export/excel");
      await exportToExcel(data, calc);
    } catch (e) {
      console.error(e);
      notify({ kind: "bad", text: `Ошибка выгрузки Excel: ${(e as Error).message}` });
    } finally {
      setBusy(null);
    }
  }

  async function exportPptx() {
    if (!data || !calc) return;
    setBusy("pptx");
    try {
      const { exportToPptx } = await import("./export/pptx");
      await exportToPptx(data, calc);
    } catch (e) {
      console.error(e);
      notify({ kind: "bad", text: `Ошибка формирования презентации: ${(e as Error).message}` });
    } finally {
      setBusy(null);
    }
  }

  if (!data || !calc || !catalog) {
    return (
      <div className="loading">
        {toast ? <div className="alert bad">{toast.text}</div> : "Загрузка…"}
      </div>
    );
  }

  // ---------- editors ----------
  const blocks = data.blocks.filter((b) => b.trim());
  const blockOptions = [...blocks, ALL_BLOCK];
  const gradesFor = (role: string) => catalog.roles.find((r) => r.role === role)?.grades ?? [];
  const completeMembers = data.team.filter((m) => m.role && m.grade && m.tier);

  function setTeam(i: number, patch: Partial<TeamMember>) {
    const cur = data!.team[i];
    const next: TeamMember = { ...cur, ...patch };
    if (patch.role !== undefined && patch.role !== cur.role) {
      const gs = gradesFor(patch.role);
      next.grade = gs.length === 1 ? gs[0] : gs.includes(cur.grade) ? cur.grade : "";
    }
    const comboChanged = !sameCombo(cur, next);
    if (comboChanged) next.rate = 0;
    update((d) => ({
      ...d,
      team: d.team.map((m, k) => (k === i ? next : m)),
      // keep linked engagement rows in sync with the team member
      allocations: comboChanged
        ? d.allocations.map((a) =>
            a.memberId === cur.id ? { ...a, role: next.role, grade: next.grade, tier: next.tier } : a,
          )
        : d.allocations,
    }));
    if (comboChanged) resolveRates([next]);
  }

  function renameBlock(i: number, name: string) {
    const old = data!.blocks[i];
    update((d) => ({
      ...d,
      blocks: d.blocks.map((b, k) => (k === i ? name : b)),
      allocations: old ? d.allocations.map((a) => (a.block === old ? { ...a, block: name } : a)) : d.allocations,
      packages: old ? d.packages.map((p) => (p.block === old ? { ...p, block: name } : p)) : d.packages,
    }));
  }

  const busyLabel = (k: string, label: string) => (busy === k ? "…" : label);

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="logo">▦</span>
          <div>
            <div className="brand-title">Project Cost and Labor Estimation</div>
            <div className="brand-sub">
              {projectId ? "Сохранённый проект" : "Новый проект (не сохранён)"}
              {dirty && <span className="dirty"> · есть изменения</span>}
            </div>
          </div>
        </div>
        <nav className="actions">
          <button onClick={() => confirmDiscard() && setShowCreate(true)}>＋ Создать проект</button>
          <button className="primary" onClick={save} disabled={!!busy}>{busyLabel("save", "💾 Сохранить проект")}</button>
          <button onClick={() => setShowOpen(true)} disabled={!!busy}>📂 Выбрать на редактирование</button>
          <button onClick={exportExcel} disabled={!!busy}>{busyLabel("xlsx", "⬇ Выгрузить в Excel")}</button>
          <button onClick={exportPptx} disabled={!!busy}>{busyLabel("pptx", "📊 Сформировать презентацию")}</button>
        </nav>
      </header>

      {toast && <div className={`toast ${toast.kind}`} onClick={() => setToast(null)}>{toast.text}</div>}

      <main>
        <section className="card project-head">
          <label>
            <span>Project Name</span>
            <input value={data.name} onChange={(e) => update((d) => ({ ...d, name: e.target.value }))} />
          </label>
          <label>
            <span>Client</span>
            <input value={data.client} onChange={(e) => update((d) => ({ ...d, client: e.target.value }))} />
          </label>
        </section>

        <div className="grid-top">
          {/* ---------- Total for the project ---------- */}
          <section className="card">
            <h2>Total for the project</h2>
            <table className="kv">
              <tbody>
                <tr className="kv-group"><td colSpan={3}>Estimation</td></tr>
                <tr>
                  <td>Work (roles)</td>
                  <td className={`v ${chk(calc.checkTotal)}`} title="Зелёный — совпадает с Total efforts amount by Block, красный — расходится">
                    {fmt.num(calc.workRoles)}
                  </td>
                  <td className="u">m/d</td>
                </tr>
                <tr>
                  <td>Contingency Reserve m/d</td>
                  <td className="v"><NumInput value={data.contingency} onChange={(v) => update((d) => ({ ...d, contingency: v }))} /></td>
                  <td className="u">m/d</td>
                </tr>
                <tr><td>Contingency Reserve %</td><td className="v">{fmt.pct(calc.contingencyPct)}</td><td className="u">%</td></tr>
                <tr className="strong"><td>Total</td><td className="v">{fmt.num(calc.totalWork)}</td><td className="u">m/d</td></tr>
                <tr className="kv-group"><td colSpan={3}>Finance</td></tr>
                <tr><td>Project Cost</td><td className="v">{fmt.usd(calc.projectCost)}</td><td className="u">USD</td></tr>
                <tr>
                  <td>Discount %</td>
                  <td className="v"><NumInput value={data.discount} scale={100} onChange={(v) => update((d) => ({ ...d, discount: v }))} /></td>
                  <td className="u">%</td>
                </tr>
                <tr><td>Discount</td><td className="v">{fmt.usd(calc.discountAmount)}</td><td className="u">USD</td></tr>
                <tr><td>Blended Rate</td><td className="v">{fmt.usd(calc.blendedRate)}</td><td className="u">USD</td></tr>
                <tr><td>Project Duration</td><td className="v">{fmt.num(calc.durationMonths)}</td><td className="u">Months</td></tr>
                <tr className="strong total"><td>Project Cost after Discount</td><td className="v">{fmt.usd(calc.costAfterDiscount)}</td><td className="u">USD</td></tr>
              </tbody>
            </table>
          </section>

          {/* ---------- Team ---------- */}
          <section className="card">
            <h2>
              Team <small>Рейт подставляется из Rate Card (Fix-Price / A-la-Carte) только для выбранных ролей</small>
            </h2>
            <div className="table-wrap">
              <table className="grid">
                <thead>
                  <tr><th>Role</th><th>Grade</th><th>Tier</th><th className="r">Rate (USD)</th><th className="r">%</th><th className="r">Work</th><th></th></tr>
                </thead>
                <tbody>
                  {data.team.map((m, i) => (
                    <tr key={m.id}>
                      <td>
                        <select value={m.role} onChange={(e) => setTeam(i, { role: e.target.value })}>
                          <option value="">— роль —</option>
                          {catalog.roles.map((r) => <option key={r.role} value={r.role}>{r.role}</option>)}
                          {m.role && !catalog.roles.some((r) => r.role === m.role) && <option value={m.role}>{m.role}</option>}
                        </select>
                      </td>
                      <td>
                        <select value={m.grade} onChange={(e) => setTeam(i, { grade: e.target.value })} disabled={!m.role}>
                          <option value="">—</option>
                          {gradesFor(m.role).map((g) => <option key={g} value={g}>{g}</option>)}
                        </select>
                      </td>
                      <td>
                        <select value={m.tier} onChange={(e) => setTeam(i, { tier: e.target.value })}>
                          <option value="">—</option>
                          {catalog.tiers.map((t) => <option key={t} value={t}>{t}</option>)}
                        </select>
                      </td>
                      <td className={`r ${m.role && m.grade && m.tier && !m.rate ? "chk-bad" : ""}`} title={m.role && m.grade && m.tier && !m.rate ? "Для этой комбинации нет рейта в Rate Card" : undefined}>
                        {m.role && m.grade && m.tier ? fmt.num(m.rate) : ""}
                      </td>
                      <td className="r">{fmt.pct(calc.team[i]?.share ?? 0)}</td>
                      <td className="r">{fmt.num(calc.team[i]?.work ?? 0)}</td>
                      <td><button className="icon-btn danger" title="Удалить" onClick={() => update((d) => ({ ...d, team: d.team.filter((_, k) => k !== i) }))}>✕</button></td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr>
                    <td colSpan={4}>Total</td>
                    <td className="r">{fmt.pct(calc.teamShareTotal)}</td>
                    <td className={`r ${chk(calc.checkTotal)}`}>{fmt.num(calc.teamWorkTotal)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            </div>
            <button className="add" onClick={() => update((d) => ({ ...d, team: [...d.team, newTeamMember()] }))}>＋ Добавить участника</button>
          </section>
        </div>

        <div className="grid-two">
          {/* ---------- Block Estimation ---------- */}
          <section className="card">
            <h2>Block Estimation <small>Блоки проекта; часы блока «All» распределяются пропорционально</small></h2>
            <table className="grid">
              <thead><tr><th>Block</th><th className="r">Estimation, m/d</th><th></th></tr></thead>
              <tbody>
                {data.blocks.map((b, i) => (
                  <tr key={i}>
                    <td><input value={b} placeholder="Название блока" onChange={(e) => renameBlock(i, e.target.value)} /></td>
                    <td className="r">{b.trim() ? fmt.num(calc.blockEstimation.find((x) => x.block === b)?.value ?? 0) : ""}</td>
                    <td><button className="icon-btn danger" title="Удалить" onClick={() => update((d) => ({ ...d, blocks: d.blocks.filter((_, k) => k !== i) }))}>✕</button></td>
                  </tr>
                ))}
                <tr className="muted"><td>{ALL_BLOCK} <small>(служебный блок)</small></td><td className="r">×{fmt.num(calc.upliftTotal)}</td><td></td></tr>
              </tbody>
              <tfoot><tr><td>Total</td><td className="r">{fmt.num(calc.blockEstimationTotal)}</td><td></td></tr></tfoot>
            </table>
            <button className="add" onClick={() => update((d) => ({ ...d, blocks: [...d.blocks, ""] }))}>＋ Добавить блок</button>
          </section>

          {/* ---------- Phase Estimation + mini Gantt ---------- */}
          <section className="card">
            <h2>Phase Estimation</h2>
            <table className="grid">
              <thead><tr><th>Phase</th><th>Start</th><th>End</th><th className="r">m/d</th></tr></thead>
              <tbody>
                {calc.phaseEstimation.map((p, i) => (
                  <tr key={i}>
                    <td>{p.phase}</td>
                    <td>{fmtDate(calc.phaseStart[i])}</td>
                    <td>{fmtDate(calc.phaseEnd[i])}</td>
                    <td className="r">{fmt.num(p.value)}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot><tr><td colSpan={3}>Total</td><td className="r">{fmt.num(calc.phaseEstimationTotal)}</td></tr></tfoot>
            </table>
            <MiniGantt data={data} calc={calc} />
          </section>
        </div>

        {/* ---------- Phases and roles engagement ---------- */}
        <section className="card">
          <h2>Phases &amp; roles engagement <small>Engagement в %: 100% = 1 FTE, 200% = 2 FTE</small></h2>
          <div className="table-wrap">
            <table className="grid phases aligned">
              <colgroup>
                <col className="block-col" />
                <col />
                {PH.map((i) => <col key={i} className="ph-col" />)}
                <col className="tot-col" />
                <col className="del-col" />
              </colgroup>
              <thead>
                <tr>
                  <th>Block</th>
                  <th>Role · Grade · Tier</th>
                  {PH.map((i) => (
                    <th key={i} className="phase-h">
                      <textarea
                        rows={2}
                        title={data.phases[i].name}
                        value={data.phases[i].name}
                        onChange={(e) => update((d) => ({ ...d, phases: d.phases.map((p, k) => (k === i ? { ...p, name: e.target.value.replace(/\n/g, " ") } : p)) }))}
                      />
                    </th>
                  ))}
                  <th className="r">TOTAL</th>
                  <th></th>
                </tr>
              </thead>
              <tbody className="phase-meta">
                <tr>
                  <td colSpan={2}>Phase Start Date</td>
                  {PH.map((i) => (
                    <td key={i}>
                      {i === 0 ? (
                        <input type="date" value={data.startDate} onChange={(e) => update((d) => ({ ...d, startDate: e.target.value }))} />
                      ) : (
                        <span className="calc">{fmtDate(calc.phaseStart[i])}</span>
                      )}
                    </td>
                  ))}
                  <td className="r">{fmtDate(calc.projectStart)}</td><td></td>
                </tr>
                <tr>
                  <td colSpan={2}>Phase End Date</td>
                  {PH.map((i) => <td key={i}><span className="calc">{fmtDate(calc.phaseEnd[i])}</span></td>)}
                  <td className="r">{fmtDate(calc.projectEnd)}</td><td></td>
                </tr>
                <tr>
                  <td colSpan={2}>Amount of working days</td>
                  {PH.map((i) => (
                    <td key={i}>
                      <NumInput value={data.phases[i].workingDays} min={0} onChange={(v) => update((d) => ({ ...d, phases: d.phases.map((p, k) => (k === i ? { ...p, workingDays: Math.round(v) } : p)) }))} />
                    </td>
                  ))}
                  <td className="r">{fmt.int(calc.totalWorkingDays)}</td><td></td>
                </tr>
                <tr>
                  <td colSpan={2}>FTE, total</td>
                  {PH.map((i) => <td key={i} className="r">{fmt.num(calc.fteByPhase[i])}</td>)}
                  <td className="r">{fmt.num(calc.fteTotal)}</td><td></td>
                </tr>
                <tr className="strong">
                  <td colSpan={2}>Total efforts amount by Roles</td>
                  {PH.map((i) => <td key={i} className={`r ${chk(calc.checkPhases[i])}`}>{fmt.num(calc.effortByRoles[i])}</td>)}
                  <td className={`r ${chk(calc.checkPhasesTotal)}`}>{fmt.num(calc.effortByRolesTotal)}</td><td></td>
                </tr>
              </tbody>
              <tbody>
                {data.allocations.map((a, i) => {
                  const inTeam = calc.allocationInTeam[i];
                  const linked = completeMembers.find((m) => m.id === a.memberId && sameCombo(m, a)) ?? completeMembers.find((m) => sameCombo(m, a));
                  return (
                    <tr key={a.id}>
                      <td>
                        <select value={a.block} onChange={(e) => update((d) => ({ ...d, allocations: d.allocations.map((x, k) => (k === i ? { ...x, block: e.target.value } : x)) }))}>
                          <option value="">—</option>
                          {blockOptions.map((b) => <option key={b} value={b}>{b}</option>)}
                          {a.block && !blockOptions.includes(a.block) && <option value={a.block}>{a.block}</option>}
                        </select>
                      </td>
                      <td>
                        <select
                          className={a.role && !inTeam ? "chk-bad" : ""}
                          title={a.role && !inTeam ? "Этой комбинации Role/Grade/Tier нет в Team — трудозатраты не попадут в Work" : undefined}
                          value={linked?.id ?? (a.role ? "__orphan" : "")}
                          onChange={(e) => {
                            const m = completeMembers.find((x) => x.id === e.target.value);
                            update((d) => ({
                              ...d,
                              allocations: d.allocations.map((x, k) =>
                                k === i ? { ...x, memberId: m?.id, role: m?.role ?? "", grade: m?.grade ?? "", tier: m?.tier ?? "" } : x,
                              ),
                            }));
                          }}
                        >
                          <option value="">— участник команды —</option>
                          {completeMembers.map((m) => <option key={m.id} value={m.id}>{memberLabel(m)}</option>)}
                          {a.role && !linked && <option value="__orphan">⚠ {memberLabel(a)} (нет в Team)</option>}
                        </select>
                      </td>
                      {PH.map((p) => (
                        <td key={p}>
                          <NumInput value={a.fte[p]} scale={100} min={0} placeholder="" className="pct" onChange={(v) => update((d) => ({ ...d, allocations: d.allocations.map((x, k) => (k === i ? { ...x, fte: x.fte.map((f, j) => (j === p ? v : f)) } : x)) }))} />
                        </td>
                      ))}
                      <td className="r">{fmt.num(calc.allocationAvgFte[i])}</td>
                      <td><button className="icon-btn danger" title="Удалить" onClick={() => update((d) => ({ ...d, allocations: d.allocations.filter((_, k) => k !== i) }))}>✕</button></td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className="muted">
                  <td colSpan={2}>Uplift factor for block "All"</td>
                  {PH.map((i) => <td key={i} className="r">{fmt.num(calc.uplift[i])}</td>)}
                  <td className="r">{fmt.num(calc.upliftTotal)}</td><td></td>
                </tr>
              </tfoot>
            </table>
          </div>
          <button className="add" onClick={() => update((d) => ({ ...d, allocations: [...d.allocations, newAllocation()] }))}>＋ Добавить строку вовлечённости</button>
        </section>

        {/* ---------- Entry table ---------- */}
        <section className="card">
          <h2>Service packages <small>Трудозатраты по блокам, m/d</small></h2>
          <div className="table-wrap">
            <table className="grid packages aligned">
              <colgroup>
                <col className="block-col" />
                <col className="id-col" />
                <col />
                {PH.map((i) => <col key={i} className="ph-col" />)}
                <col className="tot-col" />
                <col className="del-col" />
              </colgroup>
              <thead>
                <tr>
                  <th>Block</th>
                  <th>Package ID</th>
                  <th>Name of the service package</th>
                  {PH.map((i) => <th key={i} className="c ph-name">{data.phases[i].name}</th>)}
                  <th className="r">Total, m/d</th>
                  <th></th>
                </tr>
                <tr className="strong">
                  <td colSpan={3}>Total efforts amount by Block</td>
                  {PH.map((i) => <td key={i} className={`r ${chk(calc.checkPhases[i])}`}>{fmt.num(calc.effortByBlock[i])}</td>)}
                  <td className={`r ${chk(calc.checkTotal)}`}>{fmt.num(calc.effortByBlockTotal)}</td><td></td>
                </tr>
              </thead>
              <tbody>
                {data.packages.map((p, i) => {
                  const set = (patch: Partial<typeof p>) => update((d) => ({ ...d, packages: d.packages.map((x, k) => (k === i ? { ...x, ...patch } : x)) }));
                  return (
                    <tr key={p.id}>
                      <td>
                        <select value={p.block} onChange={(e) => set({ block: e.target.value })}>
                          <option value="">—</option>
                          {blockOptions.map((b) => <option key={b} value={b}>{b}</option>)}
                          {p.block && !blockOptions.includes(p.block) && <option value={p.block}>{p.block}</option>}
                        </select>
                      </td>
                      <td><input value={p.packageId} onChange={(e) => set({ packageId: e.target.value })} /></td>
                      <td><input value={p.name} onChange={(e) => set({ name: e.target.value })} /></td>
                      {PH.map((j) => (
                        <td key={j}><NumInput value={p.days[j]} min={0} placeholder="" onChange={(v) => set({ days: p.days.map((x, k) => (k === j ? v : x)) })} /></td>
                      ))}
                      <td className="r">{fmt.num(calc.packageTotals[i])}</td>
                      <td><button className="icon-btn danger" title="Удалить" onClick={() => update((d) => ({ ...d, packages: d.packages.filter((_, k) => k !== i) }))}>✕</button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button className="add" onClick={() => update((d) => ({ ...d, packages: [...d.packages, newPackage()] }))}>＋ Добавить пакет</button>
        </section>

        <div className="grid-two">
          <section className="card">
            <h2>Assumptions</h2>
            {data.assumptions.length === 0 && <div className="muted">Допущений пока нет.</div>}
            <ol className="assumptions">
              {data.assumptions.map((a, i) => (
                <li key={a.id}>
                  <textarea rows={2} value={a.text} onChange={(e) => update((d) => ({ ...d, assumptions: d.assumptions.map((x, k) => (k === i ? { ...x, text: e.target.value } : x)) }))} />
                  <button className="icon-btn danger" onClick={() => update((d) => ({ ...d, assumptions: d.assumptions.filter((_, k) => k !== i) }))}>✕</button>
                </li>
              ))}
            </ol>
            <button className="add" onClick={() => update((d) => ({ ...d, assumptions: [...d.assumptions, { id: uid(), text: "" }] }))}>＋ Добавить допущение</button>
          </section>
          <section className="card">
            <h2>
              Holidays <small>Пропускаются при расчёте дат фаз (WORKDAY)</small>
              {catalog.defaultHolidays.length > 0 && (
                <button className="link-btn" onClick={() => update((d) => ({ ...d, holidays: [...catalog.defaultHolidays] }))}>
                  Сбросить на стандартные (США 2027–2028)
                </button>
              )}
            </h2>
            <div className="holidays">
              {data.holidays.map((h, i) => (
                <span key={i} className="chip" title={holidayName(h)}>
                  <input type="date" value={h} onChange={(e) => update((d) => ({ ...d, holidays: d.holidays.map((x, k) => (k === i ? e.target.value : x)) }))} />
                  <button className="icon-btn danger" onClick={() => update((d) => ({ ...d, holidays: d.holidays.filter((_, k) => k !== i) }))}>✕</button>
                </span>
              ))}
            </div>
            <button className="add" onClick={() => update((d) => ({ ...d, holidays: [...d.holidays, d.startDate || new Date().toISOString().slice(0, 10)] }))}>＋ Добавить дату</button>
          </section>
        </div>

        <details className="card help">
          <summary>Как заполнять</summary>
          <ol>
            <li><b>Team</b> — выберите Role, Grade и Tier для каждого участника; рейт подставится автоматически.</li>
            <li><b>Project start date</b> — дата начала только первой фазы.</li>
            <li><b>Service packages</b> — выберите блок и укажите трудозатраты (m/d) по фазам.</li>
            <li><b>Amount of working days</b> — длительность каждой фазы в рабочих днях.</li>
            <li><b>Roles engagement</b> — выберите блок и участника команды, укажите вовлечённость в % по фазам.</li>
            <li><b>Проверка</b> — «Total efforts amount by Roles» должен совпадать с «by Block»: зелёный — совпадает, красный — расходится.</li>
            <li><b>Contingency Reserve</b> и <b>Discount %</b>.</li>
          </ol>
        </details>
      </main>

      {showOpen && (
        <OpenDialog
          currentId={projectId}
          onOpen={openProject}
          onDeleted={(id) => {
            if (id === projectId) {
              setProjectId(null);
              setDirty(true);
              localStorage.removeItem("pe:lastProject");
            }
          }}
          onClose={() => setShowOpen(false)}
        />
      )}

      {showCreate && (
        <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && setShowCreate(false)}>
          <div className="modal small" role="dialog" aria-label="Создать проект">
            <div className="modal-head">
              <h2>Создать проект</h2>
              <button className="icon-btn" onClick={() => setShowCreate(false)} aria-label="Закрыть">✕</button>
            </div>
            <div className="create-options">
              <button onClick={() => createProject("blank")}>
                <b>Пустой проект</b>
                <span>Стандартные 5 фаз, пустые таблицы</span>
              </button>
              <button onClick={() => createProject("sample")}>
                <b>Из примера шаблона</b>
                <span>Данные из Project_Estimation_Template.xlsx</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function MiniGantt({ data, calc }: { data: ProjectData; calc: ReturnType<typeof calculate> }) {
  const toDay = (s: string | null) => (s ? Date.parse(s) / 86_400_000 : null);
  const s0 = toDay(calc.projectStart);
  const s1 = toDay(calc.projectEnd);
  if (s0 === null || s1 === null || s1 < s0) return null;
  const span = s1 - s0 + 1;
  return (
    <div className="gantt" aria-label="Project schedule">
      {PH.map((i) => {
        const a = toDay(calc.phaseStart[i]);
        const b = toDay(calc.phaseEnd[i]);
        return (
          <div className="g-row" key={i}>
            <div className="g-label" title={data.phases[i].name}>{data.phases[i].name}</div>
            <div className="g-track">
              {a !== null && b !== null && (
                <div
                  className={`g-bar c${i}`}
                  style={{ left: `${((a - s0) / span) * 100}%`, width: `${((b - a + 1) / span) * 100}%` }}
                  title={`${fmtDate(calc.phaseStart[i])} – ${fmtDate(calc.phaseEnd[i])}`}
                />
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}

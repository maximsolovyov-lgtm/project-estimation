import { useEffect, useMemo, useState } from "react";
import { api } from "../api";
import type { ProjectSummary } from "../shared/types";

interface Props {
  currentId: string | null;
  onOpen: (id: string) => void;
  onDeleted: (id: string) => void;
  onClose: () => void;
}

export function OpenDialog({ currentId, onOpen, onDeleted, onClose }: Props) {
  const [items, setItems] = useState<ProjectSummary[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [q, setQ] = useState("");

  useEffect(() => {
    api.list().then(setItems).catch((e: Error) => setErr(e.message));
  }, []);

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    return (items ?? []).filter((p) => !s || p.name.toLowerCase().includes(s) || p.client.toLowerCase().includes(s));
  }, [items, q]);

  async function remove(p: ProjectSummary) {
    if (!confirm(`Удалить проект «${p.name}»? Это действие нельзя отменить.`)) return;
    try {
      await api.remove(p.id);
      setItems((xs) => (xs ?? []).filter((x) => x.id !== p.id));
      onDeleted(p.id);
    } catch (e) {
      setErr((e as Error).message);
    }
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" role="dialog" aria-label="Выбор проекта">
        <div className="modal-head">
          <h2>Выбрать проект для редактирования</h2>
          <button className="icon-btn" onClick={onClose} aria-label="Закрыть">✕</button>
        </div>
        <input className="search" placeholder="Поиск по названию или клиенту…" value={q} onChange={(e) => setQ(e.target.value)} autoFocus />
        {err && <div className="alert bad">{err}</div>}
        {!items && !err && <div className="muted pad">Загрузка…</div>}
        {items && filtered.length === 0 && <div className="muted pad">Сохранённых проектов нет.</div>}
        <ul className="project-list">
          {filtered.map((p) => (
            <li key={p.id} className={p.id === currentId ? "current" : ""}>
              <button className="project-pick" onClick={() => onOpen(p.id)}>
                <span className="pname">{p.name}</span>
                <span className="pmeta">
                  {p.client && <>{p.client} · </>}изменён {new Date(p.updatedAt).toLocaleString("ru-RU")}
                </span>
              </button>
              <button className="icon-btn danger" title="Удалить" onClick={() => remove(p)}>🗑</button>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}

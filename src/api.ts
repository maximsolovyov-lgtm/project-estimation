import type { Catalog, ProjectData, ProjectSummary, RateQuery, StoredProject } from "./shared/types";

async function req<T>(url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(url, {
    ...init,
    headers: { "content-type": "application/json", ...(init?.headers ?? {}) },
  });
  if (!res.ok) {
    const body = (await res.json().catch(() => ({}))) as { error?: string };
    throw new Error(body.error || `HTTP ${res.status}`);
  }
  return res.json() as Promise<T>;
}

export const api = {
  catalog: () => req<Catalog>("/api/catalog"),
  rates: (items: RateQuery[]) =>
    req<{ rates: number[] }>("/api/rates", { method: "POST", body: JSON.stringify({ items }) }).then((r) => r.rates),
  list: () => req<ProjectSummary[]>("/api/projects"),
  get: (id: string) => req<StoredProject>(`/api/projects/${id}`),
  create: (data: ProjectData) => req<StoredProject>("/api/projects", { method: "POST", body: JSON.stringify({ data }) }),
  update: (id: string, data: ProjectData) =>
    req<StoredProject>(`/api/projects/${id}`, { method: "PUT", body: JSON.stringify({ data }) }),
  remove: (id: string) => req<{ ok: boolean }>(`/api/projects/${id}`, { method: "DELETE" }),
};

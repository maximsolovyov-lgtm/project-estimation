// Cloudflare Worker: REST API + static SPA (served from ./dist via the ASSETS binding).
// The full Rate Card lives only here; the browser receives roles/grades (no prices)
// and the rates of the exact Role+Grade+Tier combinations placed in the Team.

import { buildCatalog, lookupRate } from "./rates";
import { ASSETS } from "./assets.generated";
import type { ProjectData, ProjectSummary, RateQuery, StoredProject } from "../src/shared/types";
import { uniqueProjectName } from "../src/shared/naming";

interface Env {
  DB: D1Database;
}

const CATALOG = buildCatalog();

const json = (data: unknown, status = 200) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });
const error = (status: number, message: string) => json({ error: message }, status);

/** Re-resolve team rates on the server so stored rates always come from the Rate Card. */
function withServerRates(data: ProjectData): ProjectData {
  return { ...data, team: (data.team ?? []).map((m) => ({ ...m, rate: lookupRate(m) })) };
}

function validate(data: unknown): data is ProjectData {
  const d = data as ProjectData;
  return !!d && typeof d === "object" && Array.isArray(d.phases) && Array.isArray(d.team) &&
    Array.isArray(d.allocations) && Array.isArray(d.packages) && Array.isArray(d.blocks);
}

interface Row {
  id: string;
  name: string;
  data: string;
  created_at: string;
  updated_at: string;
}

const toSummary = (r: Row): ProjectSummary => {
  let client = "";
  try { client = (JSON.parse(r.data) as ProjectData).client ?? ""; } catch { /* ignore */ }
  return { id: r.id, name: r.name, client, createdAt: r.created_at, updatedAt: r.updated_at };
};
const toStored = (r: Row): StoredProject => ({ ...toSummary(r), data: withServerRates(JSON.parse(r.data)) });

async function api(req: Request, env: Env, url: URL): Promise<Response> {
  const path = url.pathname.replace(/\/+$/, "");
  const method = req.method;

  if (path === "/api/catalog" && method === "GET") return json(CATALOG);

  if (path === "/api/rates" && method === "POST") {
    const body = (await req.json().catch(() => null)) as { items?: RateQuery[] } | null;
    if (!body || !Array.isArray(body.items) || body.items.length > 200) return error(400, "items[] expected");
    return json({ rates: body.items.map(lookupRate) });
  }

  if (path === "/api/projects" && method === "GET") {
    const { results } = await env.DB.prepare(
      "SELECT id, name, data, created_at, updated_at FROM projects ORDER BY updated_at DESC",
    ).all<Row>();
    return json(results.map(toSummary));
  }

  if (path === "/api/projects" && method === "POST") {
    const body = (await req.json().catch(() => null)) as { data?: unknown; uniqueName?: boolean } | null;
    if (!body || !validate(body.data)) return error(400, "Invalid project data");
    let data = withServerRates(body.data);
    if (body.uniqueName) {
      // Import: "<name>" when free, otherwise "<name> v2", "<name> v3", …
      const { results } = await env.DB.prepare("SELECT name FROM projects").all<{ name: string }>();
      data = { ...data, name: uniqueProjectName(data.name, results.map((r) => r.name)) };
    }
    const id = crypto.randomUUID();
    const now = new Date().toISOString();
    await env.DB.prepare("INSERT INTO projects (id, name, data, created_at, updated_at) VALUES (?, ?, ?, ?, ?)")
      .bind(id, data.name || "Untitled", JSON.stringify(data), now, now)
      .run();
    return json({ id, name: data.name, client: data.client, createdAt: now, updatedAt: now, data } satisfies StoredProject, 201);
  }

  const m = /^\/api\/projects\/([\w-]+)$/.exec(path);
  if (m) {
    const id = m[1];
    if (method === "GET") {
      const row = await env.DB.prepare("SELECT * FROM projects WHERE id = ?").bind(id).first<Row>();
      return row ? json(toStored(row)) : error(404, "Project not found");
    }
    if (method === "PUT") {
      const body = (await req.json().catch(() => null)) as { data?: unknown } | null;
      if (!body || !validate(body.data)) return error(400, "Invalid project data");
      const data = withServerRates(body.data);
      const now = new Date().toISOString();
      const res = await env.DB.prepare("UPDATE projects SET name = ?, data = ?, updated_at = ? WHERE id = ?")
        .bind(data.name || "Untitled", JSON.stringify(data), now, id)
        .run();
      if (!res.meta.changes) return error(404, "Project not found");
      const row = await env.DB.prepare("SELECT * FROM projects WHERE id = ?").bind(id).first<Row>();
      return json(toStored(row!));
    }
    if (method === "DELETE") {
      await env.DB.prepare("DELETE FROM projects WHERE id = ?").bind(id).run();
      return json({ ok: true });
    }
  }

  return error(404, "Not found");
}

/** Serves the embedded SPA; unknown paths fall back to index.html (client-side routing). */
function serveAsset(pathname: string): Response {
  const hit = ASSETS[pathname];
  const file = hit ?? ASSETS["/index.html"];
  if (!file) return new Response("Frontend not built", { status: 500 });
  const immutable = !!hit && pathname.startsWith("/assets/");
  return new Response(file.body, {
    headers: {
      "content-type": file.type,
      "cache-control": immutable ? "public, max-age=31536000, immutable" : "no-cache",
      "x-content-type-options": "nosniff",
    },
  });
}

export default {
  async fetch(req: Request, env: Env): Promise<Response> {
    const url = new URL(req.url);
    if (url.pathname.startsWith("/api/")) {
      try {
        return await api(req, env, url);
      } catch (e) {
        console.error(e);
        return error(500, e instanceof Error ? e.message : "Internal error");
      }
    }
    return serveAsset(url.pathname);
  },
} satisfies ExportedHandler<Env>;

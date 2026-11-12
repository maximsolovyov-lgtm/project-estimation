/**
 * Name for a newly imported project: the name itself when it is free, otherwise
 * "<name> v2", "<name> v3", … (a trailing " vN" of the incoming name is treated as the base).
 */
export function uniqueProjectName(name: string, existing: string[]): string {
  const taken = new Set(existing.map((n) => n.trim().toLowerCase()));
  const clean = name.trim() || "Imported project";
  if (!taken.has(clean.toLowerCase())) return clean;
  const base = clean.replace(/\s+v\d+$/i, "");
  for (let n = 2; ; n++) {
    const candidate = `${base} v${n}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}

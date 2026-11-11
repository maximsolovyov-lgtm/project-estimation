import { DEFAULT_PHASE_COUNT, type Allocation, type ProjectData, type TeamMember, type WorkPackage } from "./types";

export const DEFAULT_PHASES = [
  "Discovery & Design",
  "Build",
  "Testing & Validation (UAT)",
  "Closure & Training",
  "Go-Live & Hypercare",
];

export function uid(): string {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
}

export function emptyPhaseArray(n: number = DEFAULT_PHASE_COUNT): number[] {
  return Array(n).fill(0);
}

export function newTeamMember(): TeamMember {
  return { id: uid(), role: "", grade: "", tier: "", rate: 0 };
}

export function newAllocation(phases: number = DEFAULT_PHASE_COUNT): Allocation {
  return { id: uid(), block: "", role: "", grade: "", tier: "", fte: emptyPhaseArray(phases) };
}

export function newPackage(phases: number = DEFAULT_PHASE_COUNT): WorkPackage {
  return { id: uid(), block: "", packageId: "", name: "", days: emptyPhaseArray(phases) };
}

function todayIso(): string {
  return new Date().toISOString().slice(0, 10);
}

export function blankProject(holidays: string[] = []): ProjectData {
  return {
    name: "New project",
    client: "",
    startDate: todayIso(),
    phases: DEFAULT_PHASES.map((name) => ({ name, workingDays: 0 })),
    contingency: 0,
    discount: 0,
    blocks: [""],
    team: [newTeamMember()],
    allocations: [newAllocation()],
    packages: [newPackage()],
    holidays: [...holidays],
    assumptions: [],
  };
}

/** The example filled in on Project_Estimation_Template.xlsx. Rates are resolved by the server. */
export function sampleProject(holidays: string[] = []): ProjectData {
  const t = (role: string, grade: string, tier: string): TeamMember => ({ id: uid(), role, grade, tier, rate: 0 });
  const a = (block: string, role: string, grade: string, tier: string, fte: number[]): Allocation => ({
    id: uid(), block, role, grade, tier, fte,
  });
  const w = (block: string, packageId: string, name: string, days: number[]): WorkPackage => ({
    id: uid(), block, packageId, name, days,
  });
  const team = [
    t("Salesforce Architect", "Technical", "Tier 4"),
    t("AI Agent", "Senior", "Tier 5"),
    t("Integration", "Lead", "Tier 3"),
    t("SPIFF Project Manager", "Middle", "Tier 4"),
    t("SPIFF Configuration Consultant", "Middle", "Tier 4"),
    t("SPIFF Developer", "Middle", "Tier 5"),
    t("Salesforce Business Analyst/Consultant", "Senior", "Tier 4"),
    t("QA Manual", "Middle", "Tier 5"),
    t("Integration developer", "Senior", "Tier 4"),
    t("BI Consultant", "Senior", "Tier 4"),
  ];
  const allocations = [
    a("All", "Salesforce Architect", "Technical", "Tier 4", [1, 0.5, 0.2, 0.2, 0.2]),
    a("All", "SPIFF Project Manager", "Middle", "Tier 4", [1, 1, 0.5, 0.5, 0.5]),
    a("SPIFF", "SPIFF Consultant", "Lead", "Tier 4", [1, 1, 0.3, 0.2, 0.3]),
    a("SPIFF", "AI Agent", "Senior", "Tier 5", [0.5, 1.5, 0.2, 0.2, 0.2]),
    a("SPIFF", "SPIFF Developer", "Middle", "Tier 5", [0.5, 1, 0.2, 0, 0.2]),
    a("SPIFF", "QA Manual", "Middle", "Tier 5", [0, 0.4, 0.6, 0, 0]),
    a("Integration", "Integration", "Lead", "Tier 3", [1, 0.5, 0, 0.1, 0.1]),
    a("Integration", "Integration developer", "Senior", "Tier 4", [0.2, 1, 0, 0, 0.1]),
    a("Sales Planning", "Salesforce Business Analyst/Consultant", "Senior", "Tier 4", [1.5, 0.5, 0, 0.2, 0]),
    a("Sales Planning", "BI Consultant", "Senior", "Tier 4", [0.5, 1, 0, 0, 0]),
  ];
  for (const al of allocations) {
    al.memberId = team.find((m) => m.role === al.role && m.grade === al.grade && m.tier === al.tier)?.id;
  }
  return {
    name: "Project Estimation Template (example)",
    client: "",
    startDate: "2026-03-31",
    phases: DEFAULT_PHASES.map((name, i) => ({ name, workingDays: [10, 20, 15, 10, 15][i] })),
    contingency: 3,
    discount: 0.1,
    blocks: ["SPIFF", "Maps", "Sales Planning", "Sales Program", "Integration"],
    team,
    allocations,
    packages: [
      w("SPIFF", "SPIFF01", "Comission Plan 1", [10, 15, 4, 5, 5]),
      w("SPIFF", "SPIFF02", "Comission Plan 2", [8, 5, 4, 5, 6]),
      w("SPIFF", "SPIFF03", "Workflow", [8, 3, 4, 0, 6]),
      w("SPIFF", "SPIFF04", "Exports capabilities", [10, 2, 3, 0, 1]),
      w("SPIFF", "RP1", "Report: 1", [7, 1, 0, 0, 0]),
      w("SPIFF", "RP2", "Report 2", [7, 1, 0, 0, 0]),
      w("Integration", "INT1", "Integration: Workday", [7, 1, 0, 0, 0]),
      w("SPIFF", "NF1", "Non Functional: Branding ", [7, 1, 0, 0, 0]),
      w("Sales Planning", "MM1", "Wave 2 · Pipeline PRODUCT_COMPLIANCE_REFRESH — Gen3-archetype redesign", [0, 109, 0, 0, 0]),
      w("Integration", "INT2", "Integration: Salesforce ↔ ERP", [3, 10, 6, 1, 2]),
      w("SPIFF", "SPIFF05", "Data migration & historical load", [2, 8, 4, 1, 2]),
      w("Sales Planning", "SP1", "Territory & quota planning", [3, 12, 5, 2, 2]),
    ],
    holidays: [...holidays],
    assumptions: [],
  };
}

// Shared data model for the frontend and the Worker.
// Mirrors the "Project Estimation" sheet of Project_Estimation_Template.xlsx.

/** A new project always starts with 5 phases; the phase manager can add/remove phases. */
export const DEFAULT_PHASE_COUNT = 5;
export const MIN_PHASES = 1;
export const MAX_PHASES = 12;
export const ALL_BLOCK = "All";
export const TIERS = ["Tier 1", "Tier 2", "Tier 3", "Tier 4", "Tier 5"] as const;

/** Team row (Project Estimation!I4:N13). `rate` is resolved server-side from the hidden Rate Card. */
export interface TeamMember {
  id: string;
  role: string;
  grade: string;
  tier: string;
  rate: number;
}

/** Roles & engagement row (Project Estimation!A23:N34). `fte` = engagement per phase, 1 = 100% = 1 FTE. */
export interface Allocation {
  id: string;
  block: string;
  /** Team member the row was picked from (UI convenience; calculations match role+grade+tier like Excel). */
  memberId?: string;
  role: string;
  grade: string;
  tier: string;
  fte: number[];
}

/** Entry table row (Project Estimation!A37:N…): service package, man-days per phase. */
export interface WorkPackage {
  id: string;
  block: string;
  packageId: string;
  name: string;
  days: number[];
}

export interface Phase {
  name: string;
  workingDays: number;
}

export interface Assumption {
  id: string;
  text: string;
}

export interface ProjectData {
  name: string;
  client: string;
  /** Start date of the FIRST phase, ISO yyyy-mm-dd (Project Estimation!I18). */
  startDate: string;
  phases: Phase[];
  /** Contingency Reserve, m/d (F5). */
  contingency: number;
  /** Discount, fraction (F10). 0.1 = 10%. */
  discount: number;
  /** Block names (P3:P14). "All" (P15) is implicit. */
  blocks: string[];
  team: TeamMember[];
  allocations: Allocation[];
  packages: WorkPackage[];
  /** Holidays skipped by WORKDAY (hidden Holidays sheet), ISO dates. */
  holidays: string[];
  assumptions: Assumption[];
}

export interface ProjectSummary {
  id: string;
  name: string;
  client: string;
  createdAt: string;
  updatedAt: string;
}

export interface StoredProject extends ProjectSummary {
  data: ProjectData;
}

export interface CatalogRole {
  role: string;
  grades: string[];
}

export interface Catalog {
  roles: CatalogRole[];
  tiers: string[];
  defaultHolidays: string[];
}

export interface RateQuery {
  role: string;
  grade: string;
  tier: string;
}

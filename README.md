# Project Estimation (Cloudflare)

A web app for building project estimates. It reproduces the logic of `Project_Estimation_Template.xlsx`
(the "Project Estimation" sheet) and runs on Cloudflare Workers with a D1 database.

## Features

| Button | What it does |
|---|---|
| **Create project** | New project: blank, or prefilled with the example from the template |
| **Save project** | Saves the project to D1 (creates it or updates it) |
| **Open project** | List of saved projects with search, open and delete |
| **Export to Excel** | `.xlsx` with the same layout and live formulas as the template (WORKDAY, SUMPRODUCT, SUMIF, uplift for "All", conditional formatting for the checks, drop-downs) |
| **Generate presentation** | `.pptx` with one element per slide: title, Total for the project, Team, Workload, **Gantt chart of the schedule**, Phases, Roles engagement, Block Estimation, Phase Estimation, Service packages, Cost structure, Consistency check, Assumptions |

### The Rate Card stays hidden
* The Rate Card (`worker/rateCard.json`) is bundled **only into the Worker**. It is not in the frontend bundle.
* Rate Card prices are hourly, estimates are in man-days, so the app uses a **daily rate = hourly rate × 8**
  (`HOURS_PER_DAY` in `src/shared/types.ts`). The Team table, cost, Blended Rate, Excel and the deck all use it.
* The browser gets the list of roles and grades without prices (`GET /api/catalog`), plus the rates for the
  Role+Grade+Tier combinations placed in Team (`POST /api/rates`).
* When a project is saved or loaded, the server recomputes the rates from the Rate Card, so a rate
  changed on the client is ignored.
* The Excel export does not include the Rate Card sheet. The Rate column holds only the team's rates, as values.

## Phases

A new project always starts with the 5 standard phases (Discovery & Design, Build, Testing & Validation (UAT),
Closure & Training, Go-Live & Hypercare). The **Phases** block lets you rename phases, set working days, add
(up to 12), insert, move, delete (at least 1) and reset to the 5 standard phases. Engagement % and package man-days
move with their phase. The Excel export and the presentation adapt to any number of phases.

## Calculation logic

`src/shared/calc.ts` is a direct port of the template formulas. Each value is labelled with its source cell (F4…F14, L/M/N4:13,
I15:N15, I18:N22, N23:N34, I36:N36, Q3:Q14, V3:V14). `test/calc.test.ts` checks the results against the values Excel
calculated in the original file.

One difference from the template: on the template, the Phase Estimation cells V6/V7 summed the wrong columns,
so "Closure & Training" and "Go-Live & Hypercare" showed each other's values. The app gives each phase its own column.

## Running locally

```bash
npm install
```

```bash
npm run db:migrate:local
```

```bash
npm run preview   # build + wrangler dev → http://localhost:8787
```

```bash
npm test
```

## Deploying to Cloudflare

**Live URL: https://project-estimation.maxim-solovyov.workers.dev**

The account, D1 database (`project-estimation`) and its id are already set in `wrangler.jsonc`, and the migration has been applied.
To deploy again:

```bash
npm run deploy
```

`npm run build` builds the frontend and embeds it into the Worker (`worker/assets.generated.ts`), so a deploy is
a single script upload with no separate Static Assets upload. For a new schema migration: `npm run db:migrate:remote`.

The app has no built-in authentication. To restrict access, turn on Cloudflare Access (Zero Trust) for the Worker's domain.

## Updating the Rate Card

```bash
python3 scripts/extract_rate_card.py path/to/Project_Estimation_Template.xlsx
```

The script regenerates `worker/rateCard.json` (Fix-Price / A-la-Carte tiers plus the default holidays). Run `npm run deploy` afterwards.

## Structure

```
worker/            Cloudflare Worker: API, D1, Rate Card (server-side only)
src/shared/        data model, calculation engine, WORKDAY, templates
src/App.tsx        estimation form UI
src/export/        Excel (exceljs) and PowerPoint (pptxgenjs) generation in the browser
migrations/        D1 schema
test/              tests for the calculations and the exports
```

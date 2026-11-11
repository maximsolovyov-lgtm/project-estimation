// Default holidays for new projects: US federal holidays 2027–2028 (observed dates —
// a holiday on Saturday is observed on Friday, on Sunday — on Monday).
// Skipped by WORKDAY when phase dates are calculated.

export const DEFAULT_HOLIDAYS: { date: string; name: string }[] = [
  { date: "2027-01-01", name: "New Year's Day" },
  { date: "2027-01-18", name: "Martin Luther King Jr. Day" },
  { date: "2027-02-15", name: "Washington's Birthday" },
  { date: "2027-05-31", name: "Memorial Day" },
  { date: "2027-06-18", name: "Juneteenth (observed)" },
  { date: "2027-07-05", name: "Independence Day (observed)" },
  { date: "2027-09-06", name: "Labor Day" },
  { date: "2027-10-11", name: "Columbus Day" },
  { date: "2027-11-11", name: "Veterans Day" },
  { date: "2027-11-25", name: "Thanksgiving Day" },
  { date: "2027-12-24", name: "Christmas Day (observed)" },
  { date: "2027-12-31", name: "New Year's Day 2028 (observed)" },
  { date: "2028-01-17", name: "Martin Luther King Jr. Day" },
  { date: "2028-02-21", name: "Washington's Birthday" },
  { date: "2028-05-29", name: "Memorial Day" },
  { date: "2028-06-19", name: "Juneteenth" },
  { date: "2028-07-04", name: "Independence Day" },
  { date: "2028-09-04", name: "Labor Day" },
  { date: "2028-10-09", name: "Columbus Day" },
  { date: "2028-11-10", name: "Veterans Day (observed)" },
  { date: "2028-11-23", name: "Thanksgiving Day" },
  { date: "2028-12-25", name: "Christmas Day" },
];

const NAMES = new Map(DEFAULT_HOLIDAYS.map((h) => [h.date, h.name]));

export function holidayName(iso: string): string {
  return NAMES.get(iso) ?? "";
}

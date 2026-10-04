import * as XLSX from "xlsx";

export type Contractor = {
  id: string;
  name: string;
  email: string | null;
  department: string | null;
  function: string | null;
  role: string | null;
  work_type: string;
  contract_type: string | null;
  country: string | null;
  vendor: string | null;
  monthly_rate: number | null;
  currency: string | null;
  renewal_count: number;
  sow_name: string | null;
  sow_start_date: string | null;
  sow_end_date: string | null;
  termination_date: string | null;
  addis_status: string | null;
  manager: string | null;
  notes: string | null;
  created_at?: string;
  updated_at?: string;
};

export type ContractorInput = Omit<Contractor, "id" | "created_at" | "updated_at">;

export type Status = "Terminated" | "Expired" | "Expiring" | "Active";

export const EXPIRY_WINDOW_DAYS = 180;
export const WORK_TYPES = ["Contractor", "FTE"] as const;

export const FIELDS: { key: keyof ContractorInput; label: string; type: "text" | "date" | "number" }[] = [
  { key: "name", label: "Name", type: "text" },
  { key: "email", label: "Email", type: "text" },
  { key: "work_type", label: "Work Type", type: "text" },
  { key: "department", label: "Department", type: "text" },
  { key: "function", label: "Function", type: "text" },
  { key: "role", label: "Role", type: "text" },
  { key: "country", label: "Country", type: "text" },
  { key: "vendor", label: "Vendor", type: "text" },
  { key: "sow_name", label: "SoW Name", type: "text" },
  { key: "sow_start_date", label: "SoW Start Date", type: "date" },
  { key: "sow_end_date", label: "SoW / Project End Date", type: "date" },
  { key: "termination_date", label: "Termination Date", type: "date" },
  { key: "contract_type", label: "Contract Type", type: "text" },
  { key: "monthly_rate", label: "Monthly Rate", type: "number" },
  { key: "currency", label: "Currency", type: "text" },
  { key: "renewal_count", label: "Renewals", type: "number" },
  { key: "addis_status", label: "ADDIS Status", type: "text" },
  { key: "manager", label: "Direct Manager", type: "text" },
  { key: "notes", label: "Notes", type: "text" },
];

const today = () => {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
};

export function parseDate(value: string | null | undefined): Date | null {
  if (!value) return null;
  const d = new Date(value + (value.length === 10 ? "T00:00:00" : ""));
  return isNaN(d.getTime()) ? null : d;
}

export function daysUntil(value: string | null | undefined): number | null {
  const d = parseDate(value);
  if (!d) return null;
  return Math.round((d.getTime() - today().getTime()) / 86400000);
}

export function getStatus(c: Contractor): Status {
  const term = daysUntil(c.termination_date);
  if (term !== null && term <= 0) return "Terminated";
  const end = daysUntil(c.sow_end_date);
  if (end === null) return "Active";
  if (end < 0) return "Expired";
  if (end <= EXPIRY_WINDOW_DAYS) return "Expiring";
  return "Active";
}

/** Still flagged as in the internal system, but the SoW already ended. */
export function needsAttention(c: Contractor): boolean {
  const status = getStatus(c);
  const inSystem = (c.addis_status ?? "").trim().toLowerCase();
  const stillIn = inSystem === "" ? false : !/(inactive|removed|disabled|terminated|no)$/.test(inSystem);
  return status === "Expired" && stillIn;
}

export const STATUS_STYLES: Record<Status, string> = {
  Active: "bg-emerald-500/15 text-emerald-700 dark:text-emerald-300",
  Expiring: "bg-amber-500/20 text-amber-700 dark:text-amber-300",
  Expired: "bg-destructive/15 text-destructive",
  Terminated: "bg-muted text-muted-foreground",
};

/* ------------------------------ spreadsheet ------------------------------ */

export type ParsedSheet = { headers: string[]; rows: Record<string, unknown>[] };

export async function readSheet(file: File): Promise<ParsedSheet> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0] ?? ""];
  if (!sheet) return { headers: [], rows: [] };
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: null, raw: false });
  const headerRow = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, blankrows: false })[0] ?? [];
  const headers = headerRow.map((h) => String(h ?? "").trim()).filter(Boolean);
  return { headers, rows };
}

function normalize(s: string) {
  return s.toLowerCase().replace(/[^a-z0-9]/g, "");
}

const HINTS: Record<keyof ContractorInput, string[]> = {
  name: ["name", "fullname", "contractor", "contractorname", "employeename"],
  email: ["email", "emailaddress", "mail", "workemail"],
  work_type: ["worktype", "workertype", "employmenttype", "employeetype", "fteorcontractor"],
  department: ["department", "dept", "team"],
  function: ["function", "jobfunction"],
  role: ["role", "jobtitle", "title", "position", "jobrole"],
  country: ["country", "location", "region", "site"],
  vendor: ["vendor", "supplier", "agency", "vendorname", "suppliername", "staffingpartner"],
  sow_name: ["sowname", "sow", "statementofwork", "projectname"],
  sow_start_date: ["sowstartdate", "startdate", "start", "contractstart", "projectstartdate"],
  sow_end_date: ["sowenddate", "projectenddate", "projectend", "enddate", "end", "contractend", "expirydate", "expirationdate"],
  termination_date: ["terminationdate", "terminatedate", "termdate", "exitdate", "lastworkingday"],
  contract_type: ["contracttype", "agreementtype", "engagementtype"],
  monthly_rate: ["monthlyrate", "rate", "monthlyfee", "fee", "monthrate"],
  currency: ["currency", "curr", "ratecurrency", "currencycode"],
  renewal_count: ["renewalcount", "renewals", "renewaltimes", "extensions", "renewed"],
  addis_status: ["addisstatus", "addis", "systemstatus", "accountstatus", "status"],
  manager: ["directmanager", "manager", "linemanager", "reportingmanager", "supervisor", "owner"],
  notes: ["notes", "comment", "comments", "remarks"],
};

export function guessMapping(headers: string[]): Record<string, string> {
  const map: Record<string, string> = {};
  const used = new Set<string>();
  for (const f of FIELDS) {
    const hints = HINTS[f.key];
    const found =
      headers.find((h) => !used.has(h) && hints.includes(normalize(h))) ??
      headers.find((h) => !used.has(h) && hints.some((x) => normalize(h).includes(x)));
    if (found) {
      map[f.key] = found;
      used.add(found);
    }
  }
  return map;
}

export function toISODate(value: unknown): string | null {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  const raw = String(value).trim();
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);
  // dd/mm/yyyy or mm/dd/yyyy — assume day first when the first part > 12
  const m = raw.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
  if (m) {
    const a = m[1] ?? "";
    const b = m[2] ?? "";
    const y = m[3] ?? "";
    let day = Number(a);
    let month = Number(b);
    if (day <= 12 && month > 12) {
      [day, month] = [month, day];
    }
    const year = y.length === 2 ? 2000 + Number(y) : Number(y);
    const d = new Date(Date.UTC(year, month - 1, day));
    return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }
  const serial = Number(raw);
  if (!isNaN(serial) && serial > 20000 && serial < 60000) {
    const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    return d.toISOString().slice(0, 10);
  }
  const parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed.toISOString().slice(0, 10);
}

export type MappedRow = { record: ContractorInput; issues: string[] };

export function normalizeWorkType(v: string | null | undefined, fallback: string): string {
  const s = (v ?? "").trim().toLowerCase();
  if (!s) return fallback;
  if (/fte|full|perm|employee|staff/.test(s)) return "FTE";
  if (/contract|temp|freelanc|consult/.test(s)) return "Contractor";
  return v!.trim();
}

export function mapRows(
  rows: Record<string, unknown>[],
  mapping: Record<string, string>,
  defaultWorkType = "Contractor",
): MappedRow[] {
  return rows.map((row) => {
    const issues: string[] = [];
    const record = {} as ContractorInput;
    for (const f of FIELDS) {
      const col = mapping[f.key];
      const raw = col ? row[col] : null;
      if (f.type === "date") {
        const iso = toISODate(raw);
        if (raw && !iso) issues.push(`${f.label} not a valid date`);
        record[f.key] = iso as never;
      } else if (f.type === "number") {
        const rawStr = raw === null || raw === undefined ? "" : String(raw).trim();
        if (!rawStr) {
          record[f.key] = (f.key === "renewal_count" ? 0 : null) as never;
        } else {
          const n = Number(rawStr.replace(/[,$\s]/g, ""));
          if (isNaN(n)) issues.push(`${f.label} not a valid number`);
          record[f.key] = (isNaN(n) ? (f.key === "renewal_count" ? 0 : null) : n) as never;
        }
      } else {
        const v = raw === null || raw === undefined ? null : String(raw).trim();
        record[f.key] = (v === "" ? null : v) as never;
      }
    }
    record.work_type = normalizeWorkType(record.work_type, defaultWorkType);
    if (!record.name) issues.push("Missing name");
    return { record, issues };
  });
}

/* ------------------------------ table columns ----------------------------- */

export type ColumnDef = { key: string; label: string; value: (c: Contractor) => string | number };

export const COLUMNS: ColumnDef[] = [
  { key: "name", label: "Name", value: (c) => c.name },
  { key: "email", label: "Email", value: (c) => c.email ?? "" },
  { key: "work_type", label: "Work Type", value: (c) => c.work_type ?? "" },
  { key: "department", label: "Department", value: (c) => c.department ?? "" },
  { key: "function", label: "Function", value: (c) => c.function ?? "" },
  { key: "role", label: "Role", value: (c) => c.role ?? "" },
  { key: "manager", label: "Direct Manager", value: (c) => c.manager ?? "" },
  { key: "contract_type", label: "Contract Type", value: (c) => c.contract_type ?? "" },
  { key: "country", label: "Country", value: (c) => c.country ?? "" },
  { key: "vendor", label: "Vendor", value: (c) => c.vendor ?? "" },
  { key: "monthly_rate", label: "Monthly Rate", value: (c) => c.monthly_rate ?? "" },
  { key: "currency", label: "Currency", value: (c) => c.currency ?? "" },
  { key: "sow_name", label: "SoW Name", value: (c) => c.sow_name ?? "" },
  { key: "sow_start_date", label: "SoW Start", value: (c) => c.sow_start_date ?? "" },
  { key: "sow_end_date", label: "SoW / Project End", value: (c) => c.sow_end_date ?? "" },
  { key: "termination_date", label: "Termination", value: (c) => c.termination_date ?? "" },
  { key: "renewal_count", label: "Renewals", value: (c) => c.renewal_count ?? 0 },
  { key: "addis_status", label: "ADDIS", value: (c) => c.addis_status ?? "" },
  { key: "notes", label: "Notes", value: (c) => c.notes ?? "" },
  { key: "status", label: "Status", value: (c) => getStatus(c) },
  { key: "days", label: "Days To Expiry", value: (c) => daysUntil(c.sow_end_date) ?? "" },
];

export const DEFAULT_COLUMNS = [
  "name",
  "work_type",
  "role",
  "manager",
  "contract_type",
  "function",
  "country",
  "vendor",
  "sow_end_date",
  "termination_date",
  "addis_status",
  "status",
];

export function exportToExcel(contractors: Contractor[], filename = "contractors.xlsx") {
  const data = contractors.map((c) =>
    Object.fromEntries(COLUMNS.map((col) => [col.label, col.value(c)])),
  );
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = COLUMNS.map(() => ({ wch: 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Contractors");
  XLSX.writeFile(wb, filename);
}

/** Multi-sheet report: summary + breakdowns + details using chosen columns. */
export function exportReport(
  rows: Contractor[],
  opts: { title: string; filters: [string, string][]; columns: string[]; filename?: string },
) {
  const statusCount = (s: Status) => rows.filter((c) => getStatus(c) === s).length;
  const within = (n: number) =>
    rows.filter((c) => {
      const d = daysUntil(c.sow_end_date);
      return d !== null && d >= 0 && d <= n && getStatus(c) !== "Terminated";
    }).length;
  const summary: (string | number)[][] = [
    ["Report", opts.title],
    ["Generated", new Date().toISOString().slice(0, 16).replace("T", " ")],
    [],
    ["Filters"],
    ...(opts.filters.length ? opts.filters : [["(none)", ""]]),
    [],
    ["Metric", "Count"],
    ["Total", rows.length],
    ["Active (incl. expiring)", statusCount("Active") + statusCount("Expiring")],
    ["Ending within 90 days", within(90)],
    ["Ending within 180 days", within(180)],
    ["Expired", statusCount("Expired")],
    ["Terminated", statusCount("Terminated")],
    ["SoW ended, still in system", rows.filter(needsAttention).length],
  ];
  const breakdown: (string | number)[][] = [];
  const dims: [string, (c: Contractor) => string][] = [
    ["Work Type", (c) => c.work_type],
    ["Function", (c) => c.function ?? ""],
    ["Country", (c) => c.country ?? ""],
    ["Vendor", (c) => c.vendor ?? ""],
    ["Status", (c) => getStatus(c)],
  ];
  for (const [label, get] of dims) {
    const m = new Map<string, number>();
    for (const c of rows) {
      const k = get(c).trim() || "Unspecified";
      m.set(k, (m.get(k) ?? 0) + 1);
    }
    breakdown.push([label, "Count"]);
    [...m].sort((a, b) => b[1] - a[1]).forEach(([k, v]) => breakdown.push([k, v]));
    breakdown.push([]);
  }
  const cols = COLUMNS.filter((c) => opts.columns.includes(c.key));
  const details = rows.map((c) => Object.fromEntries(cols.map((col) => [col.label, col.value(c)])));

  const wb = XLSX.utils.book_new();
  const s1 = XLSX.utils.aoa_to_sheet(summary);
  s1["!cols"] = [{ wch: 30 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, s1, "Summary");
  const s2 = XLSX.utils.aoa_to_sheet(breakdown);
  s2["!cols"] = [{ wch: 28 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, s2, "Breakdown");
  const s3 = XLSX.utils.json_to_sheet(details.length ? details : [{}]);
  s3["!cols"] = cols.map(() => ({ wch: 18 }));
  XLSX.utils.book_append_sheet(wb, s3, "Details");
  XLSX.writeFile(wb, opts.filename ?? "report.xlsx");
}

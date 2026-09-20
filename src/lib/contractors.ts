import * as XLSX from "xlsx";

export type Contractor = {
  id: string;
  name: string;
  email: string | null;
  department: string | null;
  function: string | null;
  country: string | null;
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

export const FIELDS: { key: keyof ContractorInput; label: string; type: "text" | "date" }[] = [
  { key: "name", label: "Name", type: "text" },
  { key: "email", label: "Email", type: "text" },
  { key: "department", label: "Department", type: "text" },
  { key: "function", label: "Function", type: "text" },
  { key: "country", label: "Country", type: "text" },
  { key: "sow_name", label: "SoW Name", type: "text" },
  { key: "sow_start_date", label: "SoW Start Date", type: "date" },
  { key: "sow_end_date", label: "SoW End Date", type: "date" },
  { key: "termination_date", label: "Termination Date", type: "date" },
  { key: "addis_status", label: "ADDIS Status", type: "text" },
  { key: "manager", label: "Manager", type: "text" },
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
  department: ["department", "dept", "team"],
  function: ["function", "role", "jobfunction", "position", "title"],
  country: ["country", "location", "region", "site"],
  sow_name: ["sowname", "sow", "statementofwork", "projectname", "project"],
  sow_start_date: ["sowstartdate", "startdate", "start", "contractstart"],
  sow_end_date: ["sowenddate", "enddate", "end", "contractend", "expirydate", "expirationdate"],
  termination_date: ["terminationdate", "terminatedate", "termdate", "exitdate", "lastworkingday"],
  addis_status: ["addisstatus", "addis", "systemstatus", "accountstatus", "status"],
  manager: ["manager", "linemanager", "reportingmanager", "supervisor", "owner"],
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

export function mapRows(rows: Record<string, unknown>[], mapping: Record<string, string>): MappedRow[] {
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
      } else {
        const v = raw === null || raw === undefined ? null : String(raw).trim();
        record[f.key] = (v === "" ? null : v) as never;
      }
    }
    if (!record.name) issues.push("Missing name");
    return { record, issues };
  });
}

export function exportToExcel(contractors: Contractor[], filename = "contractors.xlsx") {
  const data = contractors.map((c) => ({
    Name: c.name,
    Email: c.email ?? "",
    Department: c.department ?? "",
    Function: c.function ?? "",
    Country: c.country ?? "",
    "SoW Name": c.sow_name ?? "",
    "SoW Start Date": c.sow_start_date ?? "",
    "SoW End Date": c.sow_end_date ?? "",
    "Termination Date": c.termination_date ?? "",
    "ADDIS Status": c.addis_status ?? "",
    Manager: c.manager ?? "",
    Notes: c.notes ?? "",
    Status: getStatus(c),
    "Days To Expiry": daysUntil(c.sow_end_date) ?? "",
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = Object.keys(data[0] ?? { Name: "" }).map(() => ({ wch: 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Contractors");
  XLSX.writeFile(wb, filename);
}

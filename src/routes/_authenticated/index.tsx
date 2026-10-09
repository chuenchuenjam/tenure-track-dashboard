import { useEffect, useMemo, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import {
  Bar,
  BarChart,
  Cell,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import {
  Download,
  Plus,
  Upload,
  Users,
  ShieldCheck,
  CalendarClock,
  CalendarX,
  AlertTriangle,
  Wallet,
  X,
  Search,
  Columns3,
  LayoutGrid,
  History,
  FileText,
  Save,
  Trash2,
} from "lucide-react";
import { toast } from "sonner";
import * as XLSX from "xlsx";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import {
  DropdownMenu,
  DropdownMenuCheckboxItem,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { supabase } from "@/integrations/supabase/client";
import { UploadDialog } from "@/components/contractors/UploadDialog";
import { ContractorSheet } from "@/components/contractors/ContractorSheet";
import {
  COLUMNS,
  DEFAULT_COLUMNS,
  STATUS_STYLES,
  daysUntil,
  exportReport,
  exportToExcel,
  getStatus,
  needsAttention,
  type Contractor,
  type Status,
} from "@/lib/contractors";

export const Route = createFileRoute("/_authenticated/")({
  head: () => ({
    meta: [
      { title: "Contractor SoW Dashboard | HR Contract Tracker" },
      {
        name: "description",
        content:
          "Upload contractor and FTE lists, track SoW expiry within 90 and 180 days, resolve conflicts, build reports and export to Excel.",
      },
      { property: "og:title", content: "Contractor SoW Dashboard | HR Contract Tracker" },
      {
        property: "og:description",
        content:
          "Upload contractor and FTE lists, track SoW expiry within 90 and 180 days, resolve conflicts, build reports and export to Excel.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Dashboard,
});

const ALL = "__all__";
const COLS_KEY = "dashboard.columns.v1";
function downloadUpload(name: string, rows: Record<string, unknown>[], overrides: Record<string, unknown>[]) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), "Uploaded rows");
  if (overrides.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(overrides), "Conflicts");
  XLSX.writeFile(wb, `upload-${name.replace(/\.[^.]+$/, "")}.xlsx`);
}

const WIDGETS_KEY = "dashboard.widgets.v1";
const WIDGETS: { key: string; label: string }[] = [
  { key: "kpi:Total", label: "Total (card)" },
  { key: "kpi:Active", label: "Active (card)" },
  { key: "kpi:Ending within 90 days", label: "Ending within 90 days (card)" },
  { key: "kpi:Ending within 180 days", label: "Ending within 180 days (card)" },
  { key: "kpi:SoW ended, still in system", label: "SoW ended, still in system (card)" },
  { key: "kpi:Monthly spend (active)", label: "Monthly spend (active) (card)" },
  { key: "chart:Monthly resource count by vendor", label: "Monthly resource count by vendor" },
  { key: "chart:Headcount trend", label: "Headcount trend" },
  { key: "chart:Monthly spend by vendor", label: "Monthly spend by vendor" },
  { key: "chart:Terminations by year", label: "Terminations by year" },
  { key: "chart:Upcoming expiries by month", label: "Upcoming expiries by month" },
  { key: "chart:Split by function", label: "Split by function" },
  { key: "chart:Split by country", label: "Split by country" },
  { key: "chart:Renewal queue", label: "Renewal queue" },
  { key: "uploads", label: "Upload history" },
];

const CHART_COLORS = [
  "var(--color-chart-1)",
  "var(--color-chart-2)",
  "var(--color-chart-3)",
  "var(--color-chart-4)",
  "var(--color-chart-5)",
  "var(--color-muted-foreground)",
];

type DrillKind =
  | "active"
  | "attention"
  | "within"
  | "vendor"
  | "termYear"
  | "endMonth"
  | "function"
  | "country"
  | "vendorMonth";
type Drill = { label: string; kind: DrillKind; value?: string } | null;

type Filters = {
  search: string;
  status: string;
  department: string;
  fn: string;
  country: string;
  vendor: string;
  workType: string;
  drill: Drill;
};
const EMPTY_FILTERS: Filters = {
  search: "",
  status: ALL,
  department: ALL,
  fn: ALL,
  country: ALL,
  vendor: ALL,
  workType: ALL,
  drill: null,
};

type SavedReport = {
  id: string;
  name: string;
  filters: Partial<Filters>;
  columns: string[];
  created_by: string;
};

const isActive = (c: Contractor) => ["Active", "Expiring"].includes(getStatus(c));
const orUnspec = (v: string | null) => (v ?? "").trim() || "Unspecified";
const endsWithin = (c: Contractor, n: number) => {
  if (getStatus(c) === "Terminated") return false;
  const d = daysUntil(c.sow_end_date);
  return d !== null && d >= 0 && d <= n;
};
const activeInMonth = (c: Contractor, monthStart: string, monthEnd: string) => {
  const st = c.sow_start_date ?? "";
  const se = c.sow_end_date ?? "";
  const t = c.termination_date ?? "";
  if (st && st >= monthEnd) return false;
  if (se && se < monthStart) return false;
  if (t && t < monthEnd) return false;
  return true;
};
const monthRange = (ym: string) => {
  const [y, m] = ym.split("-").map(Number);
  const start = new Date(Date.UTC(y!, m! - 1, 1)).toISOString().slice(0, 10);
  const end = new Date(Date.UTC(y!, m!, 1)).toISOString().slice(0, 10);
  return [start, end] as const;
};

function testDrill(d: NonNullable<Drill>, c: Contractor): boolean {
  const v = d.value ?? "";
  switch (d.kind) {
    case "active":
      return isActive(c);
    case "attention":
      return needsAttention(c);
    case "within":
      return endsWithin(c, Number(v));
    case "vendor":
      return orUnspec(c.vendor) === v;
    case "termYear":
      return (c.termination_date ?? "").startsWith(v);
    case "endMonth":
      return (c.sow_end_date ?? "").startsWith(v);
    case "function":
      return orUnspec(c.function) === v;
    case "country":
      return orUnspec(c.country) === v;
    case "vendorMonth": {
      const [vendor, ym] = v.split("|");
      const [s, e] = monthRange(ym!);
      return (vendor === "Other" || orUnspec(c.vendor) === vendor) && activeInMonth(c, s, e);
    }
  }
}

const fmtMoney = (n: number, currency = "USD") =>
  new Intl.NumberFormat("en-US", { style: "currency", currency, maximumFractionDigits: 0 }).format(n);

function Kpi({
  label,
  value,
  icon: Icon,
  tone,
  active,
  onClick,
}: {
  label: string;
  value: string | number;
  icon: typeof Users;
  tone: string;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button
      onClick={onClick}
      className={`group rounded-xl border bg-card p-5 text-left transition-all hover:-translate-y-0.5 hover:shadow-md ${
        active ? "border-primary ring-2 ring-primary/30" : "border-border"
      }`}
    >
      <div className="flex items-center justify-between gap-2">
        <span className="text-sm font-medium text-muted-foreground">{label}</span>
        <span className={`shrink-0 rounded-lg p-2 ${tone}`}>
          <Icon className="size-4" />
        </span>
      </div>
      <p className="mt-3 text-3xl font-semibold tabular-nums">{value}</p>
      <p className="mt-1 text-xs text-muted-foreground">Click to view the list</p>
    </button>
  );
}

function Dashboard() {
  const [uploadOpen, setUploadOpen] = useState(false);
  const [sheetOpen, setSheetOpen] = useState(false);
  const [editing, setEditing] = useState<Contractor | null>(null);
  const [f, setF] = useState<Filters>(EMPTY_FILTERS);
  const [renewalWindow, setRenewalWindow] = useState<90 | 180>(180);
  const [cols, setCols] = useState<string[]>(DEFAULT_COLUMNS);
  const [saveOpen, setSaveOpen] = useState(false);
  const [reportName, setReportName] = useState("");

  const patch = (p: Partial<Filters>) => setF((prev) => ({ ...prev, ...p }));
  const setDrill = (drill: Drill) => patch({ drill, status: ALL });

  useEffect(() => {
    try {
      const raw = localStorage.getItem(COLS_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as string[];
        if (Array.isArray(parsed) && parsed.length) setCols(parsed);
      }
    } catch {
      /* ignore */
    }
  }, []);
  const [hiddenWidgets, setHiddenWidgets] = useState<string[]>([]);
  useEffect(() => {
    try {
      const raw = localStorage.getItem(WIDGETS_KEY);
      if (raw) setHiddenWidgets(JSON.parse(raw) as string[]);
    } catch {
      /* ignore */
    }
  }, []);
  const show = (key: string) => !hiddenWidgets.includes(key);
  const updateWidgets = (next: string[]) => {
    setHiddenWidgets(next);
    localStorage.setItem(WIDGETS_KEY, JSON.stringify(next));
  };
  const updateCols = (next: string[]) => {
    setCols(next);
    localStorage.setItem(COLS_KEY, JSON.stringify(next));
  };

  const { data: contractors = [], refetch } = useQuery({
    queryKey: ["contractors"],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("contractors")
        .select("*")
        .order("sow_end_date", { ascending: true, nullsFirst: false });
      if (error) throw error;
      return (data ?? []) as Contractor[];
    },
  });

  const { data: uploads = [], refetch: refetchUploads } = useQuery({
    queryKey: ["upload_history"],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("upload_history")
        .select("*")
        .order("created_at", { ascending: false })
        .limit(50);
      if (error) throw error;
      return data ?? [];
    },
  });

  const { data: me } = useQuery({
    queryKey: ["me"],
    queryFn: async () => (await supabase.auth.getUser()).data.user?.id ?? null,
  });

  const { data: reports = [], refetch: refetchReports } = useQuery({
    queryKey: ["saved_reports"],
    queryFn: async () => {
      const { data, error } = await supabase
        .from("saved_reports")
        .select("id,name,filters,columns,created_by")
        .order("created_at", { ascending: false });
      if (error) throw error;
      return (data ?? []) as unknown as SavedReport[];
    },
  });

  /** Work-type scope applies to cards, charts and table. */
  const scoped = useMemo(
    () => (f.workType === ALL ? contractors : contractors.filter((c) => c.work_type === f.workType)),
    [contractors, f.workType],
  );

  const uniq = (key: keyof Contractor) =>
    Array.from(new Set(scoped.map((c) => (c[key] as string | null) ?? "").filter(Boolean))).sort();

  const filtered = useMemo(() => {
    const q = f.search.trim().toLowerCase();
    return scoped.filter((c) => {
      if (q && !`${c.name} ${c.email ?? ""} ${c.sow_name ?? ""} ${c.role ?? ""} ${c.manager ?? ""}`.toLowerCase().includes(q))
        return false;
      if (f.status !== ALL && getStatus(c) !== f.status) return false;
      if (f.department !== ALL && (c.department ?? "") !== f.department) return false;
      if (f.fn !== ALL && (c.function ?? "") !== f.fn) return false;
      if (f.country !== ALL && (c.country ?? "") !== f.country) return false;
      if (f.vendor !== ALL && (c.vendor ?? "") !== f.vendor) return false;
      if (f.drill && !testDrill(f.drill, c)) return false;
      return true;
    });
  }, [scoped, f]);

  const stats = useMemo(() => {
    let active = 0;
    let in90 = 0;
    let in180 = 0;
    let attention = 0;
    for (const c of scoped) {
      if (isActive(c)) active += 1;
      if (endsWithin(c, 90)) in90 += 1;
      if (endsWithin(c, 180)) in180 += 1;
      if (needsAttention(c)) attention += 1;
    }
    return { total: scoped.length, active, in90, in180, attention };
  }, [scoped]);

  const spend = useMemo(() => {
    const withRate = scoped.filter((c) => isActive(c) && c.monthly_rate != null);
    const curCount = new Map<string, number>();
    for (const c of withRate) {
      const cu = (c.currency ?? "USD").toUpperCase();
      curCount.set(cu, (curCount.get(cu) ?? 0) + 1);
    }
    const currency = [...curCount.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "USD";
    const total = withRate
      .filter((c) => (c.currency ?? "USD").toUpperCase() === currency)
      .reduce((sum, c) => sum + (c.monthly_rate ?? 0), 0);
    return { total, currency, mixed: curCount.size > 1 };
  }, [scoped]);

  const byTerminationYear = useMemo(() => {
    const m = new Map<string, number>();
    for (const c of scoped) {
      if (!c.termination_date) continue;
      const y = c.termination_date.slice(0, 4);
      m.set(y, (m.get(y) ?? 0) + 1);
    }
    return Array.from(m, ([year, count]) => ({ year, count })).sort((a, b) => a.year.localeCompare(b.year));
  }, [scoped]);

  const groupBy = (key: keyof Contractor) => {
    const m = new Map<string, number>();
    for (const c of scoped) {
      const k = orUnspec(c[key] as string | null);
      m.set(k, (m.get(k) ?? 0) + 1);
    }
    return Array.from(m, ([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count);
  };
  const byFunction = useMemo(() => groupBy("function"), [scoped]);
  const byCountry = useMemo(() => groupBy("country"), [scoped]);

  const byVendorSpend = useMemo(() => {
    const m = new Map<string, number>();
    for (const c of scoped) {
      if (!isActive(c)) continue;
      const k = orUnspec(c.vendor);
      m.set(k, (m.get(k) ?? 0) + (c.monthly_rate ?? 0));
    }
    return Array.from(m, ([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 8);
  }, [scoped]);

  const months = useMemo(() => {
    const out: { ym: string; start: string; end: string }[] = [];
    const now = new Date();
    for (let i = 11; i >= 0; i--) {
      const s = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - i, 1));
      const e = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - i + 1, 1));
      out.push({ ym: s.toISOString().slice(0, 7), start: s.toISOString().slice(0, 10), end: e.toISOString().slice(0, 10) });
    }
    return out;
  }, []);

  const headcountTrend = useMemo(
    () => months.map((m) => ({ month: m.ym, count: scoped.filter((c) => activeInMonth(c, m.start, m.end)).length })),
    [scoped, months],
  );

  /** Monthly headcount by vendor (top 5 vendors + Other), last 12 months. */
  const vendorMonthly = useMemo(() => {
    const vendorTotals = new Map<string, number>();
    for (const c of scoped) {
      if (c.work_type === "FTE" && !c.vendor) continue;
      vendorTotals.set(orUnspec(c.vendor), (vendorTotals.get(orUnspec(c.vendor)) ?? 0) + 1);
    }
    const top = [...vendorTotals].sort((a, b) => b[1] - a[1]).slice(0, 5).map(([v]) => v);
    const hasOther = vendorTotals.size > top.length;
    const keys = hasOther ? [...top, "Other"] : top;
    const data = months.map((m) => {
      const row: Record<string, string | number> = { month: m.ym };
      for (const k of keys) row[k] = 0;
      for (const c of scoped) {
        if (c.work_type === "FTE" && !c.vendor) continue;
        if (!activeInMonth(c, m.start, m.end)) continue;
        const v = orUnspec(c.vendor);
        const k = top.includes(v) ? v : "Other";
        row[k] = (row[k] as number) + 1;
      }
      return row;
    });
    return { data, keys, top };
  }, [scoped, months]);

  const renewalQueue = useMemo(
    () =>
      scoped
        .filter((c) => endsWithin(c, renewalWindow))
        .sort((a, b) => (a.sow_end_date ?? "").localeCompare(b.sow_end_date ?? "")),
    [scoped, renewalWindow],
  );

  const expiryByMonth = useMemo(() => {
    const m = new Map<string, number>();
    for (const c of scoped) {
      if (!endsWithin(c, 180)) continue;
      const key = (c.sow_end_date ?? "").slice(0, 7);
      m.set(key, (m.get(key) ?? 0) + 1);
    }
    return Array.from(m, ([month, count]) => ({ month, count })).sort((a, b) => a.month.localeCompare(b.month));
  }, [scoped]);

  const clearFilters = () => setF((p) => ({ ...EMPTY_FILTERS, workType: p.workType }));
  const hasFilters =
    Boolean(f.search) ||
    f.status !== ALL ||
    f.department !== ALL ||
    f.fn !== ALL ||
    f.country !== ALL ||
    f.vendor !== ALL ||
    Boolean(f.drill);

  const openRow = (c: Contractor | null) => {
    setEditing(c);
    setSheetOpen(true);
  };

  const filterSummary = (): [string, string][] => {
    const out: [string, string][] = [];
    if (f.workType !== ALL) out.push(["Work type", f.workType]);
    if (f.search) out.push(["Search", f.search]);
    if (f.status !== ALL) out.push(["Status", f.status]);
    if (f.department !== ALL) out.push(["Department", f.department]);
    if (f.fn !== ALL) out.push(["Function", f.fn]);
    if (f.country !== ALL) out.push(["Country", f.country]);
    if (f.vendor !== ALL) out.push(["Vendor", f.vendor]);
    if (f.drill) out.push(["Drill-down", f.drill.label]);
    return out;
  };

  async function saveReport() {
    const name = reportName.trim();
    if (!name || name.length > 100) return void toast.error("Please enter a name (max 100 characters)");
    const { error } = await supabase.from("saved_reports").insert({ name, filters: f as never, columns: cols });
    if (error) return void toast.error(error.message);
    toast.success(`Saved report "${name}"`);
    setSaveOpen(false);
    setReportName("");
    void refetchReports();
  }

  async function deleteReport(id: string) {
    const { error } = await supabase.from("saved_reports").delete().eq("id", id);
    if (error) return void toast.error(error.message);
    toast.success("Report deleted");
    void refetchReports();
  }

  const applyReport = (r: SavedReport) => {
    setF({ ...EMPTY_FILTERS, ...r.filters });
    if (Array.isArray(r.columns) && r.columns.length) updateCols(r.columns);
    toast.success(`Showing report "${r.name}"`);
  };

  const visibleCols = COLUMNS.filter((c) => cols.includes(c.key));

  const renderCell = (key: string, c: Contractor) => {
    switch (key) {
      case "name":
        return (
          <div className="flex items-center gap-2 font-medium">
            {needsAttention(c) && <AlertTriangle className="size-4 shrink-0 text-destructive" />}
            <span>
              {c.name}
              {c.email && !cols.includes("email") && (
                <span className="block text-xs font-normal text-muted-foreground">{c.email}</span>
              )}
            </span>
          </div>
        );
      case "status": {
        const s = getStatus(c);
        return (
          <Badge className={`${STATUS_STYLES[s]} border-0`} variant="secondary">
            {s}
          </Badge>
        );
      }
      case "work_type":
        return (
          <Badge variant="outline" className="font-normal">
            {c.work_type}
          </Badge>
        );
      case "sow_end_date": {
        const d = daysUntil(c.sow_end_date);
        return (
          <span className="whitespace-nowrap">
            {c.sow_end_date ?? "—"}
            {d !== null && d >= 0 && d <= 180 && (
              <span className={`block text-xs ${d <= 90 ? "text-destructive" : "text-amber-600"}`}>in {d} days</span>
            )}
          </span>
        );
      }
      case "monthly_rate":
        return (
          <span className="whitespace-nowrap tabular-nums">
            {c.monthly_rate != null ? fmtMoney(c.monthly_rate, c.currency ?? "USD") : "—"}
          </span>
        );
      case "renewal_count":
        return <span className="text-muted-foreground">{c.renewal_count > 0 ? `×${c.renewal_count}` : "—"}</span>;
      default: {
        const v = COLUMNS.find((x) => x.key === key)!.value(c);
        return <span className="text-muted-foreground">{v === "" ? "—" : v}</span>;
      }
    }
  };

  return (
    <div className="min-h-screen bg-muted/30">
      <header className="border-b border-border bg-card">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center gap-3 px-4 py-5 sm:px-6">
          <div className="mr-auto min-w-0">
            <h1 className="text-xl font-semibold tracking-tight sm:text-2xl">Workforce SoW Dashboard</h1>
            <p className="text-sm text-muted-foreground">
              Who is expiring, who is still active, and who should no longer be in the system.
            </p>
          </div>
          <div className="flex rounded-lg border border-border p-0.5">
            {[ALL, "Contractor", "FTE"].map((w) => (
              <button
                key={w}
                onClick={() => patch({ workType: w })}
                className={`rounded-md px-3 py-1.5 text-sm transition-colors ${
                  f.workType === w ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:bg-muted"
                }`}
              >
                {w === ALL ? "Everyone" : w === "FTE" ? "FTEs" : "Contractors"}
              </button>
            ))}
          </div>
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <Button variant="outline">
                <LayoutGrid className="size-4" /> Widgets
              </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" className="max-h-96 w-64 overflow-y-auto">
              <DropdownMenuLabel>Show on dashboard</DropdownMenuLabel>
              <DropdownMenuSeparator />
              {WIDGETS.map((w) => (
                <DropdownMenuCheckboxItem
                  key={w.key}
                  checked={show(w.key)}
                  onSelect={(e) => e.preventDefault()}
                  onCheckedChange={(v) =>
                    updateWidgets(v ? hiddenWidgets.filter((k) => k !== w.key) : [...hiddenWidgets, w.key])
                  }
                >
                  {w.label}
                </DropdownMenuCheckboxItem>
              ))}
              <DropdownMenuSeparator />
              <DropdownMenuItem onClick={() => updateWidgets([])}>Show all</DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
          <Button variant="outline" onClick={() => openRow(null)}>
            <Plus className="size-4" /> Add
          </Button>
          <Button onClick={() => setUploadOpen(true)}>
            <Upload className="size-4" /> Upload file
          </Button>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-6 px-4 py-6 sm:px-6">
        <section className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6">
{show("kpi:Total") && (
          <Kpi
            label="Total"
            value={stats.total}
            icon={Users}
            tone="bg-primary/10 text-primary"
            active={!hasFilters}
            onClick={clearFilters}
          />
          )}
{show("kpi:Active") && (
          <Kpi
            label="Active"
            value={stats.active}
            icon={ShieldCheck}
            tone="bg-emerald-500/15 text-emerald-600"
            active={f.drill?.kind === "active"}
            onClick={() => setDrill({ label: "Active", kind: "active" })}
          />
          )}
{show("kpi:Ending within 90 days") && (
          <Kpi
            label="Ending within 90 days"
            value={stats.in90}
            icon={CalendarX}
            tone="bg-destructive/15 text-destructive"
            active={f.drill?.kind === "within" && f.drill.value === "90"}
            onClick={() => setDrill({ label: "Ending within 90 days", kind: "within", value: "90" })}
          />
          )}
{show("kpi:Ending within 180 days") && (
          <Kpi
            label="Ending within 180 days"
            value={stats.in180}
            icon={CalendarClock}
            tone="bg-amber-500/20 text-amber-600"
            active={f.drill?.kind === "within" && f.drill.value === "180"}
            onClick={() => setDrill({ label: "Ending within 180 days", kind: "within", value: "180" })}
          />
          )}
{show("kpi:SoW ended, still in system") && (
          <Kpi
            label="SoW ended, still in system"
            value={stats.attention}
            icon={AlertTriangle}
            tone="bg-destructive/15 text-destructive"
            active={f.drill?.kind === "attention"}
            onClick={() => setDrill({ label: "SoW ended, still in system", kind: "attention" })}
          />
          )}
{show("kpi:Monthly spend (active)") && (
          <Kpi
            label="Monthly spend (active)"
            value={fmtMoney(spend.total, spend.currency) + (spend.mixed ? "+" : "")}
            icon={Wallet}
            tone="bg-sky-500/15 text-sky-600"
            active={false}
            onClick={() => setDrill({ label: "Active", kind: "active" })}
          />
          )}
        </section>

        <section className="grid gap-4 lg:grid-cols-2">
{show("chart:Monthly resource count by vendor") && (
          <Card className="lg:col-span-2">
            <CardHeader>
              <CardTitle className="text-base">Monthly resource count by vendor — last 12 months</CardTitle>
            </CardHeader>
            <CardContent className="h-72">
              {vendorMonthly.keys.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={vendorMonthly.data}>
                    <XAxis dataKey="month" tickLine={false} axisLine={false} fontSize={11} />
                    <YAxis allowDecimals={false} tickLine={false} axisLine={false} fontSize={12} width={28} />
                    <Tooltip cursor={{ fill: "var(--color-muted)" }} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    {vendorMonthly.keys.map((k, i) => (
                      <Bar
                        key={k}
                        dataKey={k}
                        stackId="v"
                        fill={CHART_COLORS[i % CHART_COLORS.length]}
                        className="cursor-pointer"
                        onClick={(d: { month?: string }) =>
                          d.month &&
                          setDrill({
                            label: `${k} — active in ${d.month}`,
                            kind: "vendorMonth",
                            value: `${k}|${d.month}`,
                          })
                        }
                      />
                    ))}
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>
)}

{show("chart:Headcount trend") && (
          <Card>
            <CardHeader>
              <CardTitle className="text-base">Headcount trend — last 12 months</CardTitle>
            </CardHeader>
            <CardContent className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={headcountTrend}>
                  <XAxis dataKey="month" tickLine={false} axisLine={false} fontSize={11} />
                  <YAxis allowDecimals={false} tickLine={false} axisLine={false} fontSize={12} width={28} />
                  <Tooltip />
                  <Line type="monotone" dataKey="count" stroke="var(--color-chart-1)" strokeWidth={2.5} dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>
)}

{show("chart:Monthly spend by vendor") && (
          <Card>
            <CardHeader>
              <CardTitle className="text-base">Monthly spend by vendor</CardTitle>
            </CardHeader>
            <CardContent className="h-64">
              {byVendorSpend.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={byVendorSpend} layout="vertical" margin={{ left: 12 }}>
                    <XAxis
                      type="number"
                      tickLine={false}
                      axisLine={false}
                      fontSize={12}
                      tickFormatter={(v: number) => fmtMoney(v, spend.currency)}
                    />
                    <YAxis type="category" dataKey="name" width={120} tickLine={false} axisLine={false} fontSize={12} />
                    <Tooltip cursor={{ fill: "var(--color-muted)" }} formatter={(v) => fmtMoney(Number(v), spend.currency)} />
                    <Bar
                      dataKey="value"
                      radius={[0, 6, 6, 0]}
                      fill="var(--color-chart-3)"
                      className="cursor-pointer"
                      onClick={(d: { name?: string }) =>
                        d.name && setDrill({ label: `Vendor: ${d.name}`, kind: "vendor", value: d.name })
                      }
                    />
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>
)}

{show("chart:Terminations by year") && (
          <Card>
            <CardHeader>
              <CardTitle className="text-base">Terminations by year</CardTitle>
            </CardHeader>
            <CardContent className="h-64">
              {byTerminationYear.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={byTerminationYear}>
                    <XAxis dataKey="year" tickLine={false} axisLine={false} fontSize={12} />
                    <YAxis allowDecimals={false} tickLine={false} axisLine={false} fontSize={12} width={28} />
                    <Tooltip cursor={{ fill: "var(--color-muted)" }} />
                    <Bar
                      dataKey="count"
                      radius={[6, 6, 0, 0]}
                      fill="var(--color-chart-1)"
                      className="cursor-pointer"
                      onClick={(d: { year?: string }) =>
                        d.year && setDrill({ label: `Terminated in ${d.year}`, kind: "termYear", value: d.year })
                      }
                    />
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>
)}

{show("chart:Upcoming expiries by month") && (
          <Card>
            <CardHeader>
              <CardTitle className="text-base">Upcoming expiries by month</CardTitle>
            </CardHeader>
            <CardContent className="h-64">
              {expiryByMonth.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={expiryByMonth}>
                    <XAxis dataKey="month" tickLine={false} axisLine={false} fontSize={12} />
                    <YAxis allowDecimals={false} tickLine={false} axisLine={false} fontSize={12} width={28} />
                    <Tooltip cursor={{ fill: "var(--color-muted)" }} />
                    <Bar
                      dataKey="count"
                      radius={[6, 6, 0, 0]}
                      fill="var(--color-chart-4)"
                      className="cursor-pointer"
                      onClick={(d: { month?: string }) =>
                        d.month && setDrill({ label: `SoW ending ${d.month}`, kind: "endMonth", value: d.month })
                      }
                    />
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>
)}

{show("chart:Split by function") && (
          <Card>
            <CardHeader>
              <CardTitle className="text-base">Split by function</CardTitle>
            </CardHeader>
            <CardContent className="h-64">
              {byFunction.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={byFunction} layout="vertical" margin={{ left: 12 }}>
                    <XAxis type="number" allowDecimals={false} tickLine={false} axisLine={false} fontSize={12} />
                    <YAxis type="category" dataKey="name" width={110} tickLine={false} axisLine={false} fontSize={12} />
                    <Tooltip cursor={{ fill: "var(--color-muted)" }} />
                    <Bar
                      dataKey="count"
                      radius={[0, 6, 6, 0]}
                      fill="var(--color-chart-2)"
                      className="cursor-pointer"
                      onClick={(d: { name?: string }) =>
                        d.name && setDrill({ label: `Function: ${d.name}`, kind: "function", value: d.name })
                      }
                    />
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>
)}

{show("chart:Split by country") && (
          <Card>
            <CardHeader>
              <CardTitle className="text-base">Split by country</CardTitle>
            </CardHeader>
            <CardContent className="h-64">
              {byCountry.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={byCountry}
                      dataKey="count"
                      nameKey="name"
                      innerRadius={50}
                      outerRadius={85}
                      paddingAngle={2}
                      className="cursor-pointer"
                      onClick={(d: { name?: string }) =>
                        d.name && setDrill({ label: `Country: ${d.name}`, kind: "country", value: d.name })
                      }
                      label={(e: { name?: string }) => e.name ?? ""}
                    >
                      {byCountry.map((_, i) => (
                        <Cell key={i} fill={CHART_COLORS[i % 5]} />
                      ))}
                    </Pie>
                    <Tooltip />
                  </PieChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>
)}
        </section>

{show("chart:Renewal queue") && (
        <Card>
          <CardHeader className="flex flex-row flex-wrap items-center justify-between gap-3 space-y-0">
            <div>
              <CardTitle className="text-base">
                Renewal queue — next {renewalWindow} days{" "}
                <Badge variant="secondary" className="ml-1">
                  {renewalQueue.length}
                </Badge>
              </CardTitle>
              <p className="mt-1 text-sm text-muted-foreground">
                SoWs ending soon, sorted by date. Click a person to update dates, rate or vendor.
              </p>
            </div>
            <div className="flex items-center gap-2">
              <div className="flex rounded-lg border border-border p-0.5">
                {([90, 180] as const).map((n) => (
                  <button
                    key={n}
                    onClick={() => setRenewalWindow(n)}
                    className={`rounded-md px-3 py-1 text-sm ${
                      renewalWindow === n ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:bg-muted"
                    }`}
                  >
                    {n} days
                  </button>
                ))}
              </div>
              <Button
                variant="outline"
                size="sm"
                onClick={() => exportToExcel(renewalQueue, `renewal-queue-${renewalWindow}d.xlsx`)}
                disabled={!renewalQueue.length}
              >
                <Download className="size-4" /> Export queue
              </Button>
            </div>
          </CardHeader>
          <CardContent className="px-0 pb-0">
            {renewalQueue.length ? (
              <div className="max-h-96 overflow-y-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Name</TableHead>
                      <TableHead>Role</TableHead>
                      <TableHead>Vendor</TableHead>
                      <TableHead>Direct manager</TableHead>
                      <TableHead>SoW / Project end</TableHead>
                      <TableHead>Monthly rate</TableHead>
                      <TableHead>Renewals</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {renewalQueue.map((c) => (
                      <TableRow key={c.id} className="cursor-pointer" onClick={() => openRow(c)}>
                        <TableCell>
                          <span className="font-medium">{c.name}</span>
                          {c.email && <span className="block text-xs text-muted-foreground">{c.email}</span>}
                        </TableCell>
                        <TableCell className="text-muted-foreground">{c.role ?? "—"}</TableCell>
                        <TableCell className="text-muted-foreground">{c.vendor ?? "—"}</TableCell>
                        <TableCell className="text-muted-foreground">{c.manager ?? "—"}</TableCell>
                        <TableCell>{renderCell("sow_end_date", c)}</TableCell>
                        <TableCell>{renderCell("monthly_rate", c)}</TableCell>
                        <TableCell>{renderCell("renewal_count", c)}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            ) : (
              <div className="px-4 py-10 text-center text-sm text-muted-foreground">
                Nothing expiring in the next {renewalWindow} days.
              </div>
            )}
          </CardContent>
        </Card>
)}

        <section className="rounded-xl border border-border bg-card">
          <div className="flex flex-wrap items-center gap-3 border-b border-border p-4">
            <div className="relative min-w-52 flex-1">
              <Search className="absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
              <Input
                className="pl-9"
                placeholder="Search name, email, role, manager or SoW"
                value={f.search}
                onChange={(e) => patch({ search: e.target.value })}
              />
            </div>
            <FilterSelect value={f.workType} onChange={(v) => patch({ workType: v })} placeholder="Work type" options={["Contractor", "FTE"]} />
            <FilterSelect
              value={f.status}
              onChange={(v) => patch({ status: v as Status, drill: null })}
              placeholder="Status"
              options={["Active", "Expiring", "Expired", "Terminated"]}
            />
            <FilterSelect value={f.department} onChange={(v) => patch({ department: v })} placeholder="Department" options={uniq("department")} />
            <FilterSelect value={f.fn} onChange={(v) => patch({ fn: v })} placeholder="Function" options={uniq("function")} />
            <FilterSelect value={f.country} onChange={(v) => patch({ country: v })} placeholder="Country" options={uniq("country")} />
            <FilterSelect value={f.vendor} onChange={(v) => patch({ vendor: v })} placeholder="Vendor" options={uniq("vendor")} />
            {hasFilters && (
              <Button variant="ghost" size="sm" onClick={clearFilters}>
                <X className="size-4" /> Clear
              </Button>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2 border-b border-border px-4 py-2">
            {f.drill && (
              <span className="mr-auto flex items-center gap-1 rounded-md bg-primary/10 px-2 py-1 text-sm font-medium">
                Showing: {f.drill.label}
                <button onClick={() => patch({ drill: null })} aria-label="Clear drill-down">
                  <X className="size-3.5" />
                </button>
              </span>
            )}
            <div className="ml-auto flex flex-wrap gap-2">
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="outline" size="sm">
                    <Columns3 className="size-4" /> Columns
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end" className="max-h-96 w-56 overflow-y-auto">
                  <DropdownMenuLabel>Show columns</DropdownMenuLabel>
                  <DropdownMenuSeparator />
                  {COLUMNS.map((c) => (
                    <DropdownMenuCheckboxItem
                      key={c.key}
                      checked={cols.includes(c.key)}
                      disabled={c.key === "name"}
                      onSelect={(e) => e.preventDefault()}
                      onCheckedChange={(v) =>
                        updateCols(
                          v
                            ? COLUMNS.map((x) => x.key).filter((k) => k === c.key || cols.includes(k))
                            : cols.filter((k) => k !== c.key),
                        )
                      }
                    >
                      {c.label}
                    </DropdownMenuCheckboxItem>
                  ))}
                  <DropdownMenuSeparator />
                  <DropdownMenuItem onSelect={() => updateCols(DEFAULT_COLUMNS)}>Reset to default</DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>

              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="outline" size="sm">
                    <FileText className="size-4" /> Reports
                    {reports.length > 0 && <Badge variant="secondary">{reports.length}</Badge>}
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end" className="w-64">
                  <DropdownMenuItem onSelect={() => setSaveOpen(true)}>
                    <Save className="size-4" /> Save current view as report…
                  </DropdownMenuItem>
                  <DropdownMenuSeparator />
                  <DropdownMenuLabel>Saved reports</DropdownMenuLabel>
                  {reports.length ? (
                    reports.map((r) => (
                      <DropdownMenuItem key={r.id} onSelect={() => applyReport(r)} className="justify-between">
                        <span className="truncate">{r.name}</span>
                        {r.created_by === me && (
                          <button
                            aria-label={`Delete ${r.name}`}
                            className="text-muted-foreground hover:text-destructive"
                            onClick={(e) => {
                              e.stopPropagation();
                              e.preventDefault();
                              void deleteReport(r.id);
                            }}
                          >
                            <Trash2 className="size-3.5" />
                          </button>
                        )}
                      </DropdownMenuItem>
                    ))
                  ) : (
                    <div className="px-2 py-1.5 text-sm text-muted-foreground">None yet</div>
                  )}
                </DropdownMenuContent>
              </DropdownMenu>

              <Button
                variant="outline"
                size="sm"
                disabled={!filtered.length}
                onClick={() =>
                  exportReport(filtered, {
                    title: f.drill?.label ?? "Workforce report",
                    filters: filterSummary(),
                    columns: cols,
                    filename: `report-${new Date().toISOString().slice(0, 10)}.xlsx`,
                  })
                }
              >
                <FileText className="size-4" /> Download report
              </Button>
              <Button variant="outline" size="sm" onClick={() => exportToExcel(filtered)} disabled={!filtered.length}>
                <Download className="size-4" /> Export Excel
              </Button>
            </div>
          </div>

          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  {visibleCols.map((c) => (
                    <TableHead key={c.key} className="whitespace-nowrap">
                      {c.label}
                    </TableHead>
                  ))}
                </TableRow>
              </TableHeader>
              <TableBody>
                {filtered.map((c) => (
                  <TableRow key={c.id} className="cursor-pointer" onClick={() => openRow(c)}>
                    {visibleCols.map((col) => (
                      <TableCell key={col.key}>{renderCell(col.key, c)}</TableCell>
                    ))}
                  </TableRow>
                ))}
                {!filtered.length && (
                  <TableRow>
                    <TableCell colSpan={visibleCols.length} className="py-14 text-center text-muted-foreground">
                      {contractors.length
                        ? "No records match these filters."
                        : "No data yet — upload your contractor or FTE Excel file to get started."}
                    </TableCell>
                  </TableRow>
                )}
              </TableBody>
            </Table>
          </div>
          <div className="border-t border-border px-4 py-3 text-sm text-muted-foreground">
            Showing {filtered.length} of {contractors.length} records
          </div>
        </section>
        {show("uploads") && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-base">
                <History className="size-4" /> Upload history
              </CardTitle>
              <p className="text-sm text-muted-foreground">
                Every file uploaded is kept here. Download the original rows or see which conflicting values were kept.
              </p>
            </CardHeader>
            <CardContent className="px-0 pb-0">
              {uploads.length ? (
                <div className="max-h-80 overflow-auto">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>When</TableHead>
                        <TableHead>File</TableHead>
                        <TableHead>Type</TableHead>
                        <TableHead className="text-right">Rows</TableHead>
                        <TableHead className="text-right">Added</TableHead>
                        <TableHead className="text-right">Updated</TableHead>
                        <TableHead className="text-right">Conflicts</TableHead>
                        <TableHead />
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {uploads.map((u) => (
                        <TableRow key={u.id}>
                          <TableCell className="whitespace-nowrap text-muted-foreground">
                            {u.created_at.slice(0, 16).replace("T", " ")}
                          </TableCell>
                          <TableCell className="font-medium">{u.file_name}</TableCell>
                          <TableCell>{u.work_type}</TableCell>
                          <TableCell className="text-right tabular-nums">{u.row_count}</TableCell>
                          <TableCell className="text-right tabular-nums">{u.inserted_count}</TableCell>
                          <TableCell className="text-right tabular-nums">{u.updated_count}</TableCell>
                          <TableCell className="text-right">
                            {u.conflict_count > 0 ? (
                              <Badge className="border-0 bg-amber-500/20 text-amber-700">{u.conflict_count}</Badge>
                            ) : (
                              "—"
                            )}
                          </TableCell>
                          <TableCell className="text-right">
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() =>
                                downloadUpload(
                                  u.file_name,
                                  u.rows as Record<string, unknown>[],
                                  u.overrides as Record<string, unknown>[],
                                )
                              }
                            >
                              <Download className="size-4" />
                            </Button>
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              ) : (
                <p className="px-6 pb-6 text-sm text-muted-foreground">No uploads yet.</p>
              )}
            </CardContent>
          </Card>
        )}
      </main>

      <Dialog open={saveOpen} onOpenChange={setSaveOpen}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle>Save as report</DialogTitle>
            <DialogDescription>
              Saves the current filters and columns. Everyone signed in can open it.
            </DialogDescription>
          </DialogHeader>
          <Input
            autoFocus
            maxLength={100}
            placeholder="e.g. HK contractors ending in 90 days"
            value={reportName}
            onChange={(e) => setReportName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && void saveReport()}
          />
          <div className="text-xs text-muted-foreground">
            {filterSummary().length
              ? filterSummary().map(([k, v]) => `${k}: ${v}`).join(" · ")
              : "No filters — all records"}
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setSaveOpen(false)}>
              Cancel
            </Button>
            <Button onClick={() => void saveReport()}>Save</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <UploadDialog open={uploadOpen} onOpenChange={setUploadOpen} onImported={() => void refetch()} />
      <ContractorSheet contractor={editing} open={sheetOpen} onOpenChange={setSheetOpen} onSaved={() => void refetch()} />
    </div>
  );
}

function FilterSelect({
  value,
  onChange,
  placeholder,
  options,
}: {
  value: string;
  onChange: (v: string) => void;
  placeholder: string;
  options: string[];
}) {
  return (
    <Select value={value} onValueChange={onChange}>
      <SelectTrigger className="w-[150px]">
        <SelectValue placeholder={placeholder}>{value === ALL ? `All ${placeholder.toLowerCase()}` : value}</SelectValue>
      </SelectTrigger>
      <SelectContent>
        <SelectItem value={ALL}>All {placeholder.toLowerCase()}</SelectItem>
        {options.map((o) => (
          <SelectItem key={o} value={o}>
            {o}
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}

function Empty() {
  return <div className="flex h-full items-center justify-center text-sm text-muted-foreground">No data yet</div>;
}

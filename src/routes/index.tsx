import { useMemo, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import {
  Bar,
  BarChart,
  Cell,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { Download, Plus, Upload, Users, ShieldCheck, CalendarClock, AlertTriangle, X, Search } from "lucide-react";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { supabase } from "@/integrations/supabase/client";
import { UploadDialog } from "@/components/contractors/UploadDialog";
import { ContractorSheet } from "@/components/contractors/ContractorSheet";
import {
  EXPIRY_WINDOW_DAYS,
  STATUS_STYLES,
  daysUntil,
  exportToExcel,
  getStatus,
  needsAttention,
  type Contractor,
  type Status,
} from "@/lib/contractors";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Contractor SoW Dashboard | HR Contract Tracker" },
      {
        name: "description",
        content:
          "Upload your contractor list, track SoW expiry and termination dates, spot accounts that should be closed, and export to Excel.",
      },
      { property: "og:title", content: "Contractor SoW Dashboard | HR Contract Tracker" },
      {
        property: "og:description",
        content:
          "Upload your contractor list, track SoW expiry and termination dates, spot accounts that should be closed, and export to Excel.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Dashboard,
});

const ALL = "__all__";
const CHART_COLORS = [
  "var(--color-chart-1)",
  "var(--color-chart-2)",
  "var(--color-chart-3)",
  "var(--color-chart-4)",
  "var(--color-chart-5)",
];

type Drill = { label: string; test: (c: Contractor) => boolean } | null;

function Kpi({
  label,
  value,
  icon: Icon,
  tone,
  active,
  onClick,
}: {
  label: string;
  value: number;
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
      <div className="flex items-center justify-between">
        <span className="text-sm font-medium text-muted-foreground">{label}</span>
        <span className={`rounded-lg p-2 ${tone}`}>
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
  const [search, setSearch] = useState("");
  const [status, setStatus] = useState<string>(ALL);
  const [department, setDepartment] = useState<string>(ALL);
  const [fn, setFn] = useState<string>(ALL);
  const [country, setCountry] = useState<string>(ALL);
  const [drill, setDrill] = useState<Drill>(null);

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

  const uniq = (key: keyof Contractor) =>
    Array.from(new Set(contractors.map((c) => (c[key] as string | null) ?? "").filter(Boolean))).sort();

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return contractors.filter((c) => {
      if (q && !`${c.name} ${c.email ?? ""} ${c.sow_name ?? ""}`.toLowerCase().includes(q)) return false;
      if (status !== ALL && getStatus(c) !== status) return false;
      if (department !== ALL && (c.department ?? "") !== department) return false;
      if (fn !== ALL && (c.function ?? "") !== fn) return false;
      if (country !== ALL && (c.country ?? "") !== country) return false;
      if (drill && !drill.test(c)) return false;
      return true;
    });
  }, [contractors, search, status, department, fn, country, drill]);

  const stats = useMemo(() => {
    let active = 0;
    let expiring = 0;
    let attention = 0;
    for (const c of contractors) {
      const s = getStatus(c);
      if (s === "Active" || s === "Expiring") active += 1;
      if (s === "Expiring") expiring += 1;
      if (needsAttention(c)) attention += 1;
    }
    return { total: contractors.length, active, expiring, attention };
  }, [contractors]);

  const byTerminationYear = useMemo(() => {
    const m = new Map<string, number>();
    for (const c of contractors) {
      if (!c.termination_date) continue;
      const y = c.termination_date.slice(0, 4);
      m.set(y, (m.get(y) ?? 0) + 1);
    }
    return Array.from(m, ([year, count]) => ({ year, count })).sort((a, b) => a.year.localeCompare(b.year));
  }, [contractors]);

  const groupBy = (key: keyof Contractor) => {
    const m = new Map<string, number>();
    for (const c of contractors) {
      const k = ((c[key] as string | null) ?? "Unspecified").trim() || "Unspecified";
      m.set(k, (m.get(k) ?? 0) + 1);
    }
    return Array.from(m, ([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count);
  };

  const byFunction = useMemo(() => groupBy("function"), [contractors]);
  const byCountry = useMemo(() => groupBy("country"), [contractors]);

  const expiryByMonth = useMemo(() => {
    const m = new Map<string, number>();
    for (const c of contractors) {
      const d = daysUntil(c.sow_end_date);
      if (d === null || d < 0 || d > EXPIRY_WINDOW_DAYS) continue;
      const key = (c.sow_end_date ?? "").slice(0, 7);
      m.set(key, (m.get(key) ?? 0) + 1);
    }
    return Array.from(m, ([month, count]) => ({ month, count })).sort((a, b) => a.month.localeCompare(b.month));
  }, [contractors]);

  const clearFilters = () => {
    setSearch("");
    setStatus(ALL);
    setDepartment(ALL);
    setFn(ALL);
    setCountry(ALL);
    setDrill(null);
  };

  const hasFilters =
    search || status !== ALL || department !== ALL || fn !== ALL || country !== ALL || Boolean(drill);

  const openRow = (c: Contractor | null) => {
    setEditing(c);
    setSheetOpen(true);
  };

  const setStatusFilter = (s: Status | typeof ALL) => {
    setDrill(null);
    setStatus(s);
  };

  return (
    <div className="min-h-screen bg-muted/30">
      <header className="border-b border-border bg-card">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center gap-3 px-4 py-5 sm:px-6">
          <div className="mr-auto">
            <h1 className="text-xl font-semibold tracking-tight sm:text-2xl">Contractor SoW Dashboard</h1>
            <p className="text-sm text-muted-foreground">
              Who is expiring, who is still active, and who should no longer be in the system.
            </p>
          </div>
          <Button variant="outline" onClick={() => openRow(null)}>
            <Plus className="size-4" /> Add
          </Button>
          <Button
            variant="outline"
            onClick={() => exportToExcel(filtered.length ? filtered : contractors)}
            disabled={!contractors.length}
          >
            <Download className="size-4" /> Export Excel
          </Button>
          <Button onClick={() => setUploadOpen(true)}>
            <Upload className="size-4" /> Upload file
          </Button>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-6 px-4 py-6 sm:px-6">
        <section className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
          <Kpi
            label="Total contractors"
            value={stats.total}
            icon={Users}
            tone="bg-primary/10 text-primary"
            active={!hasFilters}
            onClick={clearFilters}
          />
          <Kpi
            label="Active contractors"
            value={stats.active}
            icon={ShieldCheck}
            tone="bg-emerald-500/15 text-emerald-600"
            active={drill?.label === "Active contractors"}
            onClick={() => {
              setStatus(ALL);
              setDrill({
                label: "Active contractors",
                test: (c) => ["Active", "Expiring"].includes(getStatus(c)),
              });
            }}
          />
          <Kpi
            label={`Expiring in ${EXPIRY_WINDOW_DAYS} days`}
            value={stats.expiring}
            icon={CalendarClock}
            tone="bg-amber-500/20 text-amber-600"
            active={status === "Expiring"}
            onClick={() => setStatusFilter("Expiring")}
          />
          <Kpi
            label="SoW ended, still in system"
            value={stats.attention}
            icon={AlertTriangle}
            tone="bg-destructive/15 text-destructive"
            active={drill?.label === "SoW ended, still in system"}
            onClick={() => {
              setStatus(ALL);
              setDrill({ label: "SoW ended, still in system", test: needsAttention });
            }}
          />
        </section>

        <section className="grid gap-4 lg:grid-cols-2">
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
                        d.year &&
                        setDrill({
                          label: `Terminated in ${d.year}`,
                          test: (c) => (c.termination_date ?? "").startsWith(d.year!),
                        })
                      }
                    />
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>

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
                        d.month &&
                        setDrill({
                          label: `SoW ending ${d.month}`,
                          test: (c) => (c.sow_end_date ?? "").startsWith(d.month!),
                        })
                      }
                    />
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>

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
                        d.name &&
                        setDrill({
                          label: `Function: ${d.name}`,
                          test: (c) => ((c.function ?? "Unspecified") || "Unspecified") === d.name,
                        })
                      }
                    />
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <Empty />
              )}
            </CardContent>
          </Card>

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
                        d.name &&
                        setDrill({
                          label: `Country: ${d.name}`,
                          test: (c) => ((c.country ?? "Unspecified") || "Unspecified") === d.name,
                        })
                      }
                      label={(e: { name?: string }) => e.name ?? ""}
                    >
                      {byCountry.map((_, i) => (
                        <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />
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
        </section>

        <section className="rounded-xl border border-border bg-card">
          <div className="flex flex-wrap items-center gap-3 border-b border-border p-4">
            <div className="relative min-w-52 flex-1">
              <Search className="absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
              <Input
                className="pl-9"
                placeholder="Search name, email or SoW"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>
            <FilterSelect value={status} onChange={setStatusFilter} placeholder="Status" options={["Active", "Expiring", "Expired", "Terminated"]} />
            <FilterSelect value={department} onChange={setDepartment} placeholder="Department" options={uniq("department")} />
            <FilterSelect value={fn} onChange={setFn} placeholder="Function" options={uniq("function")} />
            <FilterSelect value={country} onChange={setCountry} placeholder="Country" options={uniq("country")} />
            {hasFilters && (
              <Button variant="ghost" size="sm" onClick={clearFilters}>
                <X className="size-4" /> Clear
              </Button>
            )}
          </div>

          {drill && (
            <div className="flex items-center gap-2 border-b border-border bg-primary/5 px-4 py-2 text-sm">
              <span className="font-medium">Showing: {drill.label}</span>
              <Button variant="ghost" size="sm" onClick={() => setDrill(null)}>
                <X className="size-3.5" />
              </Button>
            </div>
          )}

          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Name</TableHead>
                  <TableHead>Department</TableHead>
                  <TableHead>Function</TableHead>
                  <TableHead>Country</TableHead>
                  <TableHead>SoW end</TableHead>
                  <TableHead>Termination</TableHead>
                  <TableHead>ADDIS</TableHead>
                  <TableHead>Status</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {filtered.map((c) => {
                  const s = getStatus(c);
                  const d = daysUntil(c.sow_end_date);
                  return (
                    <TableRow key={c.id} className="cursor-pointer" onClick={() => openRow(c)}>
                      <TableCell>
                        <div className="flex items-center gap-2 font-medium">
                          {needsAttention(c) && <AlertTriangle className="size-4 shrink-0 text-destructive" />}
                          <span>
                            {c.name}
                            {c.email && <span className="block text-xs font-normal text-muted-foreground">{c.email}</span>}
                          </span>
                        </div>
                      </TableCell>
                      <TableCell className="text-muted-foreground">{c.department ?? "—"}</TableCell>
                      <TableCell className="text-muted-foreground">{c.function ?? "—"}</TableCell>
                      <TableCell className="text-muted-foreground">{c.country ?? "—"}</TableCell>
                      <TableCell className="whitespace-nowrap">
                        {c.sow_end_date ?? "—"}
                        {d !== null && d >= 0 && d <= EXPIRY_WINDOW_DAYS && (
                          <span className="block text-xs text-amber-600">in {d} days</span>
                        )}
                      </TableCell>
                      <TableCell className="whitespace-nowrap text-muted-foreground">
                        {c.termination_date ?? "—"}
                      </TableCell>
                      <TableCell className="text-muted-foreground">{c.addis_status ?? "—"}</TableCell>
                      <TableCell>
                        <Badge className={`${STATUS_STYLES[s]} border-0`} variant="secondary">
                          {s}
                        </Badge>
                      </TableCell>
                    </TableRow>
                  );
                })}
                {!filtered.length && (
                  <TableRow>
                    <TableCell colSpan={8} className="py-14 text-center text-muted-foreground">
                      {contractors.length
                        ? "No contractors match these filters."
                        : "No data yet — upload your contractor Excel file to get started."}
                    </TableCell>
                  </TableRow>
                )}
              </TableBody>
            </Table>
          </div>
          <div className="border-t border-border px-4 py-3 text-sm text-muted-foreground">
            Showing {filtered.length} of {contractors.length} contractors
          </div>
        </section>
      </main>

      <UploadDialog open={uploadOpen} onOpenChange={setUploadOpen} onImported={() => void refetch()} />
      <ContractorSheet
        contractor={editing}
        open={sheetOpen}
        onOpenChange={setSheetOpen}
        onSaved={() => void refetch()}
      />
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
  onChange: (v: never) => void;
  placeholder: string;
  options: string[];
}) {
  return (
    <Select value={value} onValueChange={onChange as (v: string) => void}>
      <SelectTrigger className="w-[150px]">
        <SelectValue placeholder={placeholder}>
          {value === ALL ? `All ${placeholder.toLowerCase()}` : value}
        </SelectValue>
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
  return (
    <div className="flex h-full items-center justify-center text-sm text-muted-foreground">No data yet</div>
  );
}

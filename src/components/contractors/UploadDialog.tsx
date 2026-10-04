import { useState } from "react";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { toast } from "sonner";
import { Upload, FileSpreadsheet, AlertTriangle, ArrowLeft } from "lucide-react";
import { supabase } from "@/integrations/supabase/client";
import {
  FIELDS,
  guessMapping,
  mapRows,
  readSheet,
  type Contractor,
  type ContractorInput,
  type MappedRow,
  type ParsedSheet,
} from "@/lib/contractors";

const NONE = "__none__";
const SKIP_COMPARE = new Set<keyof ContractorInput>(["renewal_count"]);

type FieldDiff = { key: keyof ContractorInput; label: string; oldVal: unknown; newVal: unknown };
type Conflict = { existing: Contractor; incoming: ContractorInput; diffs: FieldDiff[] };

const isEmpty = (v: unknown) => v === null || v === undefined || v === "";
const same = (a: unknown, b: unknown) =>
  typeof a === "number" || typeof b === "number"
    ? Number(a) === Number(b)
    : String(a ?? "").trim().toLowerCase() === String(b ?? "").trim().toLowerCase();

function diffRecord(existing: Contractor, incoming: ContractorInput): FieldDiff[] {
  const out: FieldDiff[] = [];
  for (const f of FIELDS) {
    if (SKIP_COMPARE.has(f.key)) continue;
    const nv = incoming[f.key];
    const ov = existing[f.key];
    if (isEmpty(nv) || isEmpty(ov)) continue;
    if (!same(ov, nv)) out.push({ key: f.key, label: f.label, oldVal: ov, newVal: nv });
  }
  return out;
}

export function UploadDialog({
  open,
  onOpenChange,
  onImported,
}: {
  open: boolean;
  onOpenChange: (v: boolean) => void;
  onImported: () => void;
}) {
  const [sheet, setSheet] = useState<ParsedSheet | null>(null);
  const [fileName, setFileName] = useState("");
  const [mapping, setMapping] = useState<Record<string, string>>({});
  const [mode, setMode] = useState<"merge" | "replace">("merge");
  const [workType, setWorkType] = useState<"Contractor" | "FTE">("Contractor");
  const [busy, setBusy] = useState(false);
  const [conflicts, setConflicts] = useState<Conflict[] | null>(null);
  const [existingRows, setExistingRows] = useState<Contractor[]>([]);
  /** choice[existingId][field] = "old" | "new" */
  const [choice, setChoice] = useState<Record<string, Record<string, "old" | "new">>>({});

  const reset = () => {
    setSheet(null);
    setFileName("");
    setMapping({});
    setBusy(false);
    setConflicts(null);
    setChoice({});
  };

  async function handleFile(file: File) {
    try {
      const parsed = await readSheet(file);
      if (!parsed.headers.length) {
        toast.error("No columns found in that file");
        return;
      }
      setSheet(parsed);
      setFileName(file.name);
      setMapping(guessMapping(parsed.headers));
      if (/fte|employee|staff/i.test(file.name)) setWorkType("FTE");
    } catch {
      toast.error("Could not read that file. Please use .xlsx or .csv");
    }
  }

  const mapped: MappedRow[] = sheet ? mapRows(sheet.rows, mapping, workType) : [];
  const valid = mapped.filter((r) => r.record.name);
  const problems = mapped.filter((r) => r.issues.length);

  /** Step 1 (merge): look for differences against current records. */
  async function checkAndImport() {
    if (!valid.length) {
      toast.error("Nothing to import");
      return;
    }
    if (mode === "replace") return void writeReplace();
    setBusy(true);
    const { data, error } = await supabase.from("contractors").select("*");
    setBusy(false);
    if (error) return void toast.error(error.message);
    const rows = (data ?? []) as Contractor[];
    setExistingRows(rows);
    const byEmail = new Map(rows.filter((e) => e.email).map((e) => [e.email!.toLowerCase(), e]));
    const found: Conflict[] = [];
    for (const r of valid) {
      const ex = r.record.email ? byEmail.get(r.record.email.toLowerCase()) : undefined;
      if (!ex) continue;
      const diffs = diffRecord(ex, r.record);
      if (diffs.length) found.push({ existing: ex, incoming: r.record, diffs });
    }
    if (found.length) {
      const init: typeof choice = {};
      for (const c of found) init[c.existing.id] = Object.fromEntries(c.diffs.map((d) => [d.key, "new"]));
      setChoice(init);
      setConflicts(found);
      return;
    }
    await writeMerge(rows, {});
  }

  async function writeReplace() {
    setBusy(true);
    try {
      const { error } = await supabase.from("contractors").delete().neq("id", "00000000-0000-0000-0000-000000000000");
      if (error) throw error;
      const { error: insErr } = await supabase.from("contractors").insert(valid.map((r) => r.record));
      if (insErr) throw insErr;
      done();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Import failed");
    } finally {
      setBusy(false);
    }
  }

  async function writeMerge(rows: Contractor[], picks: typeof choice) {
    setBusy(true);
    try {
      const byEmail = new Map(rows.filter((e) => e.email).map((e) => [e.email!.toLowerCase(), e]));
      const toInsert: ContractorInput[] = [];
      for (const r of valid) {
        const key = r.record.email?.toLowerCase();
        const found = key ? byEmail.get(key) : undefined;
        if (!found) {
          toInsert.push(r.record);
          continue;
        }
        // Only fields with a value in the file overwrite; per-field "keep current" respected.
        const update: Record<string, unknown> = {};
        for (const f of FIELDS) {
          if (f.key === "renewal_count") continue;
          const nv = r.record[f.key];
          if (isEmpty(nv)) continue;
          if (picks[found.id]?.[f.key] === "old") continue;
          update[f.key] = nv;
        }
        const finalEnd = (update.sow_end_date as string | undefined) ?? found.sow_end_date;
        const renewed =
          found.renewal_count + (finalEnd && found.sow_end_date && finalEnd > found.sow_end_date ? 1 : 0);
        update.renewal_count = renewed;
        const { error: upErr } = await supabase.from("contractors").update(update).eq("id", found.id);
        if (upErr) throw upErr;
      }
      if (toInsert.length) {
        const { error: insErr } = await supabase.from("contractors").insert(toInsert);
        if (insErr) throw insErr;
      }
      done();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Import failed");
    } finally {
      setBusy(false);
    }
  }

  function done() {
    toast.success(`Imported ${valid.length} record${valid.length === 1 ? "" : "s"}`);
    onImported();
    onOpenChange(false);
    reset();
  }

  const setAll = (v: "old" | "new") =>
    setChoice((prev) =>
      Object.fromEntries(Object.entries(prev).map(([id, m]) => [id, Object.fromEntries(Object.keys(m).map((k) => [k, v]))])),
    );

  return (
    <Dialog
      open={open}
      onOpenChange={(v) => {
        onOpenChange(v);
        if (!v) reset();
      }}
    >
      <DialogContent className="max-h-[90vh] max-w-3xl overflow-y-auto">
        <DialogHeader>
          <DialogTitle>{conflicts ? "Resolve conflicts" : "Upload contractor or FTE list"}</DialogTitle>
          <DialogDescription>
            {conflicts
              ? `${conflicts.length} existing ${conflicts.length === 1 ? "person has" : "people have"} different details in this file. Choose which value to keep.`
              : "Excel (.xlsx) or CSV. Check the column matching before importing."}
          </DialogDescription>
        </DialogHeader>

        {conflicts ? (
          <div className="space-y-4">
            <div className="flex flex-wrap gap-2">
              <Button variant="outline" size="sm" onClick={() => setAll("old")}>
                Keep all current
              </Button>
              <Button variant="outline" size="sm" onClick={() => setAll("new")}>
                Use all from file
              </Button>
            </div>
            <div className="space-y-3">
              {conflicts.map((c) => (
                <div key={c.existing.id} className="rounded-lg border border-border">
                  <div className="border-b border-border bg-muted/40 px-3 py-2 text-sm">
                    <span className="font-medium">{c.existing.name}</span>{" "}
                    <span className="text-muted-foreground">{c.existing.email}</span>
                  </div>
                  <div className="divide-y divide-border">
                    {c.diffs.map((d) => {
                      const pick = choice[c.existing.id]?.[d.key] ?? "new";
                      const opt = (v: "old" | "new", label: string, val: unknown) => (
                        <button
                          type="button"
                          onClick={() =>
                            setChoice((p) => ({ ...p, [c.existing.id]: { ...p[c.existing.id], [d.key]: v } }))
                          }
                          className={`flex-1 rounded-md border px-2 py-1.5 text-left text-sm transition-colors ${
                            pick === v ? "border-primary bg-primary/10" : "border-border hover:bg-muted"
                          }`}
                        >
                          <span className="block text-[11px] uppercase tracking-wide text-muted-foreground">{label}</span>
                          {String(val)}
                        </button>
                      );
                      return (
                        <div key={d.key} className="grid gap-2 px-3 py-2 sm:grid-cols-[140px_1fr]">
                          <span className="pt-1.5 text-xs font-medium text-muted-foreground">{d.label}</span>
                          <div className="flex gap-2">
                            {opt("old", "Current", d.oldVal)}
                            {opt("new", "From file", d.newVal)}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))}
            </div>
            <div className="flex justify-between gap-2">
              <Button variant="ghost" onClick={() => setConflicts(null)}>
                <ArrowLeft className="size-4" /> Back
              </Button>
              <Button onClick={() => writeMerge(existingRows, choice)} disabled={busy}>
                {busy ? "Importing…" : `Import ${valid.length} records`}
              </Button>
            </div>
          </div>
        ) : !sheet ? (
          <label
            className="flex cursor-pointer flex-col items-center justify-center gap-3 rounded-xl border-2 border-dashed border-border bg-muted/40 px-6 py-14 text-center transition-colors hover:border-primary/60 hover:bg-muted"
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault();
              const f = e.dataTransfer.files?.[0];
              if (f) void handleFile(f);
            }}
          >
            <Upload className="size-8 text-muted-foreground" />
            <div>
              <p className="font-medium">Drop your file here or click to browse</p>
              <p className="text-sm text-muted-foreground">.xlsx, .xls or .csv</p>
            </div>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) void handleFile(f);
              }}
            />
          </label>
        ) : (
          <div className="space-y-5">
            <div className="flex items-center gap-2 rounded-lg border border-border bg-muted/40 px-3 py-2 text-sm">
              <FileSpreadsheet className="size-4 text-primary" />
              <span className="font-medium">{fileName}</span>
              <span className="text-muted-foreground">· {sheet.rows.length} rows</span>
              <Button variant="ghost" size="sm" className="ml-auto" onClick={reset}>
                Change file
              </Button>
            </div>

            <div>
              <h3 className="mb-2 text-sm font-semibold">This file is a list of</h3>
              <RadioGroup
                value={workType}
                onValueChange={(v) => setWorkType(v as "Contractor" | "FTE")}
                className="flex gap-4"
              >
                {(["Contractor", "FTE"] as const).map((w) => (
                  <label key={w} className="flex cursor-pointer items-center gap-2 text-sm">
                    <RadioGroupItem value={w} /> {w === "Contractor" ? "Contractors" : "FTEs"}
                  </label>
                ))}
              </RadioGroup>
              <p className="mt-1 text-xs text-muted-foreground">
                Used for rows without a Work Type column value.
              </p>
            </div>

            <div>
              <h3 className="mb-2 text-sm font-semibold">Match your columns</h3>
              <div className="grid gap-3 sm:grid-cols-2">
                {FIELDS.map((f) => (
                  <div key={f.key} className="space-y-1">
                    <Label className="text-xs text-muted-foreground">{f.label}</Label>
                    <Select
                      value={mapping[f.key] ?? NONE}
                      onValueChange={(v) =>
                        setMapping((m) => {
                          const next = { ...m };
                          if (v === NONE) delete next[f.key];
                          else next[f.key] = v;
                          return next;
                        })
                      }
                    >
                      <SelectTrigger>
                        <SelectValue placeholder="Not imported" />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value={NONE}>Not imported</SelectItem>
                        {sheet.headers.map((h) => (
                          <SelectItem key={h} value={h}>
                            {h}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                ))}
              </div>
            </div>

            {problems.length > 0 && (
              <div className="rounded-lg border border-amber-500/40 bg-amber-500/10 p-3 text-sm">
                <p className="flex items-center gap-2 font-medium text-amber-700 dark:text-amber-300">
                  <AlertTriangle className="size-4" />
                  {problems.length} row{problems.length === 1 ? "" : "s"} need attention
                </p>
                <ul className="mt-2 max-h-28 list-disc space-y-0.5 overflow-y-auto pl-5 text-muted-foreground">
                  {problems.slice(0, 8).map((p, i) => (
                    <li key={i}>
                      {p.record.name || "(no name)"} — {p.issues.join(", ")}
                    </li>
                  ))}
                </ul>
                <p className="mt-2 text-xs text-muted-foreground">Rows without a name are skipped.</p>
              </div>
            )}

            <div>
              <h3 className="mb-2 text-sm font-semibold">How should this be imported?</h3>
              <RadioGroup value={mode} onValueChange={(v) => setMode(v as "merge" | "replace")} className="gap-2">
                <label className="flex cursor-pointer items-start gap-3 rounded-lg border border-border p-3">
                  <RadioGroupItem value="merge" className="mt-0.5" />
                  <span className="text-sm">
                    <span className="font-medium">Merge by email</span>
                    <span className="block text-muted-foreground">
                      Update matching people, add the new ones. You'll be asked about any conflicting values.
                    </span>
                  </span>
                </label>
                <label className="flex cursor-pointer items-start gap-3 rounded-lg border border-border p-3">
                  <RadioGroupItem value="replace" className="mt-0.5" />
                  <span className="text-sm">
                    <span className="font-medium">Replace everything</span>
                    <span className="block text-muted-foreground">Delete all current records first.</span>
                  </span>
                </label>
              </RadioGroup>
            </div>

            <div className="flex justify-end gap-2">
              <Button variant="outline" onClick={() => onOpenChange(false)}>
                Cancel
              </Button>
              <Button onClick={checkAndImport} disabled={busy || !valid.length}>
                {busy ? "Checking…" : `Import ${valid.length} records`}
              </Button>
            </div>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}

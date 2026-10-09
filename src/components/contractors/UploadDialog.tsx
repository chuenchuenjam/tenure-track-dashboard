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

const personKey = (r: ContractorInput) =>
  r.email?.trim() ? r.email.trim().toLowerCase() : `name:${r.name.trim().toLowerCase()}`;
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
  const problems = mapped.filter((r) => r.issues.length);
  /** Collapse duplicate people inside the same file (same email, or same name when no email). */
  const { valid, dupes } = (() => {
    const byKey = new Map<string, ContractorInput>();
    const dupes: string[] = [];
    for (const r of mapped) {
      if (!r.record.name) continue;
      const k = personKey(r.record);
      const prev = byKey.get(k);
      if (!prev) {
        byKey.set(k, { ...r.record });
        continue;
      }
      const merged = { ...prev } as Record<string, unknown>;
      let clash = false;
      for (const f of FIELDS) {
        const nv = r.record[f.key];
        if (isEmpty(nv)) continue;
        if (!isEmpty(merged[f.key]) && !same(merged[f.key], nv)) clash = true;
        merged[f.key] = nv;
      }
      if (clash) dupes.push(r.record.name);
      byKey.set(k, merged as ContractorInput);
    }
    return { valid: [...byKey.values()].map((record) => ({ record })), dupes };
  })();

  function findExisting(rows: Contractor[], rec: ContractorInput) {
    const email = rec.email?.trim().toLowerCase();
    if (email) {
      const hit = rows.find((e) => e.email?.trim().toLowerCase() === email);
      if (hit) return hit;
    }
    const name = rec.name.trim().toLowerCase();
    // Fall back to name only when one side has no email, to avoid merging two different people.
    return rows.find(
      (e) => e.name.trim().toLowerCase() === name && (!email || !e.email),
    );
  }

  /** Step 1: look for differences against current records. */
  async function checkAndImport() {
    if (!valid.length) {
      toast.error("Nothing to import");
      return;
    }
    setBusy(true);
    const { data, error } = await supabase.from("contractors").select("*");
    setBusy(false);
    if (error) return void toast.error(error.message);
    const rows = (data ?? []) as Contractor[];
    setExistingRows(rows);
    const found: Conflict[] = [];
    for (const r of valid) {
      const ex = findExisting(rows, r.record);
      if (!ex) continue;
      const diffs = diffRecord(ex, r.record);
      if (diffs.length) found.push({ existing: ex, incoming: r.record, diffs });
    }
    if (found.length) {
      const init: typeof choice = {};
      for (const c of found) init[c.existing.id] = Object.fromEntries(c.diffs.map((d) => [d.key, "old"]));
      setChoice(init);
      setConflicts(found);
      return;
    }
    await writeMerge(rows, {}, []);
  }

  async function writeMerge(rows: Contractor[], picks: typeof choice, found: Conflict[]) {
    setBusy(true);
    try {
      const toInsert: ContractorInput[] = [];
      const matched = new Set<string>();
      let updated = 0;
      for (const r of valid) {
        const found = findExisting(rows, r.record);
        if (!found || matched.has(found.id)) {
          toInsert.push(r.record);
          continue;
        }
        matched.add(found.id);
        // Cumulative: only non-empty file values are applied; empty cells never wipe saved data.
        const update: Partial<ContractorInput> = {};
        for (const f of FIELDS) {
          if (f.key === "renewal_count") continue;
          const nv = r.record[f.key];
          if (isEmpty(nv)) continue;
          if (picks[found.id]?.[f.key] === "old") continue;
          if (same(found[f.key], nv)) continue;
          (update as Record<string, unknown>)[f.key] = nv;
        }
        const finalEnd = update.sow_end_date ?? found.sow_end_date;
        if (finalEnd && found.sow_end_date && finalEnd > found.sow_end_date) {
          update.renewal_count = found.renewal_count + 1;
        }
        if (!Object.keys(update).length) continue;
        const { error: upErr } = await supabase.from("contractors").update(update).eq("id", found.id);
        if (upErr) throw upErr;
        updated++;
      }
      if (toInsert.length) {
        const { error: insErr } = await supabase.from("contractors").insert(toInsert);
        if (insErr) throw insErr;
      }
      const overrides = found.flatMap((c) =>
        c.diffs.map((d) => ({
          name: c.existing.name,
          email: c.existing.email,
          field: d.label,
          current: d.oldVal,
          file: d.newVal,
          kept: picks[c.existing.id]?.[d.key] === "old" ? "current" : "file",
        })),
      );
      const { error: logErr } = await supabase.from("upload_history").insert({
        file_name: fileName,
        work_type: workType,
        row_count: mapped.length,
        inserted_count: toInsert.length,
        updated_count: updated,
        conflict_count: found.length,
        overrides: overrides as never,
        rows: (sheet?.rows ?? []) as never,
      });
      if (logErr) console.warn("Upload history not saved", logErr);
      toast.success(`Added ${toInsert.length}, updated ${updated}${found.length ? `, ${found.length} conflicts resolved` : ""}`);
      onImported();
      onOpenChange(false);
      reset();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Import failed");
    } finally {
      setBusy(false);
    }
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
              ? `${conflicts.length} existing ${conflicts.length === 1 ? "person has" : "people have"} different details in this file. Nothing changes until you pick — current values are kept by default.`
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
                <div key={c.existing.id} className="rounded-lg border border-amber-500/50">
                  <div className="flex items-center gap-2 border-b border-amber-500/40 bg-amber-500/10 px-3 py-2 text-sm">
                    <AlertTriangle className="size-4 text-amber-600" />
                    <span className="font-medium">{c.existing.name}</span>
                    <span className="text-muted-foreground">{c.existing.email}</span>
                    <span className="ml-auto text-xs text-muted-foreground">{c.diffs.length} field{c.diffs.length === 1 ? "" : "s"} differ</span>
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
              <Button onClick={() => writeMerge(existingRows, choice, conflicts)} disabled={busy}>
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

            {dupes.length > 0 && (
              <div className="rounded-lg border border-amber-500/40 bg-amber-500/10 p-3 text-sm">
                <p className="flex items-center gap-2 font-medium text-amber-700 dark:text-amber-300">
                  <AlertTriangle className="size-4" />
                  {dupes.length} person{dupes.length === 1 ? " appears" : "s appear"} more than once with different details
                </p>
                <p className="mt-1 text-muted-foreground">
                  {dupes.slice(0, 6).join(", ")}
                  {dupes.length > 6 ? "…" : ""} — the last row in the file is used.
                </p>
              </div>
            )}

            <p className="rounded-lg border border-border bg-muted/40 p-3 text-sm text-muted-foreground">
              Records are kept and added to over time: new people are added, matching people (by email, or by name
              when there is no email) are updated, and empty cells never erase saved data. If saved details differ
              from the file, you'll choose which value to keep. Every upload is saved in Upload history.
            </p>

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

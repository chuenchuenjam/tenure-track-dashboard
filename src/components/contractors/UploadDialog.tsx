import { useState } from "react";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { toast } from "sonner";
import { Upload, FileSpreadsheet, AlertTriangle } from "lucide-react";
import { supabase } from "@/integrations/supabase/client";
import {
  FIELDS,
  guessMapping,
  mapRows,
  readSheet,
  type MappedRow,
  type ParsedSheet,
} from "@/lib/contractors";

const NONE = "__none__";

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
  const [busy, setBusy] = useState(false);

  const reset = () => {
    setSheet(null);
    setFileName("");
    setMapping({});
    setBusy(false);
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
    } catch {
      toast.error("Could not read that file. Please use .xlsx or .csv");
    }
  }

  const mapped: MappedRow[] = sheet ? mapRows(sheet.rows, mapping) : [];
  const valid = mapped.filter((r) => r.record.name);
  const problems = mapped.filter((r) => r.issues.length);

  async function doImport() {
    if (!valid.length) {
      toast.error("Nothing to import");
      return;
    }
    setBusy(true);
    try {
      if (mode === "replace") {
        const { error } = await supabase.from("contractors").delete().neq("id", "00000000-0000-0000-0000-000000000000");
        if (error) throw error;
        const { error: insErr } = await supabase.from("contractors").insert(valid.map((r) => r.record));
        if (insErr) throw insErr;
      } else {
        const { data: existing, error } = await supabase
          .from("contractors")
          .select("id,email,sow_end_date,renewal_count");
        if (error) throw error;
        const byEmail = new Map(
          (existing ?? []).filter((e) => e.email).map((e) => [String(e.email).toLowerCase(), e]),
        );
        const toInsert: (typeof valid)[number]["record"][] = [];
        for (const r of valid) {
          const key = r.record.email?.toLowerCase();
          const found = key ? byEmail.get(key) : undefined;
          if (found) {
            // A later SoW end date than before counts as a renewal.
            const renewed =
              found.renewal_count +
              (r.record.sow_end_date && found.sow_end_date && r.record.sow_end_date > found.sow_end_date
                ? 1
                : 0);
            const { error: upErr } = await supabase
              .from("contractors")
              .update({ ...r.record, renewal_count: renewed })
              .eq("id", found.id);
            if (upErr) throw upErr;
          } else {
            toInsert.push(r.record);
          }
        }
        if (toInsert.length) {
          const { error: insErr } = await supabase.from("contractors").insert(toInsert);
          if (insErr) throw insErr;
        }
      }
      toast.success(`Imported ${valid.length} contractor${valid.length === 1 ? "" : "s"}`);
      onImported();
      onOpenChange(false);
      reset();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Import failed");
    } finally {
      setBusy(false);
    }
  }

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
          <DialogTitle>Upload contractor list</DialogTitle>
          <DialogDescription>Excel (.xlsx) or CSV. Check the column matching before importing.</DialogDescription>
        </DialogHeader>

        {!sheet ? (
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
                    <span className="block text-muted-foreground">Update matching people, add the new ones.</span>
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
              <Button onClick={doImport} disabled={busy || !valid.length}>
                {busy ? "Importing…" : `Import ${valid.length} contractors`}
              </Button>
            </div>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}

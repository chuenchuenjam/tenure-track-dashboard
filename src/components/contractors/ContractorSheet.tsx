import { useEffect, useState } from "react";
import { Sheet, SheetContent, SheetHeader, SheetTitle, SheetDescription } from "@/components/ui/sheet";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { toast } from "sonner";
import { Trash2 } from "lucide-react";
import { supabase } from "@/integrations/supabase/client";
import { FIELDS, type Contractor, type ContractorInput } from "@/lib/contractors";

const empty: ContractorInput = {
  name: "",
  email: null,
  department: null,
  function: null,
  country: null,
  vendor: null,
  monthly_rate: null,
  currency: null,
  renewal_count: 0,
  sow_name: null,
  sow_start_date: null,
  sow_end_date: null,
  termination_date: null,
  addis_status: null,
  manager: null,
  notes: null,
};

export function ContractorSheet({
  contractor,
  open,
  onOpenChange,
  onSaved,
}: {
  contractor: Contractor | null;
  open: boolean;
  onOpenChange: (v: boolean) => void;
  onSaved: () => void;
}) {
  const [form, setForm] = useState<ContractorInput>(empty);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!open) return;
    if (contractor) {
      const { id: _id, created_at: _c, updated_at: _u, ...rest } = contractor;
      setForm(rest);
    } else {
      setForm(empty);
    }
  }, [contractor, open]);

  const set = (key: keyof ContractorInput, value: string) =>
    setForm((f) => ({
      ...f,
      [key]: value === "" ? (key === "renewal_count" ? 0 : null) : value,
    }));

  async function save() {
    if (!form.name?.trim()) {
      toast.error("Name is required");
      return;
    }
    setBusy(true);
    try {
      if (contractor) {
        const { error } = await supabase.from("contractors").update(form).eq("id", contractor.id);
        if (error) throw error;
        toast.success("Contractor updated");
      } else {
        const { error } = await supabase.from("contractors").insert(form);
        if (error) throw error;
        toast.success("Contractor added");
      }
      onSaved();
      onOpenChange(false);
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Could not save");
    } finally {
      setBusy(false);
    }
  }

  async function remove() {
    if (!contractor) return;
    setBusy(true);
    const { error } = await supabase.from("contractors").delete().eq("id", contractor.id);
    setBusy(false);
    if (error) {
      toast.error(error.message);
      return;
    }
    toast.success("Contractor deleted");
    onSaved();
    onOpenChange(false);
  }

  return (
    <Sheet open={open} onOpenChange={onOpenChange}>
      <SheetContent className="w-full overflow-y-auto sm:max-w-lg">
        <SheetHeader>
          <SheetTitle>{contractor ? contractor.name : "Add contractor"}</SheetTitle>
          <SheetDescription>Change any field — the dashboard updates straight away.</SheetDescription>
        </SheetHeader>

        <div className="space-y-4 px-4 pb-6">
          {FIELDS.filter((f) => f.key !== "notes").map((f) => (
            <div key={f.key} className="space-y-1.5">
              <Label htmlFor={f.key}>{f.label}</Label>
              <Input
                id={f.key}
                type={f.type === "date" ? "date" : f.type === "number" ? "number" : "text"}
                step={f.key === "monthly_rate" ? "0.01" : undefined}
                min={f.type === "number" ? 0 : undefined}
                value={((form[f.key] as string | number | null) ?? "") as string | number}
                onChange={(e) => set(f.key, e.target.value)}
              />
            </div>
          ))}
          <div className="space-y-1.5">
            <Label htmlFor="notes">Notes</Label>
            <Textarea id="notes" rows={3} value={form.notes ?? ""} onChange={(e) => set("notes", e.target.value)} />
          </div>

          <div className="flex gap-2 pt-2">
            <Button onClick={save} disabled={busy} className="flex-1">
              {busy ? "Saving…" : "Save"}
            </Button>
            {contractor && (
              <Button variant="outline" onClick={remove} disabled={busy} className="text-destructive">
                <Trash2 className="size-4" />
              </Button>
            )}
          </div>
        </div>
      </SheetContent>
    </Sheet>
  );
}

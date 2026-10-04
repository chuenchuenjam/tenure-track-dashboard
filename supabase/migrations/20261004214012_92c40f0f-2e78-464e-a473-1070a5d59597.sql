ALTER TABLE public.contractors ADD COLUMN role text, ADD COLUMN work_type text NOT NULL DEFAULT 'Contractor', ADD COLUMN contract_type text;
CREATE TABLE public.saved_reports (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  filters jsonb NOT NULL DEFAULT '{}'::jsonb,
  columns jsonb NOT NULL DEFAULT '[]'::jsonb,
  created_by uuid NOT NULL DEFAULT auth.uid(),
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.saved_reports TO authenticated;
GRANT ALL ON public.saved_reports TO service_role;
ALTER TABLE public.saved_reports ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Signed-in users can view reports" ON public.saved_reports FOR SELECT TO authenticated USING (true);
CREATE POLICY "Users create own reports" ON public.saved_reports FOR INSERT TO authenticated WITH CHECK (created_by = auth.uid());
CREATE POLICY "Users update own reports" ON public.saved_reports FOR UPDATE TO authenticated USING (created_by = auth.uid()) WITH CHECK (created_by = auth.uid());
CREATE POLICY "Users delete own reports" ON public.saved_reports FOR DELETE TO authenticated USING (created_by = auth.uid());
CREATE TRIGGER update_saved_reports_updated_at BEFORE UPDATE ON public.saved_reports FOR EACH ROW EXECUTE FUNCTION public.update_updated_at_column();
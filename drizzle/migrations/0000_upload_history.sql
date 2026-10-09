CREATE TABLE public.upload_history (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  file_name text NOT NULL,
  work_type text NOT NULL DEFAULT 'Contractor',
  row_count integer NOT NULL DEFAULT 0,
  inserted_count integer NOT NULL DEFAULT 0,
  updated_count integer NOT NULL DEFAULT 0,
  conflict_count integer NOT NULL DEFAULT 0,
  overrides jsonb NOT NULL DEFAULT '[]'::jsonb,
  rows jsonb NOT NULL DEFAULT '[]'::jsonb,
  created_by uuid NOT NULL DEFAULT auth.uid(),
  created_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT ON public.upload_history TO authenticated;
GRANT ALL ON public.upload_history TO service_role;
ALTER TABLE public.upload_history ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Signed-in users can view upload history" ON public.upload_history FOR SELECT TO authenticated USING (true);
CREATE POLICY "Users log own uploads" ON public.upload_history FOR INSERT TO authenticated WITH CHECK (created_by = auth.uid());
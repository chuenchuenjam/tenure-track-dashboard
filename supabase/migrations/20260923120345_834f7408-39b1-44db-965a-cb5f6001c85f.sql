DROP POLICY IF EXISTS "Anyone can view contractors" ON public.contractors;
DROP POLICY IF EXISTS "Anyone can insert contractors" ON public.contractors;
DROP POLICY IF EXISTS "Anyone can update contractors" ON public.contractors;
DROP POLICY IF EXISTS "Anyone can delete contractors" ON public.contractors;

REVOKE ALL ON public.contractors FROM anon;
GRANT SELECT, INSERT, UPDATE, DELETE ON public.contractors TO authenticated;
GRANT ALL ON public.contractors TO service_role;

ALTER TABLE public.contractors ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Signed-in users can view contractors" ON public.contractors FOR SELECT TO authenticated USING (true);
CREATE POLICY "Signed-in users can insert contractors" ON public.contractors FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "Signed-in users can update contractors" ON public.contractors FOR UPDATE TO authenticated USING (true) WITH CHECK (true);
CREATE POLICY "Signed-in users can delete contractors" ON public.contractors FOR DELETE TO authenticated USING (true);
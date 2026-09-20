CREATE TABLE public.contractors (
  id UUID NOT NULL DEFAULT gen_random_uuid() PRIMARY KEY,
  name TEXT NOT NULL,
  email TEXT,
  department TEXT,
  function TEXT,
  country TEXT,
  sow_name TEXT,
  sow_start_date DATE,
  sow_end_date DATE,
  termination_date DATE,
  addis_status TEXT,
  manager TEXT,
  notes TEXT,
  created_at TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now(),
  updated_at TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.contractors TO anon;
GRANT SELECT, INSERT, UPDATE, DELETE ON public.contractors TO authenticated;
GRANT ALL ON public.contractors TO service_role;

ALTER TABLE public.contractors ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone can view contractors" ON public.contractors FOR SELECT USING (true);
CREATE POLICY "Anyone can insert contractors" ON public.contractors FOR INSERT WITH CHECK (true);
CREATE POLICY "Anyone can update contractors" ON public.contractors FOR UPDATE USING (true) WITH CHECK (true);
CREATE POLICY "Anyone can delete contractors" ON public.contractors FOR DELETE USING (true);

CREATE INDEX idx_contractors_email ON public.contractors (email);
CREATE INDEX idx_contractors_sow_end_date ON public.contractors (sow_end_date);

CREATE OR REPLACE FUNCTION public.update_updated_at_column() RETURNS TRIGGER AS $$ BEGIN NEW.updated_at = now(); RETURN NEW; END; $$ LANGUAGE plpgsql SET search_path = public;

CREATE TRIGGER update_contractors_updated_at BEFORE UPDATE ON public.contractors FOR EACH ROW EXECUTE FUNCTION public.update_updated_at_column();
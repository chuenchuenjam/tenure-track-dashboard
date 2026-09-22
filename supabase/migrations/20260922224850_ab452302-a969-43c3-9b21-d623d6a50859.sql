ALTER TABLE public.contractors
  ADD COLUMN IF NOT EXISTS vendor text,
  ADD COLUMN IF NOT EXISTS monthly_rate numeric(12,2),
  ADD COLUMN IF NOT EXISTS currency text,
  ADD COLUMN IF NOT EXISTS renewal_count integer NOT NULL DEFAULT 0;

CREATE INDEX IF NOT EXISTS idx_contractors_vendor ON public.contractors (vendor);
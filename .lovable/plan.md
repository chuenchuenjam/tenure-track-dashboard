# Dashboard upgrade: 8 items

## 1. Choose which columns the main table shows
- A "Columns" button above the table opens a list of checkboxes, so each user can pick the columns they want to see.
- Your choice is remembered on this computer. A "Reset" option goes back to the default columns.

## 2. Contracts ending in 90 / 180 days
- Two cards: "Ending within 90 days" and "Ending within 180 days". Click either one to see the list of people.
- The Renewal queue gets a 90 / 180 day switch.

## 3. Monthly resource count by vendor (new)
- At the moment there is only a "monthly spend" chart. This adds a stacked bar chart showing how many active people each vendor had in each of the past 12 months.
- Click a bar to see the matching people.

## 4. Conflict alert when uploading
- When an upload finds a person with the same email whose details differ from the current record (dates, manager, role and so on), a "Conflicts" screen appears before importing.
- Each person is shown with every differing field side by side: old value vs new value. You can pick per field, or use "Keep all current" / "Use all from file".
- New people and records with no differences are imported as usual.

## 5. Reports based on the current filters
- "Save as report": give the current filters (search, status, department, vendor, work type, drill-down, column choice) a name and save them. Saved reports appear in a "Reports" menu, so one click restores that view.
- "Download report": a multi-sheet Excel file with Summary (KPI totals plus which filters were used), Breakdown (counts by function, country, vendor and work type), and Details (the filtered list using the columns you chose).
- Saved reports are shared by everyone who is signed in, and each report records who created it.

## 6. Upload an FTE list
- Uses the same table, with a new **Work type** field (Contractor / FTE).
- The upload screen lets you choose "This file is: Contractors / FTE". Rows with no work type use that choice.
- Add a work type filter. All cards and charts can show contractors only, FTEs only, or everyone (there is a switch at the top, default is everyone).
- FTEs who have no SoW end date count as Active and do not trigger expiry alerts.

## 7. New fields in the main table
- Direct manager (the current "Manager" is renamed to this, and the existing data is kept), Role, Work type, Contract type.
- These can be edited in the side panel, are guessed automatically when matching upload columns, and are included in the Excel export.

## 8. Project end date = SoW end date
- "Project end date" is shown as the same field as the SoW end date, not stored separately. When matching upload columns, "Project end date" automatically maps to the SoW end date. The label reads "SoW / Project end date".

## Technical details
- Migration: `contractors` gets `role text`, `work_type text not null default 'Contractor'`, `contract_type text`; `manager` stays (labelled Direct manager in the UI).
- New table `saved_reports` (id, name, filters jsonb, columns jsonb, created_by uuid default auth.uid(), timestamps). GRANT to authenticated + service_role, RLS: authenticated can read; insert/update/delete only by created_by = auth.uid().
- Column visibility stored in localStorage (read inside useEffect to avoid hydration mismatch).
- Conflict diff: in UploadDialog, after mapping, fetch existing rows by email, compare each mapped field (empty values in the file do not count as conflicts), and keep a per-field choice map before writing. Renewal count logic stays unchanged.
- Multi-sheet export with xlsx `book_append_sheet`.
- Vendor monthly headcount: for each month, count people whose start <= month end and (end/termination is empty or >= month start), grouped by vendor.
- Fill sample data with role / work_type / contract_type, plus a few FTE rows.

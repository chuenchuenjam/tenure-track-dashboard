# Contractor Dashboard for HR

A single-page internal dashboard where HR can upload a contractor list from Excel/CSV, see who is expiring or still active, edit records inline, and export everything back to Excel.

## Data stored per contractor

Name, email, department, function, country, SoW name, SoW start date, SoW end date, termination date, ADDIS status (still in internal system: yes/no), account status (Active / Expiring / Expired / Terminated), manager, notes.

Status is derived automatically from the dates:
- Terminated — has a termination date in the past
- Expired — SoW end date already passed
- Expiring — SoW end date within the next 180 days
- Active — everything else

## Upload

- Drag-and-drop or pick an .xlsx / .csv file.
- Column mapping screen: the app guesses each column, HR can correct any mapping before importing.
- Preview of rows with any problems (bad dates, missing name) flagged before confirming.
- Import options: replace all data, or merge by email (update existing, add new).

## Dashboard cards

- Total contractors
- Active contractors
- Expiring within 180 days
- Still active in internal system but SoW ended (the "should not be here" group)

## Charts

- Count of termination date by year (bar chart)
- Split by function (bar chart)
- Split by country (donut chart)
- Upcoming expiries by month for the next 180 days

Every card and every chart segment is clickable and drills down to the matching list of individual contractors.

## Contractor table

- Search by name/email, filters for status, department, function, country, date range
- Sortable columns
- Click a row to open a side panel to edit any field, including changing the SoW end date or termination date; status recalculates immediately
- Add a contractor manually, delete a contractor
- Flags for rows needing attention (ended SoW but still active in system)

## Export

- Export to Excel button — exports the current filtered view or all records, with all fields plus the derived status.

## Technical notes

- Lovable Cloud for storage: one `contractors` table plus an `import_batches` table for upload history. Public access (no login) as requested; can be locked down later.
- `xlsx` (SheetJS) for parsing uploads and generating the Excel export, `recharts` for the charts.
- Imports are parsed in the browser and written in bulk; derived status is computed in app code so date edits update the dashboard instantly.
- Drill-down is done with URL filter params so a drilled view can be shared/bookmarked.

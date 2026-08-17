# Subject Exam Analysis Dashboard

A browser-based tool for BISR subject leaders. Upload exam results, get a Value Added dashboard, and export a Word report.

**Live:** https://mrsalahsalah.github.io/Exam-analysis-dashboard-visualizer--Word-export-version/

## What it does

Upload a CSV or Excel file of exam results, one row per student. The tool detects which column is which, converts each student's CAT4 target and actual grade to a common scale, and calculates Value Added (actual minus target).

The dashboard shows:

- Grade distribution, actual and CAT4 target side by side
- Value Added by class and by subgroup (gender, SEND, EAL, more able)
- A target-vs-actual scatter chart, with letter grades for AS/A-Level and numbers for GCSE
- A table of students below their target, sorted worst first, with space for intervention notes

From there, export a Word document with the data plus a few sentences of automatic commentary.

## How to use it

1. Open the live link above.
2. Upload a CSV or Excel file.
3. Check the column mapping the tool guessed, and fix anything wrong.
4. Confirm GCSE, AS, or A-Level.
5. Read the dashboard.
6. Click Export to Word.

No installation, no login, no account needed.

## Privacy

Everything runs in the browser. No server, no upload to any backend. Close the tab and the data is gone unless exported.

## Files in this repo

| File | What it is |
|---|---|
| `index.html` | The deployed app. Fully self-contained, works offline once loaded. |
| `subject-exam-analysis-word-export-v2.jsx` | The React source. Edit this, then rebuild `index.html` from it. |
| Other `.jsx` / `.html` files | Earlier versions, kept for reference. |

## Tech stack

React, Recharts, PapaParse, SheetJS. No backend, no database, no build step for the deployed file, since all libraries are embedded directly in `index.html`.

## Maintained by

Salah, Computer Science, BISR.

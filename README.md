# Fileshare Access Auditor

A small desktop app (Tkinter GUI) that supports the user-access audit workflow:
it reads raw fileshare access data from a source Excel file, extracts the
server and share-path pieces, and appends formatted rows into a structured
audit log workbook.

## Setup

```bash
cd ~/Projects/app_1
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
python3 main.py
```

## How it works

**Source file** (raw access data), read by column position:
- Column A = server
- Column B = Access Path

**Extraction rules:**
- Full server value (column A) → target column **E**
- Substring of the Access Path after `/vol/` (or `\vol\`) → target column **F**

**Target file** ("2c.Fileshares" tab):
- Rows are **appended only** — existing rows are never touched or overwritten.
- Column **A** gets an auto-incrementing ID (`FS001`, `FS002`, ...), continuing
  from whatever the highest existing ID in that column already is.
- Columns B–D are left blank; only A, E, and F are written by this tool.
- If the target workbook or the "2c.Fileshares" sheet doesn't exist yet, it's
  created automatically (with a basic header row).

## Usage steps

1. **Browse** to the source workbook, pick the sheet, and confirm whether row 1
   is a header (checkbox, on by default — it's skipped when checked).
2. **Browse** to the target audit log workbook (existing or new).
3. Click **1. Preview Extraction** — review the computed ID / server / extracted
   path for every row. Rows with issues (missing server, no `/vol/` match) are
   highlighted in yellow with a note.
4. Once it looks right, click **2. Confirm & Append to Log** to write the rows
   and save the target file.

Nothing is written to the target file until you explicitly confirm.

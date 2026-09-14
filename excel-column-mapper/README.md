# Excel Column Mapper

A small desktop app (Tkinter GUI) that copies rows from a source Excel sheet
into a target workbook using column mappings you define at runtime. You choose
which source column goes to which target column, which tab to write to, and
whether each value is copied whole or trimmed at a delimiter.

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

Or double-click `Launch Excel Column Mapper.command`.

## How it works

**Column mappings** — the core of the app. Each mapping is one row:

```
From [ A - Server ]  →  to [ E ]   ☐ extract after [        ]   ✕
```

- **From** — the source column. The dropdown shows row 1's value as a hint
  (e.g. `A - Server`), so you can pick by header name rather than counting
  letters.
- **to** — the target column the value is written to.
- **extract after** — optional. Tick it and enter a delimiter to keep only the
  text *after* that delimiter; leave it unticked to copy the cell whole.
- **✕** — removes that mapping.

Use **+ Add mapping** for as many columns as you need. The app starts with two
mappings (`A → E`, and `B → F` trimmed at `/vol/`), which match the original
fileshare-audit workflow.

**Delimiter matching** is case-insensitive, and `/` and `\` are interchangeable,
so `/vol/` matches both `/vol/finance` and `\vol\finance`. The delimiter is
treated as literal text, so regex characters like `.` behave as typed. If the
delimiter isn't found in a row, that cell is left empty and the row is flagged.

**Auto ID** (optional) writes an incrementing `FS001`, `FS002`, ... into a target
column of your choice. Both the prefix and the column are configurable, and it
continues from the highest existing ID already in that column. Untick **Auto ID**
to skip it entirely.

**Target workbook** — **Browse…** opens your *existing* workbook and edits it in
place. Other tabs, other rows, and unmapped columns are all left alone. Use
**New…** only when you actually want to start a fresh workbook.

**Target tab** — the dropdown lists the actual tabs in the chosen target
workbook. It repopulates whenever the target path changes, whether you browsed
to the file, typed the path, or pasted it, and the hint beside it shows how many
tabs were found. Pick one, or type a new name to create that tab. A header row
is written only when the tab is created.

**Appending** — rows are always appended; existing rows are never touched or
overwritten.

> **Note on formatting:** saving goes through `openpyxl`, which rebuilds the
> workbook file. Cell values, formulas, and sheet structure survive, and macros
> are preserved in `.xlsm` files. Charts, images, and pivot tables are *not*
> carried over — if your audit log contains those, keep a master copy.

If the target file is open in Excel, the save will fail with a "file is locked"
message; close it there and click Confirm again. Nothing is written until that
save succeeds.

## Usage steps

1. **Browse** to the source workbook and pick the source tab. Confirm whether
   row 1 is a header (checkbox, on by default — it's skipped when checked).
2. **Browse** to the target workbook (existing or new) and choose the target tab.
3. Set up your **column mappings** — add, remove, and configure delimiters.
4. Configure **Auto ID**, or untick it.
5. Click **Append to Log**.

**The preview is live.** There's no preview button — the table updates as you
edit, and always shows exactly what will be written. Its columns follow your
mappings, so you can see what lands where. Rows with issues (empty source cell,
delimiter not found) are highlighted in yellow with a note.

If the settings aren't valid yet, the status line below the table says why and
**Append to Log** stays greyed out. Appending recomputes everything from the
current settings first, so what gets written always matches what's on screen.

Nothing is written to the target file until you click Append.

## Validation

**Append to Log** stays greyed out, with the reason shown in the status line,
when:

- no mappings are defined,
- a source or target column is invalid,
- two mappings (or a mapping and the Auto ID) write to the **same** target
  column,
- the Auto ID prefix or target tab name is blank,
- the target tab name uses characters Excel forbids (`[ ] : * ? / \`) or is
  longer than 31 characters.

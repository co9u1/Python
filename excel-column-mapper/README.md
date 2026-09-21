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

Configuration is split across three tabs — **Files**, **Column Mappings**, and
**Lookups** — with the live preview always visible underneath.

**Column mappings** — the core of the app. Each mapping is one row:

```
From [ A - Server ]  →  to [ E - Server Name ]   ☐ extract after [        ]   ✕
```

- **From** — the source column. The dropdown shows row 1's value as a hint
  (e.g. `A - Server`), so you can pick by header name rather than counting
  letters.
- **to** — the target column the value is written to. With **First row is a
  header** ticked on the target workbook, this dropdown is labelled from the
  target tab's own header row too (e.g. `E - Server Name`), so both sides of a
  mapping read by name. Untick it — or pick a tab with no header row — and it
  falls back to plain column letters.
- **extract after** — optional. Tick it and enter a delimiter to keep only the
  text *after* that delimiter; leave it unticked to copy the cell whole.
- **✕** — removes that mapping.

Use **+ Add mapping** for as many columns as you need. The app starts with two
straight copies (`A → E` and `B → F`); **extract after** is unticked and its box
empty on every new mapping, so values are copied whole unless you ask otherwise.

**Delimiter matching** is case-insensitive, and `/` and `\` are interchangeable,
so `/vol/` matches both `/vol/finance` and `\vol\finance`. The delimiter is
treated as literal text, so regex characters like `.` behave as typed. If the
delimiter isn't found in a row, that cell is left empty and the row is flagged.

**Lookups** (the **Lookups** tab) pull a value from a *third* workbook — a
VLOOKUP. Use this when the source file has one piece of information and the
target needs a related one that lives in a reference table. Each lookup is:

```
Lookup file [ sites.xlsx ] [Browse…]   tab [ Sites ]              ✕
Key from source [ A - Server ]  match on [ A - Server Name ]  return [ B - Site ]
Write result to [ G - Site ]   if no match, write [ UNKNOWN ]
```

- **Key from source** — the source column whose value is looked up.
- **match on** / **return** — the key and value columns *in the lookup file*.
  Both dropdowns are labelled from that file's header row.
- **Write result to** — the target column the result lands in.
- **if no match** — optional fallback text. Leave it blank to write nothing.

Matching **ignores case and surrounding whitespace**, so `SRV-NAS02 ` matches
`srv-nas02`. If the lookup file has duplicate keys, the first one wins, the same
as VLOOKUP. Unmatched rows are flagged yellow in the preview with a note naming
the value that didn't match.

Add as many lookups as you need; each gets its own file, tab, and columns.

**Auto ID** (optional) writes an incrementing `FS001`, `FS002`, ... into a target
column of your choice — its dropdown is header-labelled the same way. Both the
prefix and the column are configurable, and it
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

**Matching the sheet's formatting** — openpyxl writes unstyled cells, so
appended rows would otherwise stand out against a formatted log. With **Match
the formatting of row 2** ticked (the default), every appended row copies row
2's font, fill, borders, alignment, number format, and row height. Row 2 is the
first data row under the header, so it's the natural template.

Formatting is copied across the **full width** of row 2, including columns you
haven't mapped, so banded fills and borders stay unbroken. If the target tab has
no row 2 yet — a tab the app just created, or one with only headers — appended
rows simply keep default formatting.

**Extending validation & rules** — Excel stores a dropdown as a rule over a
fixed range like `G2:G500`. Rows appended past the end of that range get no
dropdown, which looks like the validation was lost. With **Extend validation &
rules** ticked (the default), any **data validation** or **conditional
formatting** rule that already covers row 2 is stretched down to the last
appended row. Rules that don't touch row 2 are left exactly as they are.

**Preserving Excel's extended validations** — a validation whose source is a
range on *another sheet* (`=Lists!$A$1:$A$20`) isn't stored in the normal place.
Excel puts it in an `x14` extension block, and openpyxl neither reads nor writes
that block — it warns *"Data Validation extension is not supported and will be
removed"* and drops it on save. Simple inline lists like `"Yes,No"` live in the
standard element and survive, which is why those were the only ones left.

The app now reads those extension blocks straight out of the source `.xlsx`
before openpyxl opens it, and splices them back in afterwards, so cross-sheet
dropdowns survive. Their ranges are extended to cover appended rows too.

> **Excel Tables are not extended.** If your log is a real Excel Table
> (Insert → Table), its range still won't grow to include appended rows, so they
> land just outside the table. Convert the range to a normal one, or extend the
> table by hand afterwards.

> **Extended conditional formatting is lost.** Data bars, icon sets and colour
> scales configured through Excel's newer options are stored as `x14` blocks
> nested inside each rule. openpyxl drops those when it rewrites the sheet, and
> unlike the validation blocks they can't be spliced back reliably because the
> rules they attach to are themselves rewritten. Plain rules — "highlight cells
> equal to", simple colour scales — survive normally.

## Safety

The target workbook is copied before any write. After saving, the result is
checked: every XML part must parse and openpyxl must be able to reopen the file.
**If that check fails, your original is put back unchanged** and the error is
reported — a bad write can't leave you with a broken workbook.

**Keep a timestamped backup** (on by default) additionally leaves a
`yourfile.backup-YYYYMMDD-HHMMSS.xlsx` beside the original on every successful
append. Untick it once you trust the tool on your file.

Sheet protection, workbook protection and their passwords are preserved.

## Removing sheet protection

**Remove sheet protection from …** (off by default) unlocks sheets as part of
the append. The dropdown beside it chooses the scope:

- **the target tab** *(default)* — only the tab being appended to; every other
  sheet keeps its protection.
- **all sheets** — every worksheet, including lookup and reference tabs.

This clears the `<sheetProtection>` element and the stored password hash
outright. It touches nothing else: data validation, conditional formatting, cell
styles, row heights and cell contents are all unaffected, because protection is
a separate element from every one of them.

Verified by running the same workbook three ways — protection left on, target
tab unlocked, all sheets unlocked — and confirming the only difference was the
protection flag. Validation ranges, conditional formatting, fonts, fills,
borders, alignment, number formats and row heights were identical in all three.

Workbook **structure** protection (the lock on adding, deleting or renaming
sheets) is *not* touched by this option and is always preserved.

## Keeping appended rows editable

Excel blocks a cell only when **both** are true: the sheet is protected *and*
the cell is locked. Cells are locked by default, so a protected sheet blocks
everything except cells that were deliberately unlocked.

Appended rows inherit the template row's locked flag along with its formatting.
If that row is locked — or if **Match the formatting of row 2** is off, so cells
get the locked-by-default style — the new rows refuse edits on a protected
sheet. Clicking a validation dropdown in one gives *"The cell you're trying to
change is on a protected sheet."*

**Unlock the cells in appended rows** (off by default) marks just those rows
unlocked, so they behave like your existing data rows while the sheet stays
protected. It's the usual fix if you want protection *and* usable dropdowns —
removing protection isn't necessary.

Unlocking covers the sheet's full used width, not only the mapped columns, so a
dropdown in a column the app never writes to still works on the new rows.
Existing rows are never modified.

Use `check_protection.py` to see the state of any workbook:

```bash
.venv/bin/python check_protection.py Audit.xlsx Log G2 G3 G500
```

It reports workbook and per-sheet protection, and for each cell whether it's
locked and whether Excel will therefore block it.

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
5. Click **1. Generate Preview**, check the table, then **2. Append to Log**.

**The preview is generated on demand**, so nothing is re-read while you're still
setting things up — which matters on large sheets. Its columns follow your
mappings, so you can see what lands where. Rows with issues (empty source cell,
delimiter not found, no lookup match) are highlighted in yellow with a note.

**Append is locked until the preview matches your settings.** Change anything —
a mapping, a lookup, a tab — and the table clears, the status line reads
`Settings changed - click Generate Preview`, and **Append** greys out until you
regenerate. That's what stops the app from writing something different from what
you're looking at.

If the settings aren't valid, the status line says exactly why. Appending also
recomputes from the current settings immediately before writing, so a file that
changed on disk can't slip through.

Nothing is written to the target file until you click Append.

## Validation

**Append to Log** stays greyed out, with the reason shown in the status line,
when:

- no mappings **or lookups** are defined,
- a lookup's file is missing, its tab isn't chosen, or its tab can't be read,
- a source or target column is invalid,
- two outputs (any combination of mapping, lookup, or the Auto ID) write to the
  **same** target column,
- the Auto ID prefix or target tab name is blank,
- the target tab name uses characters Excel forbids (`[ ] : * ? / \`) or is
  longer than 31 characters.

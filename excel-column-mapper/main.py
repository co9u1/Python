"""
Excel Column Mapper
-------------------
Reads rows from a source Excel sheet and appends them into a chosen tab of
a target workbook, using column mappings defined at runtime.

Each mapping is: source column -> target column, with an optional
"extract after <delimiter>" rule that can be switched on or off per
mapping. Mappings can be added and removed freely.

An optional auto-incrementing ID (prefix + zero-padded number) can be
written to a target column of your choice, continuing from the highest
existing ID already in that column.

The preview table is live: it rebuilds from the current settings on every
edit, so it always shows exactly what will be written. Appending
recomputes from scratch first, so the two can never drift apart.

The target workbook is edited in place and rows are only ever appended -
existing rows and other tabs are left untouched.

Note: saving goes through openpyxl, which rewrites the workbook file.
Values, formulas and (for .xlsm) macros survive; charts, images and pivot
tables do not.
"""

import os
import re
import shutil
import zipfile
from copy import copy
from datetime import datetime
from xml.etree import ElementTree

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import MultiCellRange

APP_TITLE = "Excel Column Mapper"
DEFAULT_ID_PREFIX = "FS"
DEFAULT_ID_COL = "A"
DEFAULT_DELIMITER = "/vol/"
ID_PAD = 3

# Excel's own limits on worksheet names.
INVALID_SHEET_CHARS = "[]:*?/\\"
MAX_SHEET_NAME = 31

# Appended rows copy their formatting from this row of the target tab.
STYLE_TEMPLATE_ROW = 2

# How often to report progress (in rows) while writing a large append.
PROGRESS_EVERY = 2000

COLUMN_CHOICES = [get_column_letter(i) for i in range(1, 41)]  # A..AN


def extract_after_delimiter(value: str, delimiter: str) -> str:
    """Return the substring after `delimiter`.

    A blank delimiter means "take the whole cell". Matching is
    case-insensitive, and '/' and '\\' are interchangeable so either path
    style matches. Returns '' when the delimiter isn't found.
    """
    if value in (None, ""):
        return ""
    text = str(value).strip()

    delim = (delimiter or "").strip()
    if not delim:
        return text

    pattern = "".join(r"[\\/]" if ch in "\\/" else re.escape(ch) for ch in delim)
    match = re.search(pattern, text, re.IGNORECASE)
    return text[match.end():].strip() if match else ""


def parse_column(text: str):
    """Parse 'E' or 'A - Server' into a 1-based column index, or None."""
    if not text:
        return None
    match = re.match(r"^\s*([A-Za-z]{1,3})(?:\b|$)", str(text))
    if not match:
        return None
    try:
        return column_index_from_string(match.group(1).upper())
    except ValueError:
        return None


def next_id_number(ws, prefix: str, col_idx: int) -> int:
    """Scan a target column for existing <prefix><digits> ids, return the next number."""
    max_n = 0
    pattern = re.compile(rf"^{re.escape(prefix)}(\d+)$", re.IGNORECASE)
    for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
        val = row[0]
        if not val:
            continue
        m = pattern.match(str(val).strip())
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def row_style_template(ws):
    """Snapshot row 2's formatting so appended rows can match the sheet.

    Row 2 is the first data row under the header. Returns None when the
    sheet has no such row yet (a tab the app just created, for example),
    in which case appended rows are left with default formatting.
    """
    if ws.max_row < STYLE_TEMPLATE_ROW or not ws.max_column:
        return None

    # A cell's formatting is a StyleArray of indices into the workbook's
    # shared style tables. Copying that small array reuses the registered
    # styles; copying font/fill/border objects per cell instead is orders of
    # magnitude slower and defeats openpyxl's style interning entirely.
    styles = [
        ws.cell(row=STYLE_TEMPLATE_ROW, column=col)._style
        for col in range(1, ws.max_column + 1)
    ]
    height = ws.row_dimensions[STYLE_TEMPLATE_ROW].height
    return {"styles": styles, "height": height}


def apply_row_style(ws, row_idx, template):
    """Paint a snapshotted row format onto every column of `row_idx`."""
    if not template:
        return
    for col, style in enumerate(template["styles"], start=1):
        if style is not None:
            ws.cell(row=row_idx, column=col)._style = copy(style)
    if template["height"] is not None:
        ws.row_dimensions[row_idx].height = template["height"]


SHEET_RE = re.compile(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*/>')
REL_RE = re.compile(r"<Relationship\b[^>]*/>")
XM_SQREF_RE = re.compile(r"(<xm:sqref>)([^<]*)(</xm:sqref>)")
EXT_OPEN = "<extLst>"
EXT_CLOSE = "</extLst>"


def worksheet_extlst_span(xml):
    """Locate the worksheet-level <extLst>, returning (start, end) or None.

    <extLst> is not unique in a sheet: conditional formatting rules nest one
    for data bars, icon sets and colour scales. Only the final child of
    <worksheet> is the sheet-level list, so match that one by walking back
    from the close tag and balancing nested pairs. Matching the first
    <extLst> in the file instead splices content into a <cfRule>, which
    Excel rejects as a corrupt workbook.
    """
    end_ws = xml.rfind("</worksheet>")
    if end_ws == -1:
        return None
    close = xml.rfind(EXT_CLOSE, 0, end_ws)
    if close == -1 or xml[close + len(EXT_CLOSE):end_ws].strip():
        return None

    depth = 0
    idx = close
    while idx > 0:
        prev_open = xml.rfind(EXT_OPEN, 0, idx)
        prev_close = xml.rfind(EXT_CLOSE, 0, idx)
        if prev_open == -1:
            return None
        if prev_close > prev_open:
            depth += 1
            idx = prev_close
        elif depth == 0:
            return (prev_open, close + len(EXT_CLOSE))
        else:
            depth -= 1
            idx = prev_open
    return None


def _sheet_part_map(zf):
    """Map worksheet name -> its XML part name inside the xlsx zip."""
    try:
        workbook_xml = zf.read("xl/workbook.xml").decode("utf-8")
        rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    except KeyError:
        return {}

    targets = {}
    for rel in REL_RE.findall(rels_xml):
        rel_id = re.search(r'Id="([^"]+)"', rel)
        target = re.search(r'Target="([^"]+)"', rel)
        if rel_id and target and "worksheets/" in target.group(1):
            part = target.group(1).lstrip("/")
            targets[rel_id.group(1)] = part if part.startswith("xl/") else f"xl/{part}"

    return {
        name: targets[rel_id]
        for name, rel_id in SHEET_RE.findall(workbook_xml)
        if rel_id in targets
    }


def read_validation_extensions(path):
    """Capture each sheet's <extLst> block, keyed by sheet name.

    Excel keeps data validations whose source lives on another sheet in an
    x14 extension list rather than the standard <dataValidations> element.
    openpyxl doesn't model that block and drops it on save - it even warns
    "Data Validation extension is not supported and will be removed" - so
    those dropdowns vanish from the saved file. Grab the raw XML before
    openpyxl touches the workbook so it can be put back afterwards.
    """
    if not path or not os.path.exists(path):
        return {}
    try:
        with zipfile.ZipFile(path) as zf:
            found = {}
            for name, part in _sheet_part_map(zf).items():
                try:
                    xml = zf.read(part).decode("utf-8")
                except KeyError:
                    continue
                span = worksheet_extlst_span(xml)
                if not span:
                    continue
                block = xml[span[0]:span[1]]
                if "x14:dataValidation" in block:
                    found[name] = block
            return found
    except (zipfile.BadZipFile, OSError):
        return {}


def restore_validation_extensions(path, blocks, stretch=None):
    """Splice preserved <extLst> blocks back into a workbook openpyxl saved.

    `stretch` is (sheet_name, last_row); that sheet's extension ranges are
    extended the same way the standard rules are.
    """
    if not blocks:
        return
    try:
        with zipfile.ZipFile(path) as zf:
            parts = _sheet_part_map(zf)
            payload = [(item, zf.read(item.filename)) for item in zf.infolist()]
    except (zipfile.BadZipFile, OSError):
        return

    wanted = {parts[name]: block for name, block in blocks.items() if name in parts}
    if not wanted:
        return

    if stretch:
        sheet_name, last_row = stretch
        part = parts.get(sheet_name)
        if part in wanted:
            wanted[part] = _stretch_extension_ranges(wanted[part], last_row)

    tmp_path = f"{path}.tmp-ext"
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as out:
            for item, data in payload:
                block = wanted.get(item.filename)
                if block:
                    xml = data.decode("utf-8")
                    span = worksheet_extlst_span(xml)
                    inner = block[len(EXT_OPEN):-len(EXT_CLOSE)]
                    if span:
                        # A worksheet may carry only one sheet-level extLst,
                        # so merge into the existing one rather than append.
                        existing = xml[span[0]:span[1]]
                        merged = existing[: -len(EXT_CLOSE)] + inner + EXT_CLOSE
                        xml = xml[: span[0]] + merged + xml[span[1]:]
                    else:
                        xml = xml.replace("</worksheet>", f"{block}</worksheet>", 1)
                    data = xml.encode("utf-8")
                out.writestr(item, data)
        shutil.move(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def workbook_is_readable(path):
    """Cheap structural check that Excel stands a chance of opening this.

    Every XML part must parse and openpyxl must be able to reload the file.
    Not a guarantee Excel is happy, but it catches malformed output before
    it replaces the user's workbook.
    """
    try:
        with zipfile.ZipFile(path) as zf:
            if zf.testzip() is not None:
                return False
            for name in zf.namelist():
                if name.endswith((".xml", ".rels")):
                    ElementTree.fromstring(zf.read(name))
        openpyxl.load_workbook(path, read_only=True).close()
        return True
    except Exception:
        return False


def _stretch_extension_ranges(block, last_row):
    """Extend <xm:sqref> ranges that cover the template row."""
    def fix(match):
        refs = []
        for ref in match.group(2).split():
            bounds = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
            if bounds:
                top, bottom = int(bounds.group(2)), int(bounds.group(4))
                if top <= STYLE_TEMPLATE_ROW <= bottom < last_row:
                    ref = f"{bounds.group(1)}{top}:{bounds.group(3)}{last_row}"
            refs.append(ref)
        return match.group(1) + " ".join(refs) + match.group(3)

    return XM_SQREF_RE.sub(fix, block)


def _stretched(sqref, last_row):
    """Extend any sub-range covering the template row down to `last_row`.

    Returns (MultiCellRange, changed). MultiCellRange.ranges is a set, so the
    ranges are pulled into a list before being mutated and rebuilt.
    """
    ranges = list(MultiCellRange(str(sqref)).ranges)
    changed = False
    for cell_range in ranges:
        covers_data = cell_range.min_row <= STYLE_TEMPLATE_ROW <= cell_range.max_row
        if covers_data and cell_range.max_row < last_row:
            cell_range.max_row = last_row
            changed = True
    return MultiCellRange(ranges), changed


def extend_data_validations(ws, last_row):
    """Stretch validation rules that cover the template row down to `last_row`.

    Excel stores a dropdown as a rule over a fixed range such as G2:G500.
    Rows appended past the end of that range get no dropdown, which looks
    like the validation was lost. Any rule already covering row 2 is treated
    as applying to the data rows, so it follows them down. Rules that don't
    touch row 2 are left exactly as they are.
    """
    for rule in ws.data_validations.dataValidation:
        stretched, changed = _stretched(rule.sqref, last_row)
        if changed:
            rule.sqref = stretched


def extend_conditional_formatting(ws, last_row):
    """Same treatment for conditional formatting ranges.

    The rules have to be re-added rather than edited in place, because the
    list is keyed by range.
    """
    saved = [(str(cf.sqref), list(cf.rules)) for cf in ws.conditional_formatting]
    if not saved:
        return
    ws.conditional_formatting = ConditionalFormattingList()
    for sqref, rules in saved:
        stretched, _changed = _stretched(sqref, last_row)
        for rule in rules:
            ws.conditional_formatting.add(str(stretched), rule)


def last_used_row(ws) -> int:
    """Return the last row index (1-based) that has any content, 0 if empty."""
    last = 0
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        if any(cell.value not in (None, "") for cell in row):
            last = row_idx
    return last


# ---------- Color palette / fonts ----------
BG_MAIN = "#eef1f8"
BG_CARD = "#ffffff"
BORDER = "#dde2ec"
SHADOW = "#d2d7e6"
HEADER_BG = "#20263f"
HEADER_FG = "#ffffff"
HEADER_SUB_FG = "#aab3cc"
ACCENT = "#3b5bfd"
ACCENT_ACTIVE = "#2c46d1"
ACCENT_DISABLED = "#b9c3f7"
SECONDARY_BG = "#e7eaf3"
SECONDARY_ACTIVE = "#d9deec"
SECONDARY_FG = "#2a3050"
TEXT_DARK = "#1c2233"
TEXT_MUTED = "#6b7280"
WARN_BG = "#fff3cd"
ROW_ALT = "#f5f7fc"
SUCCESS = "#1e8f4e"
WARNTEXT = "#b8860b"
ERROR = "#d64545"

FONT = "Avenir Next"


def _tint(hex_color: str, amount: float = 0.85) -> str:
    """Blend a hex color toward white; used for soft pastel badge backgrounds."""
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = int(r + (255 - r) * amount)
    g = int(g + (255 - g) * amount)
    b = int(b + (255 - b) * amount)
    return f"#{r:02x}{g:02x}{b:02x}"


def _round_rect_points(x1, y1, x2, y2, r):
    return [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]


class PillButton(tk.Canvas):
    """A rounded, hover-responsive button drawn on a canvas (real pill shape)."""

    def __init__(
        self,
        parent,
        text,
        command=None,
        bg_page=BG_CARD,
        fill=ACCENT,
        fill_active=ACCENT_ACTIVE,
        fill_disabled=ACCENT_DISABLED,
        fg="#ffffff",
        font=(FONT, 11, "bold"),
        padx=20,
        pady=11,
    ):
        weight = font[2] if len(font) > 2 else "normal"
        f = tkfont.Font(family=font[0], size=font[1], weight=weight)
        text_w = f.measure(text)
        text_h = f.metrics("linespace")
        btn_w = text_w + padx * 2
        btn_h = text_h + pady * 2

        super().__init__(
            parent,
            width=btn_w,
            height=btn_h,
            bg=bg_page,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        # NOTE: tkinter.Widget reserves the `_w` attribute for the Tcl path
        # name, so button dimensions are kept under different names.
        self._btn_w = btn_w
        self._btn_h = btn_h
        self.command = command
        self.text = text
        self.font = font
        self.fg = fg
        self.fill = fill
        self.fill_active = fill_active
        self.fill_disabled = fill_disabled
        self._state = "normal"

        self._render(fill)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _render(self, color):
        self.delete("all")
        r = self._btn_h / 2
        pts = _round_rect_points(1, 1, self._btn_w - 1, self._btn_h - 1, r)
        self.create_polygon(pts, smooth=True, fill=color, outline="")
        self.create_text(
            self._btn_w / 2, self._btn_h / 2, text=self.text, fill=self.fg, font=self.font
        )

    def _on_enter(self, _event):
        if self._state == "normal":
            self._render(self.fill_active)

    def _on_leave(self, _event):
        if self._state == "normal":
            self._render(self.fill)

    def _on_click(self, _event):
        if self._state == "normal" and self.command:
            self.command()

    def set_state(self, state):
        self._state = state
        if state == "disabled":
            self._render(self.fill_disabled)
            self.configure(cursor="arrow")
        else:
            self._render(self.fill)
            self.configure(cursor="hand2")


class CircleBadge(tk.Canvas):
    """A small filled circle with a centered emoji/glyph, for the header icon."""

    def __init__(self, parent, glyph, diameter=48, bg_page=HEADER_BG, fill=ACCENT):
        super().__init__(
            parent, width=diameter, height=diameter, bg=bg_page,
            highlightthickness=0, bd=0,
        )
        r = diameter / 2
        self.create_oval(2, 2, diameter - 2, diameter - 2, fill=fill, outline="")
        self.create_text(r, r, text=glyph, font=(FONT, int(diameter * 0.42)))


class StatusBadge(tk.Canvas):
    """A rounded pill with a status dot + message; auto-sizes to its text."""

    def __init__(self, parent, bg_page=BG_MAIN, font=(FONT, 11)):
        super().__init__(parent, width=1, height=1, bg=bg_page, highlightthickness=0, bd=0)
        self.font_spec = font
        self._badge_font = tkfont.Font(family=font[0], size=font[1])
        self.set("", SUCCESS)

    def set(self, text, color):
        self.delete("all")
        if not text:
            self.configure(width=1, height=1)
            return
        pad_x, pad_y, dot_r, gap = 14, 8, 4, 8
        text_w = self._badge_font.measure(text)
        text_h = self._badge_font.metrics("linespace")
        w = pad_x * 2 + dot_r * 2 + gap + text_w
        h = text_h + pad_y * 2
        self.configure(width=w, height=h)
        r = h / 2
        pts = _round_rect_points(0, 0, w, h, r)
        self.create_polygon(pts, smooth=True, fill=_tint(color), outline="")
        cx, cy = pad_x + dot_r, h / 2
        self.create_oval(cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r, fill=color, outline="")
        self.create_text(
            cx + dot_r + gap, cy, text=text, anchor="w", fill=TEXT_DARK, font=self.font_spec
        )


class ScrollArea(tk.Frame):
    """A fixed-height vertically scrollable container for the mapping rows."""

    def __init__(self, parent, bg=BG_CARD, height=170):
        super().__init__(parent, bg=bg)
        self._area_canvas = tk.Canvas(
            self, bg=bg, highlightthickness=0, bd=0, height=height
        )
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._area_canvas.yview)
        self._area_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._area_canvas.pack(side="left", fill="both", expand=True)

        self.inner = tk.Frame(self._area_canvas, bg=bg)
        self._area_window = self._area_canvas.create_window(
            (0, 0), window=self.inner, anchor="nw"
        )
        self.inner.bind("<Configure>", self._on_inner_configure)
        self._area_canvas.bind("<Configure>", self._on_canvas_configure)

        # Wheel events land on whichever child is under the pointer, so grab them
        # globally while the pointer is inside this area and release on the way out.
        self.bind("<Enter>", self._bind_wheel)
        self.bind("<Leave>", self._unbind_wheel)

    def _on_inner_configure(self, _event):
        self._area_canvas.configure(scrollregion=self._area_canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self._area_canvas.itemconfigure(self._area_window, width=event.width)

    def _bind_wheel(self, _event=None):
        self._area_canvas.bind_all("<MouseWheel>", self._on_wheel)

    def _unbind_wheel(self, _event=None):
        self._area_canvas.unbind_all("<MouseWheel>")

    def _on_wheel(self, event):
        bbox = self._area_canvas.bbox("all")
        if not bbox or bbox[3] <= self._area_canvas.winfo_height():
            return  # nothing to scroll; don't jitter
        delta = event.delta
        if abs(delta) >= 120:  # Windows-style notches
            delta = int(delta / 120)
        self._area_canvas.yview_scroll(-delta, "units")


class MappingRow:
    """One source-column -> target-column mapping, with an optional delimiter."""

    def __init__(self, app, parent, source_col="A", target_col="E",
                 delim_enabled=False, delimiter=DEFAULT_DELIMITER):
        self.app = app
        self.source_col = tk.StringVar(value=source_col)
        self.target_col = tk.StringVar(value=target_col)
        self.delim_enabled = tk.BooleanVar(value=delim_enabled)
        self.delimiter = tk.StringVar(value=delimiter)

        self.frame = tk.Frame(parent, bg=BG_CARD)
        self.frame.pack(fill="x", pady=3)

        ttk.Label(self.frame, text="From").pack(side="left")
        self.source_combo = ttk.Combobox(
            self.frame, textvariable=self.source_col, width=18, values=app.source_columns
        )
        self.source_combo.pack(side="left", padx=(6, 8))

        ttk.Label(self.frame, text="→  to").pack(side="left")
        self.target_combo = ttk.Combobox(
            self.frame, textvariable=self.target_col, width=18,
            values=app.target_columns, state="readonly",
        )
        self.target_combo.pack(side="left", padx=(6, 12))

        self.delim_check = ttk.Checkbutton(
            self.frame, text="extract after", variable=self.delim_enabled,
            command=self._sync_delim_state,
        )
        self.delim_check.pack(side="left")

        self.delim_entry = ttk.Entry(self.frame, textvariable=self.delimiter, width=12)
        self.delim_entry.pack(side="left", padx=6)

        for var in (self.source_col, self.target_col, self.delim_enabled, self.delimiter):
            var.trace_add("write", app.schedule_refresh)

        PillButton(
            self.frame, "✕", command=self.remove, bg_page=BG_CARD,
            fill=SECONDARY_BG, fill_active="#f3c9c9", fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 10, "bold"), padx=10, pady=6,
        ).pack(side="right")

        self._sync_delim_state()

    def _sync_delim_state(self):
        self.delim_entry.configure(
            state="normal" if self.delim_enabled.get() else "disabled"
        )

    def remove(self):
        self.app.remove_mapping(self)

    def destroy(self):
        self.frame.destroy()

    def effective_delimiter(self) -> str:
        return self.delimiter.get().strip() if self.delim_enabled.get() else ""


class LookupRow:
    """One VLOOKUP-style rule: a source column keys into an external workbook."""

    def __init__(self, app, parent):
        self.app = app
        self.path = tk.StringVar()
        self.sheet = tk.StringVar()
        self.key_col = tk.StringVar(value="A")
        self.val_col = tk.StringVar(value="B")
        self.source_col = tk.StringVar(value="A")
        self.target_col = tk.StringVar(value="A")
        self.fallback = tk.StringVar()
        self.lookup_columns = list(COLUMN_CHOICES)

        self.frame = tk.Frame(parent, bg=BG_CARD, highlightthickness=1,
                              highlightbackground=BORDER)
        self.frame.pack(fill="x", pady=4)

        # Row A - which workbook and tab holds the lookup table
        top = tk.Frame(self.frame, bg=BG_CARD)
        top.pack(fill="x", padx=8, pady=(7, 3))
        ttk.Label(top, text="Lookup file").pack(side="left")
        ttk.Entry(top, textvariable=self.path, width=34).pack(
            side="left", fill="x", expand=True, padx=(6, 8)
        )
        PillButton(
            top, "Browse...", command=self.pick_file, bg_page=BG_CARD,
            fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE, fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 10, "bold"), padx=12, pady=6,
        ).pack(side="left")
        ttk.Label(top, text="tab").pack(side="left", padx=(10, 4))
        self.sheet_combo = ttk.Combobox(
            top, textvariable=self.sheet, width=16, state="readonly"
        )
        self.sheet_combo.pack(side="left")
        PillButton(
            top, "✕", command=self.remove, bg_page=BG_CARD,
            fill=SECONDARY_BG, fill_active="#f3c9c9", fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 10, "bold"), padx=10, pady=6,
        ).pack(side="right", padx=(8, 0))

        # Row B - which value is matched against which lookup column
        mid = tk.Frame(self.frame, bg=BG_CARD)
        mid.pack(fill="x", padx=8, pady=(0, 2))
        ttk.Label(mid, text="Key from source").pack(side="left")
        self.source_combo = ttk.Combobox(
            mid, textvariable=self.source_col, width=16, values=app.source_columns
        )
        self.source_combo.pack(side="left", padx=(6, 12))
        ttk.Label(mid, text="match on").pack(side="left")
        self.key_combo = ttk.Combobox(
            mid, textvariable=self.key_col, width=15, state="readonly",
            values=self.lookup_columns,
        )
        self.key_combo.pack(side="left", padx=(6, 12))
        ttk.Label(mid, text="return").pack(side="left")
        self.val_combo = ttk.Combobox(
            mid, textvariable=self.val_col, width=15, state="readonly",
            values=self.lookup_columns,
        )
        self.val_combo.pack(side="left", padx=6)

        # Row C - where the result lands
        bot = tk.Frame(self.frame, bg=BG_CARD)
        bot.pack(fill="x", padx=8, pady=(0, 7))
        ttk.Label(bot, text="Write result to").pack(side="left")
        self.target_combo = ttk.Combobox(
            bot, textvariable=self.target_col, width=16, state="readonly",
            values=app.target_columns,
        )
        self.target_combo.pack(side="left", padx=(6, 12))
        ttk.Label(bot, text="if no match, write").pack(side="left")
        ttk.Entry(bot, textvariable=self.fallback, width=14).pack(side="left", padx=6)
        ttk.Label(
            bot, text="(blank leaves the cell empty)", style="Muted.TLabel"
        ).pack(side="left", padx=4)

        for var in (self.source_col, self.key_col, self.val_col,
                    self.target_col, self.fallback):
            var.trace_add("write", app.schedule_refresh)
        for var in (self.path, self.sheet):
            var.trace_add("write", self._on_file_change)

    def pick_file(self):
        path = filedialog.askopenfilename(
            title="Select the lookup workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm")],
        )
        if path:
            self.path.set(path)

    def _on_file_change(self, *_args):
        """Refresh the tab list and relabel the key/return dropdowns."""
        path = self.path.get().strip()
        sheets = []
        if path and os.path.exists(path):
            try:
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                sheets = list(wb.sheetnames)
                wb.close()
            except Exception:
                sheets = []
        self.sheet_combo["values"] = sheets
        if sheets and self.sheet.get().strip() not in sheets:
            self.sheet.set(sheets[0])
            return  # setting sheet re-enters this handler

        labels = self.app.column_labels(path, self.sheet.get().strip())
        self.lookup_columns = labels
        self.app.relabel(self.key_combo, self.key_col, labels)
        self.app.relabel(self.val_combo, self.val_col, labels)
        self.app.schedule_refresh()

    def remove(self):
        self.app.remove_lookup(self)

    def destroy(self):
        self.frame.destroy()


class ColumnMapperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("940x980")
        # Keep enough height that the preview card can never be squeezed flat -
        # it's the only expanding widget, so everything above it is fixed cost.
        self.minsize(880, 860)
        self.configure(bg=BG_MAIN)

        self.source_path = tk.StringVar()
        self.source_sheet = tk.StringVar()
        self.skip_header = tk.BooleanVar(value=True)
        self.target_path = tk.StringVar()
        self.target_sheet = tk.StringVar()
        self.target_has_header = tk.BooleanVar(value=True)
        self.copy_format = tk.BooleanVar(value=True)
        self.extend_validation = tk.BooleanVar(value=True)
        self.keep_backup = tk.BooleanVar(value=True)
        self._target_scan_job = None

        self.id_enabled = tk.BooleanVar(value=True)
        self.id_prefix = tk.StringVar(value=DEFAULT_ID_PREFIX)
        self.id_col = tk.StringVar(value=DEFAULT_ID_COL)

        self.source_columns = list(COLUMN_CHOICES)
        self.target_columns = list(COLUMN_CHOICES)
        self.mappings = []
        self.lookups = []
        self._lookup_cache = {}
        self._preview_rows = []

        # Live-preview plumbing: a debounce handle plus caches so a refresh
        # doesn't re-read the workbooks on every keystroke.
        self._refresh_job = None
        self._preview_stale = True
        self._src_cache_key = None
        self._src_cache_rows = None
        self._tgt_cache_key = None
        self._tgt_cache_nextid = 1

        self._build_style()
        self._build_ui()

        # Start with the two mappings this tool originally hardcoded.
        self.add_mapping("A", "E", False, DEFAULT_DELIMITER)
        self.add_mapping("B", "F", True, DEFAULT_DELIMITER)

    # ---------- Styling ----------
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=BG_CARD)
        style.configure("Main.TFrame", background=BG_MAIN)

        style.configure(
            "Card.TLabelframe",
            background=BG_CARD,
            bordercolor=BORDER,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=BG_CARD,
            foreground=TEXT_DARK,
            font=(FONT, 12, "bold"),
        )

        style.configure("TLabel", background=BG_CARD, foreground=TEXT_DARK, font=(FONT, 11))
        style.configure(
            "Muted.TLabel", background=BG_CARD, foreground=TEXT_MUTED, font=(FONT, 10)
        )
        style.configure(
            "Status.TLabel", background=BG_MAIN, foreground=SUCCESS, font=(FONT, 11, "bold")
        )

        style.configure("TEntry", fieldbackground="#ffffff", padding=6)
        style.configure("TCombobox", fieldbackground="#ffffff", padding=4)
        style.configure("TCheckbutton", background=BG_CARD, font=(FONT, 10))

        style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground=TEXT_DARK,
            rowheight=26,
            font=(FONT, 10),
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            background="#eef1f7",
            foreground=TEXT_DARK,
            font=(FONT, 10, "bold"),
            relief="flat",
        )
        style.map(
            "Treeview",
            background=[("selected", ACCENT)],
            foreground=[("selected", "white")],
        )

    # ---------- UI ----------
    def _card(self, parent, title, expand=False):
        """A card with a soft drop-shadow: a tinted frame peeking out bottom-right."""
        wrap = tk.Frame(parent, bg=SHADOW)
        wrap.pack(fill="both" if expand else "x", expand=expand, padx=14, pady=(0, 9))
        card = ttk.LabelFrame(wrap, text=title, style="Card.TLabelframe")
        card.pack(fill="both" if expand else "x", expand=expand, padx=(0, 3), pady=(0, 3))
        return card

    def _build_ui(self):
        # Header banner
        header = tk.Frame(self, bg=HEADER_BG)
        header.pack(fill="x")

        header_inner = tk.Frame(header, bg=HEADER_BG)
        header_inner.pack(fill="x", padx=18, pady=12)

        CircleBadge(header_inner, "🧭", diameter=48, bg_page=HEADER_BG, fill=ACCENT).pack(
            side="left", padx=(0, 14)
        )
        title_col = tk.Frame(header_inner, bg=HEADER_BG)
        title_col.pack(side="left", fill="x", expand=True)
        tk.Label(
            title_col,
            text=APP_TITLE,
            bg=HEADER_BG,
            fg=HEADER_FG,
            font=(FONT, 20, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            title_col,
            text="Map source columns  →  target columns  →  append to the chosen tab",
            bg=HEADER_BG,
            fg=HEADER_SUB_FG,
            font=(FONT, 11),
            anchor="w",
        ).pack(fill="x", pady=(2, 0))

        body = tk.Frame(self, bg=BG_MAIN)
        body.pack(fill="both", expand=True)
        tk.Frame(body, bg=BG_MAIN, height=8).pack(fill="x")

        notebook = ttk.Notebook(body)
        notebook.pack(fill="x", padx=14, pady=(0, 8))
        tab_files = tk.Frame(notebook, bg=BG_MAIN)
        tab_maps = tk.Frame(notebook, bg=BG_MAIN)
        tab_lookups = tk.Frame(notebook, bg=BG_MAIN)
        notebook.add(tab_files, text="  ①  Files  ")
        notebook.add(tab_maps, text="  ②  Column Mappings  ")
        notebook.add(tab_lookups, text="  ③  Lookups  ")

        # Source
        frame_src = self._card(tab_files, "Source File")

        row1 = ttk.Frame(frame_src)
        row1.pack(fill="x", padx=12, pady=(9, 6))
        ttk.Entry(row1, textvariable=self.source_path, width=55).pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        PillButton(
            row1, "Browse...", command=self.pick_source, bg_page=BG_CARD,
            fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE, fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 11, "bold"),
        ).pack(side="left")

        row2 = ttk.Frame(frame_src)
        row2.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Label(row2, text="Source tab:").pack(side="left")
        self.source_sheet_combo = ttk.Combobox(
            row2, textvariable=self.source_sheet, state="readonly", width=28
        )
        self.source_sheet_combo.pack(side="left", padx=8)
        self.source_sheet_combo.bind("<<ComboboxSelected>>", self._on_source_sheet_change)
        ttk.Checkbutton(
            row2, text="First row is a header (skip it)", variable=self.skip_header
        ).pack(side="left", padx=14)

        # Target

        frame_dst = self._card(tab_files, "Target Workbook")

        row3 = ttk.Frame(frame_dst)
        row3.pack(fill="x", padx=12, pady=(9, 6))
        ttk.Entry(row3, textvariable=self.target_path, width=55).pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        PillButton(
            row3, "Browse...", command=self.pick_target, bg_page=BG_CARD,
            fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE, fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 11, "bold"),
        ).pack(side="left")
        PillButton(
            row3, "New...", command=self.new_target, bg_page=BG_CARD,
            fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE, fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 11), padx=14, pady=11,
        ).pack(side="left", padx=(8, 0))

        row4 = ttk.Frame(frame_dst)
        row4.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Label(row4, text="Target tab:").pack(side="left")
        self.target_sheet_combo = ttk.Combobox(
            row4, textvariable=self.target_sheet, width=28
        )
        self.target_sheet_combo.pack(side="left", padx=8)
        ttk.Checkbutton(
            row4, text="First row is a header", variable=self.target_has_header
        ).pack(side="left", padx=10)
        self.target_tabs_hint = ttk.Label(
            row4, text="choose a target workbook first", style="Muted.TLabel"
        )
        self.target_tabs_hint.pack(side="left", padx=6)

        row5 = ttk.Frame(frame_dst)
        row5.pack(fill="x", padx=12, pady=(0, 4))
        ttk.Checkbutton(
            row5,
            text=f"Match the formatting of row {STYLE_TEMPLATE_ROW}",
            variable=self.copy_format,
        ).pack(side="left")
        ttk.Checkbutton(
            row5,
            text="Extend validation & rules",
            variable=self.extend_validation,
        ).pack(side="left", padx=16)
        ttk.Label(
            row5,
            text="make appended rows look and behave like that row",
            style="Muted.TLabel",
        ).pack(side="left", padx=8)

        row6 = ttk.Frame(frame_dst)
        row6.pack(fill="x", padx=12, pady=(0, 4))
        ttk.Checkbutton(
            row6, text="Keep a timestamped backup", variable=self.keep_backup
        ).pack(side="left")
        ttk.Label(
            row6,
            text="a failed write always rolls the file back either way",
            style="Muted.TLabel",
        ).pack(side="left", padx=8)

        ttk.Label(
            frame_dst,
            text="Browse edits the workbook in place; New… starts a fresh one.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=12, pady=(0, 8))

        # Mappings
        frame_map = self._card(tab_maps, "Column Mappings")

        id_row = tk.Frame(frame_map, bg=BG_CARD)
        id_row.pack(fill="x", padx=12, pady=(12, 6))
        ttk.Checkbutton(
            id_row, text="Auto ID", variable=self.id_enabled, command=self._on_id_toggle
        ).pack(side="left")
        ttk.Label(id_row, text="prefix").pack(side="left", padx=(12, 4))
        self.id_prefix_entry = ttk.Entry(id_row, textvariable=self.id_prefix, width=8)
        self.id_prefix_entry.pack(side="left")
        ttk.Label(id_row, text="→  to").pack(side="left", padx=(12, 4))
        self.id_col_combo = ttk.Combobox(
            id_row, textvariable=self.id_col, width=18,
            values=self.target_columns, state="readonly",
        )
        self.id_col_combo.pack(side="left", padx=4)
        ttk.Label(
            id_row,
            text=f"continues from the highest existing ID ({DEFAULT_ID_PREFIX}001, "
            f"{DEFAULT_ID_PREFIX}002, ...)",
            style="Muted.TLabel",
        ).pack(side="left", padx=10)

        ttk.Separator(frame_map, orient="horizontal").pack(fill="x", padx=12, pady=6)

        self.map_area = ScrollArea(frame_map, bg=BG_CARD, height=84)
        self.map_area.pack(fill="x", padx=12, pady=(0, 6))

        add_row = tk.Frame(frame_map, bg=BG_CARD)
        add_row.pack(fill="x", padx=12, pady=(0, 12))
        PillButton(
            add_row, "+  Add mapping", command=self.add_mapping,
            bg_page=BG_CARD, fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE,
            fill_disabled=SECONDARY_BG, fg=SECONDARY_FG, font=(FONT, 11, "bold"),
            padx=16, pady=8,
        ).pack(side="left")
        ttk.Label(
            add_row,
            text="Tick “extract after” to trim a value; leave it off to copy the cell whole.",
            style="Muted.TLabel",
        ).pack(side="left", padx=12)

        # Lookups
        frame_lk = self._card(tab_lookups, "Lookups")

        ttk.Label(
            frame_lk,
            text="Look a source value up in another workbook and write the matching "
            "value to the target - like VLOOKUP.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=12, pady=(10, 4))

        self.lookup_area = ScrollArea(frame_lk, bg=BG_CARD, height=196)
        self.lookup_area.pack(fill="x", padx=12, pady=(0, 6))

        lk_add = tk.Frame(frame_lk, bg=BG_CARD)
        lk_add.pack(fill="x", padx=12, pady=(0, 12))
        PillButton(
            lk_add, "+  Add lookup", command=self.add_lookup,
            bg_page=BG_CARD, fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE,
            fill_disabled=SECONDARY_BG, fg=SECONDARY_FG, font=(FONT, 11, "bold"),
            padx=16, pady=8,
        ).pack(side="left")
        ttk.Label(
            lk_add,
            text="Matching ignores case and surrounding spaces.",
            style="Muted.TLabel",
        ).pack(side="left", padx=12)

        # Actions
        frame_actions = tk.Frame(body, bg=BG_MAIN)
        frame_actions.pack(fill="x", padx=17, pady=(0, 9))
        PillButton(
            frame_actions, "1. Generate Preview", command=self.refresh_preview,
            bg_page=BG_MAIN, fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE,
            fill_disabled=SECONDARY_BG, fg=SECONDARY_FG, font=(FONT, 11, "bold"),
        ).pack(side="left", padx=(0, 12))
        self.append_btn = PillButton(
            frame_actions, "2. Append to Log", command=self.confirm_append,
            bg_page=BG_MAIN, fill=ACCENT, fill_active=ACCENT_ACTIVE,
            fill_disabled=ACCENT_DISABLED, fg="#ffffff", font=(FONT, 11, "bold"),
        )
        self.append_btn.pack(side="left")
        self.append_btn.set_state("disabled")
        ttk.Label(
            frame_actions,
            text="Append stays locked until the preview matches your current settings.",
            background=BG_MAIN,
            foreground=TEXT_MUTED,
            font=(FONT, 10),
        ).pack(side="left", padx=14)

        # Preview table
        frame_preview = self._card(body, "Preview", expand=True)

        tree_wrap = ttk.Frame(frame_preview)
        tree_wrap.pack(fill="both", expand=True, padx=12, pady=12)

        self.tree = ttk.Treeview(tree_wrap, columns=("placeholder",), show="headings", height=6)
        self.tree.pack(fill="both", expand=True, side="left")

        scroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.tag_configure("warn", background=WARN_BG)
        self.tree.tag_configure("even", background=ROW_ALT)
        self.tree.tag_configure("odd", background="#ffffff")

        self._on_id_toggle()
        self._rebuild_tree_columns()
        self.target_path.trace_add("write", self._on_target_path_change)

        for var in (
            self.source_path, self.source_sheet, self.skip_header,
            self.target_path, self.target_sheet, self.target_has_header,
            self.id_enabled, self.id_prefix, self.id_col,
        ):
            var.trace_add("write", self.schedule_refresh)

        for var in (self.source_path, self.source_sheet):
            var.trace_add("write", self._on_source_sheet_change)
        for var in (self.target_sheet, self.target_has_header):
            var.trace_add("write", self._on_target_sheet_change)

        # Status
        self.status_badge = StatusBadge(body, bg_page=BG_MAIN, font=(FONT, 11))
        self.status_badge.pack(padx=17, pady=(0, 10), anchor="w")
        self.status_badge.set(
            "Choose your files and mappings, then click Generate Preview.", TEXT_MUTED
        )

    # ---------- Mapping management ----------
    def add_mapping(self, source_col=None, target_col=None,
                    delim_enabled=False, delimiter=DEFAULT_DELIMITER):
        if source_col is None:
            nth = min(len(self.mappings), len(self.source_columns) - 1)
            source_col = self.source_columns[nth]
        if target_col is None:
            used = {
                idx for idx in (parse_column(m.target_col.get()) for m in self.mappings)
                if idx
            }
            if self.id_enabled.get():
                id_idx = parse_column(self.id_col.get())
                if id_idx:
                    used.add(id_idx)
            free = next(
                (i for i in range(1, len(self.target_columns) + 1) if i not in used), 1
            )
            target_col = self.target_columns[free - 1]

        row = MappingRow(
            self, self.map_area.inner, source_col, target_col, delim_enabled, delimiter
        )
        self.mappings.append(row)
        self.schedule_refresh()
        return row

    def remove_mapping(self, row):
        if row not in self.mappings:
            return
        self.mappings.remove(row)
        row.destroy()
        self.schedule_refresh()

    # ---------- Lookup management ----------
    def add_lookup(self):
        row = LookupRow(self, self.lookup_area.inner)
        self.lookups.append(row)
        self.schedule_refresh()
        return row

    def remove_lookup(self, row):
        if row not in self.lookups:
            return
        self.lookups.remove(row)
        row.destroy()
        self.schedule_refresh()

    def _lookup_table(self, path, sheet, key_idx, val_idx):
        """Build {normalised key: value} from a lookup sheet, cached by mtime.

        Keys are trimmed and lowercased. The first occurrence of a duplicate
        key wins, matching how VLOOKUP returns the first match.
        """
        try:
            cache_key = (path, sheet, key_idx, val_idx, os.path.getmtime(path))
        except OSError:
            return None
        if cache_key in self._lookup_cache:
            return self._lookup_cache[cache_key]

        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if sheet not in wb.sheetnames:
                wb.close()
                return None
            width = max(key_idx, val_idx)
            table = {}
            for raw in wb[sheet].iter_rows(min_col=1, max_col=width, values_only=True):
                key = raw[key_idx - 1] if key_idx - 1 < len(raw) else None
                if key in (None, ""):
                    continue
                norm = str(key).strip().lower()
                if norm in table:
                    continue
                val = raw[val_idx - 1] if val_idx - 1 < len(raw) else None
                table[norm] = "" if val in (None, "") else str(val).strip()
            wb.close()
        except Exception:
            return None

        self._lookup_cache[cache_key] = table
        return table

    def _on_id_toggle(self):
        state = "normal" if self.id_enabled.get() else "disabled"
        self.id_prefix_entry.configure(state=state)
        self.id_col_combo.configure(state="readonly" if self.id_enabled.get() else "disabled")
        self.schedule_refresh()

    def column_labels(self, path, sheet):
        """Build ['A - Server', 'B', ...] from row 1 of a sheet.

        Falls back to bare column letters if the file, tab, or row is
        unreadable, so the dropdowns always have usable values.
        """
        if not path or not sheet or not os.path.exists(path):
            return list(COLUMN_CHOICES)
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if sheet not in wb.sheetnames:
                wb.close()
                return list(COLUMN_CHOICES)
            first = next(wb[sheet].iter_rows(min_row=1, max_row=1, values_only=True), ())
            wb.close()
        except Exception:
            return list(COLUMN_CHOICES)

        labels = []
        for i, letter in enumerate(COLUMN_CHOICES):
            head = first[i] if i < len(first) else None
            head = str(head).strip() if head not in (None, "") else ""
            labels.append(f"{letter} - {head}" if head else letter)
        return labels

    @staticmethod
    def relabel(combo, var, labels):
        """Swap a column dropdown's labels, keeping it pointed at the same column."""
        current = parse_column(var.get())
        combo["values"] = labels
        if current and current <= len(labels):
            var.set(labels[current - 1])

    def _on_source_sheet_change(self, *_args):
        """Relabel the source column dropdowns with row-1 values as hints."""
        labels = self.column_labels(
            self.source_path.get().strip(), self.source_sheet.get().strip()
        )
        self.source_columns = labels
        for m in self.mappings:
            self.relabel(m.source_combo, m.source_col, labels)
        for lk in self.lookups:
            self.relabel(lk.source_combo, lk.source_col, labels)

    def _on_target_sheet_change(self, *_args):
        """Relabel the target column dropdowns from the target tab's header row."""
        if self.target_has_header.get():
            labels = self.column_labels(
                self.target_path.get().strip(), self.target_sheet.get().strip()
            )
        else:
            labels = list(COLUMN_CHOICES)

        self.target_columns = labels
        for m in self.mappings:
            self.relabel(m.target_combo, m.target_col, labels)
        for lk in self.lookups:
            self.relabel(lk.target_combo, lk.target_col, labels)
        self.relabel(self.id_col_combo, self.id_col, labels)

    # ---------- Preview table columns ----------
    def _rebuild_tree_columns(self):
        if not hasattr(self, "tree"):
            return

        def letter(value):
            idx = parse_column(value)
            return get_column_letter(idx) if idx else "?"

        specs = []
        if self.id_enabled.get():
            specs.append(
                ("__id", f"ID → {letter(self.id_col.get())}", 90, "center")
            )
        for i, m in enumerate(self.mappings):
            label = f"{letter(m.source_col.get())} → {letter(m.target_col.get())}"
            delim = m.effective_delimiter()
            if delim:
                label += f"  (after '{delim}')"
            specs.append((f"m{i}", label, 200, "w"))
        for i, lk in enumerate(self.lookups):
            label = (
                f"{letter(lk.source_col.get())} ⇒ "
                f"{letter(lk.target_col.get())}  (lookup)"
            )
            specs.append((f"l{i}", label, 200, "w"))
        specs.append(("__note", "Note", 170, "w"))

        self.tree.configure(columns=[s[0] for s in specs])
        for key, text, width, anchor in specs:
            self.tree.heading(key, text=text)
            self.tree.column(key, width=width, anchor=anchor)

        for item in self.tree.get_children():
            self.tree.delete(item)

    # ---------- File pickers ----------
    def pick_source(self):
        path = filedialog.askopenfilename(
            title="Select source Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm")],
        )
        if not path:
            return
        self.source_path.set(path)
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheets = wb.sheetnames
            self.source_sheet_combo["values"] = sheets
            if sheets:
                self.source_sheet.set(sheets[0])
            wb.close()
            self._on_source_sheet_change()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read workbook:\n{e}")

    def pick_target(self):
        """Open an existing workbook to append into (the normal case)."""
        path = filedialog.askopenfilename(
            title="Select the existing target workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm")],
        )
        if not path:
            return
        self.target_path.set(path)
        self._load_target_sheets(path)

    def new_target(self):
        """Create a brand-new target workbook (only when you explicitly want one)."""
        path = filedialog.asksaveasfilename(
            title="Create a new target workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        self.target_path.set(path)
        self._load_target_sheets(path)

    def _on_target_path_change(self, *_args):
        """Rescan tabs whenever the target path changes - typed, pasted or browsed."""
        if self._target_scan_job is not None:
            self.after_cancel(self._target_scan_job)
        self._target_scan_job = self.after(400, self._scan_target_sheets)

    def _scan_target_sheets(self):
        self._target_scan_job = None
        self._load_target_sheets(self.target_path.get().strip(), announce=False)

    def _load_target_sheets(self, path, announce=True):
        """Populate the target tab dropdown from the workbook, if it exists.

        `announce` controls whether read failures raise a dialog - browsing
        should report problems, background rescans should stay quiet.
        """
        if not path:
            self.target_sheet_combo["values"] = []
            self.target_tabs_hint.configure(text="choose a target workbook first")
            return

        if not os.path.exists(path):
            self.target_sheet_combo["values"] = []
            self.target_tabs_hint.configure(
                text="new file - type a tab name to create it"
            )
            return

        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheets = list(wb.sheetnames)
            wb.close()
        except Exception as e:
            self.target_sheet_combo["values"] = []
            self.target_tabs_hint.configure(text="could not read this workbook")
            if announce:
                messagebox.showerror("Error", f"Could not read target workbook:\n{e}")
            return

        self.target_sheet_combo["values"] = sheets
        self.target_tabs_hint.configure(
            text=f"{len(sheets)} tab(s) in this file - or type a new name"
        )
        if sheets and self.target_sheet.get().strip() not in sheets:
            self.target_sheet.set(sheets[0])
        self._on_target_sheet_change()

    # ---------- Validation ----------
    def _validate(self):
        """Return (error_message, plan) - error is None when valid.

        The plan is one entry per output column, in preview order: every
        column mapping first, then every lookup. Each entry carries what's
        needed to compute its value from a source row.
        """
        if not self.source_path.get().strip() or not self.source_sheet.get().strip():
            return "Pick a source file and tab first.", None
        if not self.target_path.get().strip():
            return "Pick a target workbook first.", None
        tab = self.target_sheet.get().strip()
        if not tab:
            return "Enter a target tab name.", None
        bad = set(tab) & set(INVALID_SHEET_CHARS)
        if bad:
            return (
                f"A tab name can't contain {' '.join(sorted(bad))} - "
                "Excel doesn't allow it.",
                None,
            )
        if len(tab) > MAX_SHEET_NAME:
            return f"Tab names are limited to {MAX_SHEET_NAME} characters.", None
        if not self.mappings and not self.lookups:
            return "Add at least one column mapping or lookup.", None

        plan = []
        seen_targets = {}

        if self.id_enabled.get():
            if not self.id_prefix.get().strip():
                return "The Auto ID prefix can't be blank.", None
            id_idx = parse_column(self.id_col.get())
            if not id_idx:
                return f"'{self.id_col.get()}' isn't a valid ID column.", None
            seen_targets[id_idx] = "the Auto ID"

        for m in self.mappings:
            src_idx = parse_column(m.source_col.get())
            tgt_idx = parse_column(m.target_col.get())
            if not src_idx:
                return f"'{m.source_col.get()}' isn't a valid source column.", None
            if not tgt_idx:
                return f"'{m.target_col.get()}' isn't a valid target column.", None
            if tgt_idx in seen_targets:
                return (
                    f"Target column {get_column_letter(tgt_idx)} is used by "
                    f"{seen_targets[tgt_idx]} and another mapping. "
                    "Each target column must be unique.",
                    None,
                )
            seen_targets[tgt_idx] = f"the {get_column_letter(src_idx)} mapping"
            delim = m.effective_delimiter()
            label = f"{get_column_letter(src_idx)} → {get_column_letter(tgt_idx)}"
            if delim:
                label += f"  (after '{delim}')"
            plan.append({
                "kind": "map", "src": src_idx, "tgt": tgt_idx,
                "delim": delim, "label": label,
            })

        for n, lk in enumerate(self.lookups, start=1):
            path = lk.path.get().strip()
            sheet = lk.sheet.get().strip()
            if not path:
                return f"Lookup {n}: pick a lookup file.", None
            if not os.path.exists(path):
                return f"Lookup {n}: that lookup file doesn't exist.", None
            if not sheet:
                return f"Lookup {n}: pick a tab in the lookup file.", None

            src_idx = parse_column(lk.source_col.get())
            key_idx = parse_column(lk.key_col.get())
            val_idx = parse_column(lk.val_col.get())
            tgt_idx = parse_column(lk.target_col.get())
            if not src_idx:
                return f"Lookup {n}: '{lk.source_col.get()}' isn't a valid source column.", None
            if not key_idx:
                return f"Lookup {n}: '{lk.key_col.get()}' isn't a valid match column.", None
            if not val_idx:
                return f"Lookup {n}: '{lk.val_col.get()}' isn't a valid return column.", None
            if not tgt_idx:
                return f"Lookup {n}: '{lk.target_col.get()}' isn't a valid target column.", None
            if tgt_idx in seen_targets:
                return (
                    f"Target column {get_column_letter(tgt_idx)} is used by "
                    f"{seen_targets[tgt_idx]} and lookup {n}. "
                    "Each target column must be unique.",
                    None,
                )

            table = self._lookup_table(path, sheet, key_idx, val_idx)
            if table is None:
                return f"Lookup {n}: couldn't read '{sheet}' in the lookup file.", None
            if not table:
                return f"Lookup {n}: no usable rows in '{sheet}'.", None

            seen_targets[tgt_idx] = f"lookup {n}"
            plan.append({
                "kind": "lookup", "src": src_idx, "tgt": tgt_idx,
                "table": table, "fallback": lk.fallback.get().strip(), "n": n,
                "label": f"{get_column_letter(src_idx)} ⇒ {get_column_letter(tgt_idx)}  (lookup)",
            })

        return None, plan

    # ---------- Source / target reads (cached) ----------
    def _source_rows(self):
        """All source rows, cached until the file, tab, or mtime changes."""
        path = self.source_path.get().strip()
        sheet = self.source_sheet.get().strip()
        if not path or not sheet or not os.path.exists(path):
            return None
        try:
            key = (path, sheet, os.path.getmtime(path))
        except OSError:
            return None
        if self._src_cache_key != key:
            try:
                wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                if sheet not in wb.sheetnames:
                    wb.close()
                    return None
                rows = [tuple(r) for r in wb[sheet].iter_rows(values_only=True)]
                wb.close()
            except Exception:
                return None
            self._src_cache_key = key
            self._src_cache_rows = rows
        return self._src_cache_rows

    def _next_id_start(self):
        """Next free ID number in the target tab, cached by file mtime."""
        path = self.target_path.get().strip()
        sheet = self.target_sheet.get().strip()
        prefix = self.id_prefix.get().strip()
        id_idx = parse_column(self.id_col.get())
        if not (path and sheet and prefix and id_idx and os.path.exists(path)):
            return 1
        try:
            key = (path, sheet, prefix, id_idx, os.path.getmtime(path))
        except OSError:
            return 1
        if self._tgt_cache_key != key:
            start = 1
            try:
                wb = openpyxl.load_workbook(path, data_only=True)
                if sheet in wb.sheetnames:
                    start = next_id_number(wb[sheet], prefix, id_idx)
                wb.close()
            except Exception:
                start = 1
            self._tgt_cache_key = key
            self._tgt_cache_nextid = start
        return self._tgt_cache_nextid

    def _invalidate_caches(self):
        self._src_cache_key = None
        self._tgt_cache_key = None
        self._lookup_cache.clear()

    # ---------- Live preview ----------
    def schedule_refresh(self, *_args):
        """Mark the preview out of date rather than rebuilding it immediately.

        Rebuilding on every edit is slow on large sheets, so the preview is
        generated on demand. Append stays disabled while the shown preview
        doesn't match the current settings, which is what stops the two from
        silently diverging.
        """
        if self._preview_stale:
            return
        self._mark_stale("Settings changed - click Generate Preview.")

    def _mark_stale(self, message, color=WARNTEXT):
        """Drop the shown preview and lock Append until it's regenerated."""
        self._preview_stale = True
        self._preview_rows = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.append_btn.set_state("disabled")
        self.status_badge.set(message, color)

    def refresh_preview(self):
        self._refresh_job = None
        self._preview_stale = False
        self._rebuild_tree_columns()
        self._preview_rows = []

        error, plan = self._validate()
        if error:
            self.status_badge.set(error, WARNTEXT)
            self.append_btn.set_state("disabled")
            return

        rows = self._source_rows()
        if rows is None:
            self.status_badge.set("Could not read the source tab.", ERROR)
            self.append_btn.set_state("disabled")
            return

        if self.skip_header.get() and rows:
            rows = rows[1:]

        start_n = self._next_id_start() if self.id_enabled.get() else 1
        max_src = max(step["src"] for step in plan)
        prefix = self.id_prefix.get().strip()
        self._preview_rows = []
        n = start_n
        for raw in rows:
            cells = [raw[i] if i < len(raw) else None for i in range(max_src)]
            if all(c in (None, "") for c in cells):
                continue  # fully blank row

            values = []
            notes = []
            for step in plan:
                raw_val = cells[step["src"] - 1]
                letter = get_column_letter(step["src"])

                if step["kind"] == "map":
                    value = extract_after_delimiter(raw_val, step["delim"])
                    if raw_val in (None, ""):
                        notes.append(f"{letter} empty")
                    elif step["delim"] and not value:
                        notes.append(f"{letter}: '{step['delim']}' not found")
                else:
                    if raw_val in (None, ""):
                        value = step["fallback"]
                        notes.append(f"{letter} empty")
                    else:
                        hit = step["table"].get(str(raw_val).strip().lower())
                        if hit is None:
                            value = step["fallback"]
                            notes.append(f"lookup {step['n']}: no match for '{raw_val}'")
                        else:
                            value = hit
                values.append(value)

            id_str = f"{prefix}{n:0{ID_PAD}d}" if self.id_enabled.get() else None
            self._preview_rows.append((id_str, values, "; ".join(notes)))
            n += 1

        for i, (id_str, values, note) in enumerate(self._preview_rows):
            display = ([id_str] if self.id_enabled.get() else []) + values + [note]
            if note:
                tags = ("warn",)
            else:
                tags = ("even",) if i % 2 == 0 else ("odd",)
            self.tree.insert("", "end", values=display, tags=tags)

        if not self._preview_rows:
            self.status_badge.set("No data rows found to map.", ERROR)
            self.append_btn.set_state("disabled")
            return

        warn_count = sum(1 for r in self._preview_rows if r[2])
        id_note = ""
        if self.id_enabled.get():
            id_note = f"IDs {self._preview_rows[0][0]}-{self._preview_rows[-1][0]}. "
        n_maps = sum(1 for s in plan if s["kind"] == "map")
        n_lookups = len(plan) - n_maps
        what = f"{n_maps} mapping(s)"
        if n_lookups:
            what += f" + {n_lookups} lookup(s)"
        self.status_badge.set(
            (
                f"{len(self._preview_rows)} row(s) ready across {what}. "
                f"{id_note}{warn_count} row(s) flagged."
            ),
            (WARNTEXT if warn_count else SUCCESS),
        )
        self.append_btn.set_state("normal")

    def confirm_append(self):
        # Recompute from current settings so what gets written always matches
        # what's on screen, even if the workbooks changed underneath us.
        self._invalidate_caches()
        self.refresh_preview()

        if not self._preview_rows:
            messagebox.showwarning(
                "Nothing to append", "There are no rows to write - check the preview."
            )
            return

        error, plan = self._validate()
        if error:
            messagebox.showwarning("Check the settings", error)
            return

        tgt_path = self.target_path.get().strip()
        tgt_sheet = self.target_sheet.get().strip()

        # Writing a large sheet takes seconds and blocks the event loop. Show
        # what's happening and keep the window responsive, or macOS flags the
        # app as not responding and the user force-quits mid-write.
        total = len(self._preview_rows)
        self.append_btn.set_state("disabled")
        self.status_badge.set(f"Opening the target workbook ({total} row(s))...", WARNTEXT)
        self.update()

        # Keep an untouched copy so a bad write can be rolled back rather than
        # destroying the only version of the workbook.
        rollback = None
        if os.path.exists(tgt_path):
            rollback = f"{tgt_path}.rollback"
            try:
                shutil.copy2(tgt_path, rollback)
            except OSError:
                rollback = None

        try:
            # Must be read before openpyxl opens the workbook, since it drops
            # x14 validation extensions silently.
            saved_extensions = read_validation_extensions(tgt_path)

            if os.path.exists(tgt_path):
                # Open the existing workbook and edit it in place - never replace it.
                # keep_vba preserves macros in .xlsm, which openpyxl drops otherwise.
                is_macro = os.path.splitext(tgt_path)[1].lower() == ".xlsm"
                tgt_wb = openpyxl.load_workbook(tgt_path, keep_vba=is_macro)
            else:
                tgt_wb = openpyxl.Workbook()
                default_sheet = tgt_wb.active
                if default_sheet.title != tgt_sheet:
                    tgt_wb.remove(default_sheet)

            if tgt_sheet in tgt_wb.sheetnames:
                ws = tgt_wb[tgt_sheet]
            else:
                ws = tgt_wb.create_sheet(tgt_sheet)
                if self.id_enabled.get():
                    ws.cell(row=1, column=parse_column(self.id_col.get()), value="ID")
                for step in plan:
                    heading = (
                        f"Column {get_column_letter(step['src'])}"
                        if step["kind"] == "map"
                        else f"Lookup {step['n']}"
                    )
                    ws.cell(row=1, column=step["tgt"], value=heading)

            start_row = last_used_row(ws) + 1
            start_row = max(start_row, 2)  # never write into the header row

            # Snapshot before writing, while row 2 is still the last styled
            # row we know about.
            template = row_style_template(ws) if self.copy_format.get() else None

            id_idx = parse_column(self.id_col.get()) if self.id_enabled.get() else None

            for i, (id_str, values, _note) in enumerate(self._preview_rows):
                r = start_row + i
                apply_row_style(ws, r, template)
                if id_idx:
                    ws.cell(row=r, column=id_idx, value=id_str)
                for step, value in zip(plan, values):
                    ws.cell(row=r, column=step["tgt"], value=value)
                if i and i % PROGRESS_EVERY == 0:
                    self.status_badge.set(f"Writing row {i:,} of {total:,}...", WARNTEXT)
                    self.update()

            if self.extend_validation.get() and total:
                last_row = start_row + total - 1
                extend_data_validations(ws, last_row)
                extend_conditional_formatting(ws, last_row)

            self.status_badge.set(f"Saving {total:,} row(s)...", WARNTEXT)
            self.update()
            tgt_wb.save(tgt_path)

            stretch = None
            if self.extend_validation.get() and total:
                stretch = (tgt_sheet, start_row + total - 1)
            restore_validation_extensions(tgt_path, saved_extensions, stretch)

            self.status_badge.set("Verifying the saved workbook...", WARNTEXT)
            self.update()
            if not workbook_is_readable(tgt_path):
                raise RuntimeError(
                    "the saved workbook failed its integrity check, so your "
                    "original file has been put back unchanged"
                )

            if self.keep_backup.get() and rollback:
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                base, ext = os.path.splitext(tgt_path)
                shutil.copy2(rollback, f"{base}.backup-{stamp}{ext}")

            messagebox.showinfo(
                "Done", f"Appended {total:,} row(s) to '{tgt_sheet}'."
            )
            self._invalidate_caches()
            self._mark_stale(
                f"Appended {total:,} row(s) to '{tgt_sheet}'. "
                "Click Generate Preview to continue.",
                SUCCESS,
            )

        except PermissionError:
            self._restore_rollback(rollback)
            rollback = None
            self._mark_stale(
                "Target file is locked - close it in Excel and try again.", ERROR
            )
            messagebox.showerror(
                "File is locked",
                "Could not save the target workbook - it looks like it's open in "
                "Excel.\n\nClose the file there and try again. Nothing has been "
                "written yet.",
            )
        except Exception as e:
            restored = self._restore_rollback(rollback)
            rollback = None
            self._mark_stale("Append failed - your file was put back.", ERROR)
            messagebox.showerror(
                "Error",
                f"Could not append to target workbook:\n{e}\n\n"
                + (
                    "Your original workbook has been restored unchanged."
                    if restored
                    else "No backup was available, so check the file before reusing it."
                ),
            )
        finally:
            if rollback and os.path.exists(rollback):
                os.remove(rollback)

    @staticmethod
    def _restore_rollback(rollback):
        """Put the untouched copy back after a failed write."""
        if not rollback or not os.path.exists(rollback):
            return False
        try:
            shutil.move(rollback, rollback[: -len(".rollback")])
            return True
        except OSError:
            return False


def main():
    ColumnMapperApp().mainloop()


if __name__ == "__main__":
    main()

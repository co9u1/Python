"""
Fileshare Access Auditor
------------------------
Reads raw user-access data from a source Excel file (column A = server,
column B = Access Path), extracts the pieces needed for the audit log,
and appends formatted rows into the "2c.Fileshares" tab of a target
Excel workbook.

Extraction rules:
  - Column A (server, full value)              -> target column E
  - Substring of column B after "/vol/" (or "\\vol\\") -> target column F

Target rows are appended (never overwritten), with an auto-incrementing
ID in target column A, formatted as FS001, FS002, ... continuing from
whatever IDs already exist in that column.
"""

import os
import re

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

import openpyxl

TARGET_SHEET_NAME = "2c.Fileshares"
ID_PREFIX = "FS"

# Matches "/vol/" or "\vol\" (case-insensitive), capturing everything after it.
VOL_PATTERN = re.compile(r"[\\/]vol[\\/](.*)$", re.IGNORECASE)


def extract_after_vol(access_path: str) -> str:
    """Return the substring after '/vol/' (or '\\vol\\'), or '' if not found."""
    if not access_path:
        return ""
    match = VOL_PATTERN.search(str(access_path))
    return match.group(1).strip() if match else ""


def next_id_number(ws) -> int:
    """Scan target column A for existing FS### ids and return the next number."""
    max_n = 0
    for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
        val = row[0]
        if not val:
            continue
        m = re.match(rf"^{ID_PREFIX}(\d+)$", str(val).strip(), re.IGNORECASE)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def last_used_row(ws) -> int:
    """Return the last row index (1-based) that has any content, 0 if empty/header-only."""
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
        self._f = tkfont.Font(family=font[0], size=font[1])
        self.set("", SUCCESS)

    def set(self, text, color):
        self.delete("all")
        if not text:
            self.configure(width=1, height=1)
            return
        pad_x, pad_y, dot_r, gap = 14, 8, 4, 8
        text_w = self._f.measure(text)
        text_h = self._f.metrics("linespace")
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


class FileshareAuditorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Fileshare Access Auditor")
        self.geometry("860x640")
        self.minsize(760, 560)
        self.configure(bg=BG_MAIN)

        self.source_path = tk.StringVar()
        self.source_sheet = tk.StringVar()
        self.skip_header = tk.BooleanVar(value=True)
        self.target_path = tk.StringVar()

        self._preview_rows = []  # list of (id_str, server, extracted, warning)

        self._build_style()
        self._build_ui()

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
            "Primary.TButton",
            background=ACCENT,
            foreground="white",
            font=(FONT, 11, "bold"),
            padding=(16, 9),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("disabled", ACCENT_DISABLED), ("active", ACCENT_ACTIVE)],
            foreground=[("disabled", "#eef2fb")],
        )

        style.configure(
            "Secondary.TButton",
            background=SECONDARY_BG,
            foreground=TEXT_DARK,
            font=(FONT, 11),
            padding=(14, 8),
            borderwidth=0,
        )
        style.map("Secondary.TButton", background=[("active", SECONDARY_ACTIVE)])

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
        wrap.pack(fill="both" if expand else "x", expand=expand, padx=14, pady=(0, 14))
        card = ttk.LabelFrame(wrap, text=title, style="Card.TLabelframe")
        card.pack(fill="both" if expand else "x", expand=expand, padx=(0, 3), pady=(0, 3))
        return card

    def _build_ui(self):
        # Header banner
        header = tk.Frame(self, bg=HEADER_BG)
        header.pack(fill="x")

        header_inner = tk.Frame(header, bg=HEADER_BG)
        header_inner.pack(fill="x", padx=18, pady=16)

        CircleBadge(header_inner, "🗂️", diameter=48, bg_page=HEADER_BG, fill=ACCENT).pack(
            side="left", padx=(0, 14)
        )
        title_col = tk.Frame(header_inner, bg=HEADER_BG)
        title_col.pack(side="left", fill="x", expand=True)
        tk.Label(
            title_col,
            text="Fileshare Access Auditor",
            bg=HEADER_BG,
            fg=HEADER_FG,
            font=(FONT, 20, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            title_col,
            text="Extract server & share-path data  →  format  →  append to the audit log",
            bg=HEADER_BG,
            fg=HEADER_SUB_FG,
            font=(FONT, 11),
            anchor="w",
        ).pack(fill="x", pady=(2, 0))

        body = tk.Frame(self, bg=BG_MAIN)
        body.pack(fill="both", expand=True)
        tk.Frame(body, bg=BG_MAIN, height=16).pack(fill="x")  # top breathing room

        # Source
        frame_src = self._card(body, "①  Source File — raw user-access data")

        row1 = ttk.Frame(frame_src)
        row1.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Entry(row1, textvariable=self.source_path, width=55).pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        PillButton(
            row1, "Browse...", command=self.pick_source, bg_page=BG_CARD,
            fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE, fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 11, "bold"),
        ).pack(side="left")

        row2 = ttk.Frame(frame_src)
        row2.pack(fill="x", padx=12, pady=(0, 8))
        ttk.Label(row2, text="Source sheet:").pack(side="left")
        self.source_sheet_combo = ttk.Combobox(
            row2, textvariable=self.source_sheet, state="readonly", width=28
        )
        self.source_sheet_combo.pack(side="left", padx=8)
        ttk.Checkbutton(
            row2, text="First row is a header (skip it)", variable=self.skip_header
        ).pack(side="left", padx=14)

        ttk.Label(
            frame_src,
            text="Reads column A (server) and column B (Access Path) by position.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=12, pady=(0, 12))

        # Target
        frame_dst = self._card(body, "②  Target Audit Log Workbook")

        row3 = ttk.Frame(frame_dst)
        row3.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Entry(row3, textvariable=self.target_path, width=55).pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        PillButton(
            row3, "Browse...", command=self.pick_target, bg_page=BG_CARD,
            fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE, fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 11, "bold"),
        ).pack(side="left")

        ttk.Label(
            frame_dst,
            text=f"Appends to sheet '{TARGET_SHEET_NAME}' (created automatically if missing). "
            "ID goes in column A, server in column E, extracted path in column F.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=12, pady=(0, 12))

        # Actions
        frame_actions = tk.Frame(body, bg=BG_MAIN)
        frame_actions.pack(fill="x", padx=17, pady=(0, 14))
        PillButton(
            frame_actions, "1. Preview Extraction", command=self.preview, bg_page=BG_MAIN,
            fill=SECONDARY_BG, fill_active=SECONDARY_ACTIVE, fill_disabled=SECONDARY_BG,
            fg=SECONDARY_FG, font=(FONT, 11, "bold"),
        ).pack(side="left", padx=(0, 12))
        self.confirm_btn = PillButton(
            frame_actions, "2. Confirm & Append to Log", command=self.confirm_append,
            bg_page=BG_MAIN, fill=ACCENT, fill_active=ACCENT_ACTIVE,
            fill_disabled=ACCENT_DISABLED, fg="#ffffff", font=(FONT, 11, "bold"),
        )
        self.confirm_btn.pack(side="left")
        self.confirm_btn.set_state("disabled")

        # Preview table
        frame_preview = self._card(body, "Preview", expand=True)

        tree_wrap = ttk.Frame(frame_preview)
        tree_wrap.pack(fill="both", expand=True, padx=12, pady=12)

        columns = ("id", "server", "extracted", "note")
        self.tree = ttk.Treeview(
            tree_wrap, columns=columns, show="headings", height=14
        )
        self.tree.heading("id", text="ID")
        self.tree.heading("server", text="Server  →  col E")
        self.tree.heading("extracted", text="After /vol/  →  col F")
        self.tree.heading("note", text="Note")
        self.tree.column("id", width=70, anchor="center")
        self.tree.column("server", width=220)
        self.tree.column("extracted", width=280)
        self.tree.column("note", width=150)
        self.tree.pack(fill="both", expand=True, side="left")

        scroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.tag_configure("warn", background=WARN_BG)
        self.tree.tag_configure("even", background=ROW_ALT)
        self.tree.tag_configure("odd", background="#ffffff")

        # Status
        self.status_badge = StatusBadge(body, bg_page=BG_MAIN, font=(FONT, 11))
        self.status_badge.pack(padx=17, pady=(0, 16), anchor="w")

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
        except Exception as e:
            messagebox.showerror("Error", f"Could not read workbook:\n{e}")

    def pick_target(self):
        path = filedialog.asksaveasfilename(
            title="Select or create the target audit log workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        self.target_path.set(path)

    # ---------- Core logic ----------
    def preview(self):
        src_path = self.source_path.get().strip()
        sheet_name = self.source_sheet.get().strip()
        tgt_path = self.target_path.get().strip()

        if not src_path or not sheet_name:
            messagebox.showwarning("Missing info", "Pick a source file and sheet first.")
            return
        if not tgt_path:
            messagebox.showwarning("Missing info", "Pick a target log workbook first.")
            return

        try:
            src_wb = openpyxl.load_workbook(src_path, data_only=True)
            src_ws = src_wb[sheet_name]
        except Exception as e:
            messagebox.showerror("Error", f"Could not read source sheet:\n{e}")
            return

        # Determine starting ID number from the target workbook, if it exists.
        start_n = 1
        if os.path.exists(tgt_path):
            try:
                tgt_wb = openpyxl.load_workbook(tgt_path, data_only=True)
                if TARGET_SHEET_NAME in tgt_wb.sheetnames:
                    start_n = next_id_number(tgt_wb[TARGET_SHEET_NAME])
                tgt_wb.close()
            except Exception as e:
                messagebox.showerror("Error", f"Could not read target workbook:\n{e}")
                return

        rows = list(src_ws.iter_rows(min_col=1, max_col=2, values_only=True))
        if self.skip_header.get() and rows:
            rows = rows[1:]

        self._preview_rows = []
        n = start_n
        for server, access_path in rows:
            if server in (None, "") and access_path in (None, ""):
                continue  # fully blank row, skip silently
            server_val = str(server).strip() if server not in (None, "") else ""
            extracted = extract_after_vol(access_path)
            note = ""
            if not server_val:
                note = "missing server"
            elif not extracted:
                note = "'/vol/' not found"
            id_str = f"{ID_PREFIX}{n:03d}"
            self._preview_rows.append((id_str, server_val, extracted, note))
            n += 1

        # Populate the tree
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, row in enumerate(self._preview_rows):
            if row[3]:
                tags = ("warn",)
            else:
                tags = ("even",) if i % 2 == 0 else ("odd",)
            self.tree.insert("", "end", values=row, tags=tags)

        if not self._preview_rows:
            self.status_badge.set("No data rows found to extract.", ERROR)
            self.confirm_btn.set_state("disabled")
            return

        warn_count = sum(1 for r in self._preview_rows if r[3])
        self.status_badge.set(
            (
                f"Previewing {len(self._preview_rows)} row(s), IDs "
                f"{self._preview_rows[0][0]}-{self._preview_rows[-1][0]}. "
                f"{warn_count} row(s) flagged. Review, then click Confirm to append."
            ),
            (WARNTEXT if warn_count else SUCCESS),
        )
        self.confirm_btn.set_state("normal")

    def confirm_append(self):
        if not self._preview_rows:
            messagebox.showwarning("Nothing to append", "Run Preview first.")
            return

        tgt_path = self.target_path.get().strip()
        try:
            if os.path.exists(tgt_path):
                tgt_wb = openpyxl.load_workbook(tgt_path)
            else:
                tgt_wb = openpyxl.Workbook()
                default_sheet = tgt_wb.active
                if default_sheet.title != TARGET_SHEET_NAME:
                    tgt_wb.remove(default_sheet)

            if TARGET_SHEET_NAME in tgt_wb.sheetnames:
                ws = tgt_wb[TARGET_SHEET_NAME]
            else:
                ws = tgt_wb.create_sheet(TARGET_SHEET_NAME)
                ws.cell(row=1, column=1, value="ID")
                ws.cell(row=1, column=5, value="Server")
                ws.cell(row=1, column=6, value="Path (after /vol/)")

            start_row = last_used_row(ws) + 1
            if start_row < 2:
                start_row = 2  # never write into the header row

            for i, (id_str, server_val, extracted, _note) in enumerate(self._preview_rows):
                r = start_row + i
                ws.cell(row=r, column=1, value=id_str)   # A: ID
                ws.cell(row=r, column=5, value=server_val)  # E: server
                ws.cell(row=r, column=6, value=extracted)   # F: after /vol/

            tgt_wb.save(tgt_path)

            self.status_badge.set(
                (
                    f"Appended {len(self._preview_rows)} row(s) to "
                    f"'{TARGET_SHEET_NAME}' in {tgt_path}."
                ),
                SUCCESS,
            )
            messagebox.showinfo(
                "Done", f"Appended {len(self._preview_rows)} row(s) to the audit log."
            )
            self._preview_rows = []
            self.confirm_btn.set_state("disabled")

        except Exception as e:
            messagebox.showerror("Error", f"Could not append to target workbook:\n{e}")


if __name__ == "__main__":
    app = FileshareAuditorApp()
    app.mainloop()

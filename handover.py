# -*- coding: utf-8 -*-
"""Ứng dụng quản lý sổ bàn giao Excel."""
import json
import sys
import tkinter as tk
import unicodedata
from copy import copy
from datetime import date, datetime
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk
from tkinter.font import Font as TkFont, families

import customtkinter as ctk
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.table import Table, TableStyleInfo

APP_NAME = "Handover Workspace"
CONFIG = Path.home() / ".handover_workspace.json"
PAGE_SIZE = 200
MAX_EXCEL_ROW, MAX_EXCEL_COL = 1_048_576, 16_384
PAGE_BG = ("#EEF2F7", "#0D1523")
SURFACE = ("#FFFFFF", "#151F30")
BORDER = ("#DCE4EF", "#2A3A50")
TEXT = ("#14243B", "#EEF4FC")
MUTED = ("#50627C", "#A6B7CB")
PRIMARY, PRIMARY_HOVER = "#244DC6", "#1C3EA7"
BUTTON_BG = ("#EEF3FB", "#223149")
BUTTON_HOVER = ("#DFE9F8", "#30445E")
NAV, NAV_HOVER = "#14223A", "#263B57"

RECEIVER_HEADERS = (
    "receiver", "nguoi nhan", "receiver nguoi nhan", "nguoi nhan receiver", "receiver code", "receiver id",
    "ma nguoi nhan", "receiver name", "ten nguoi nhan",
)
DEFAULT_FILL_HEADERS = {
    "nguoi phu trach", "person in charge", "assignee", "department", "dept",
    "bo phan", "ma bo phan", "phong ban", "don vi", "team", "to", "khu vuc",
    "area", "location", "vi tri", "storage location", "vi tri luu kho",
    "cost center", "cost centre", "trung tam chi phi", "email", "receiver email",
    "phone", "phone number", "receiver phone", "so dien thoai", "dien thoai",
    "receiver name", "ten nguoi nhan", "receiver id", "receiver code", "ma nguoi nhan",
}


def header_key(value):
    text = unicodedata.normalize("NFKD", str(value).casefold()).replace("đ", "d")
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join("".join(char if char.isalnum() else " " for char in text).split())


def has_value(value):
    return value is not None and (not isinstance(value, str) or bool(value.strip()))


def sheet_headers(ws):
    # Reading the sparse cells avoids creating blank cells or expanding Excel's used range.
    headers = {
        col: cell for (row, col), cell in ws._cells.items()
        if row == 1 and has_value(cell.value)
    }
    return dict(sorted(headers.items()))


def receiver_column(headers):
    columns_by_header = {}
    for col, cell in headers.items():
        columns_by_header.setdefault(header_key(cell.value), col)
    return next((columns_by_header[name] for name in RECEIVER_HEADERS if name in columns_by_header), None)


def receiver_key(cell):
    if cell is None or cell.data_type in ("f", "e") or not has_value(cell.value):
        return ""
    value = cell.value
    if type(value) is int or (type(value) is float and value.is_integer()):
        value = int(value)
        if value >= 0 and set(cell.number_format) == {"0"}:
            value = str(value).zfill(len(cell.number_format))
    return " ".join(unicodedata.normalize("NFKC", str(value)).casefold().split())


def fill_receiver_rows(ws, rows, receiver_col, target_cols):
    targets = set(rows) - {1}
    if not targets or not receiver_col or not target_cols:
        return 0, 0

    last_target = max(targets)
    receiver_cells = sorted(
        (row, cell) for (row, col), cell in ws._cells.items()
        if col == receiver_col and 1 < row <= last_target
    )
    history = {}
    filled_rows = filled_cells = 0
    for row, receiver in receiver_cells:
        key = receiver_key(receiver)
        if not key:
            continue
        previous_values = history.setdefault(key, {})
        row_count = 0
        for col in target_cols:
            cell = ws._cells.get((row, col))
            if cell is not None and has_value(cell.value):
                if cell.data_type not in ("f", "e"):
                    previous_values[col] = (cell.value, cell.data_type)
            elif row in targets and col in previous_values and not isinstance(cell, MergedCell):
                value, data_type = previous_values[col]
                destination = ws.cell(row, col)
                destination.value = value
                destination.data_type = data_type
                row_count += 1
        if row_count:
            filled_rows += 1
            filled_cells += row_count
    return filled_rows, filled_cells


class ModernHandover(ctk.CTk):
    def __init__(self, initial_file=None):
        super().__init__()
        self.title(APP_NAME)
        width = max(960, min(1380, self.winfo_screenwidth() - 60))
        height = max(580, min(820, self.winfo_screenheight() - 95))
        self.geometry(f"{width}x{height}")
        self.minsize(960, 580)
        self.protocol("WM_DELETE_WINDOW", self.close_app)
        self.wb = None
        self.path = None
        self.dirty = False
        self.page = 0
        self.page_count = 1
        self.active_cell = (1, 1)
        self.editor = None
        self.edit_cell = None
        self.fill_settings = {}
        self.receiver_only = tk.BooleanVar()
        self.auto_fill = tk.BooleanVar(value=True)
        self.formula_var = tk.StringVar()
        self.cell_var = tk.StringVar(value="A1")
        self.status_var = tk.StringVar(value="Sẵn sàng")
        self.file_var = tk.StringVar(value="Sổ mới chưa lưu")
        self.fill_var = tk.StringVar()
        self.page_var = tk.StringVar()
        self.result_var = tk.StringVar()

        config = self.read_config()
        ctk.set_default_color_theme("blue")
        ctk.set_appearance_mode(config.get("theme", "System"))
        available_fonts = set(families(self))
        self.grid_font = next((font for font in ("Segoe UI", "DejaVu Sans") if font in available_fonts), "TkDefaultFont")
        self.heading_font = TkFont(self, family=self.grid_font, size=10, weight="bold")
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.apply_tree_theme()
        self.build_ui()
        self.create_new(skip_confirm=True)
        candidate = initial_file or config.get("last_file")
        if candidate and Path(candidate).is_file():
            self.open_path(candidate, silent=True)
        self.after(300000, self.autosave)

    @property
    def ws(self):
        return self.wb.active if self.wb else None

    def apply_tree_theme(self):
        dark = ctk.get_appearance_mode() == "Dark"
        bg, fg, border = (color[dark] for color in (SURFACE, TEXT, BORDER))
        heading = "#213049" if dark else "#E5EDF8"
        self.style.configure(
            "Modern.Treeview", background=bg, fieldbackground=bg, foreground=fg,
            rowheight=31, borderwidth=0, relief="flat", font=(self.grid_font, 10),
        )
        self.style.configure(
            "Modern.Treeview.Heading", background=heading, foreground=fg,
            relief="flat", borderwidth=0, padding=(8, 8), font=self.heading_font,
        )
        self.style.map("Modern.Treeview", background=[("selected", PRIMARY)], foreground=[("selected", "#FFFFFF")])
        self.style.map("Modern.Treeview.Heading", background=[("active", BUTTON_HOVER[dark])])
        self.style.configure("Handover.TEntry", fieldbackground=bg, foreground=fg, insertcolor=fg, padding=3)
        for orientation in ("Vertical", "Horizontal"):
            self.style.configure(
                f"{orientation}.TScrollbar", gripcount=0, background=border,
                troughcolor=bg, bordercolor=bg, lightcolor=border, darkcolor=border,
                borderwidth=0, arrowsize=11,
            )

    def build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.build_sidebar()
        self.content = ctk.CTkFrame(self, corner_radius=0, fg_color=PAGE_BG)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(4, weight=1)
        self.build_header()
        self.build_metrics()
        self.build_toolbar()
        self.build_formula_bar()
        self.build_grid()
        self.build_statusbar()
        self.bind_all("<Control-n>", lambda e: self.create_new())
        self.bind_all("<Control-o>", lambda e: self.open_file())
        self.bind_all("<Control-s>", lambda e: self.save())
        self.bind_all("<Control-Shift-S>", lambda e: self.save_as())
        self.bind_all("<Control-f>", lambda e: self.search_entry.focus_set())

    def button(self, parent, text, command, width, primary=False, **kwargs):
        return ctk.CTkButton(
            parent, text=text, command=command, width=width, height=32, corner_radius=8,
            fg_color=PRIMARY if primary else BUTTON_BG,
            hover_color=PRIMARY_HOVER if primary else BUTTON_HOVER,
            text_color="#FFFFFF" if primary else TEXT, **kwargs,
        )

    def build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=192, corner_radius=0, fg_color=NAV)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)
        sidebar.grid_rowconfigure(4, weight=1)
        ctk.CTkLabel(
            sidebar, text="Handover", text_color="white", font=ctk.CTkFont(size=21, weight="bold"),
        ).grid(row=0, column=0, padx=16, pady=(19, 15), sticky="w")

        actions = [
            ("＋  Tạo sổ mới", self.create_new),
            ("Mở tệp Excel", self.open_file),
            ("Lưu thành tệp mới", self.save_as),
        ]
        for row, (text, command) in enumerate(actions, 1):
            ctk.CTkButton(
                sidebar, text=text, command=command, anchor="w", height=34, corner_radius=8,
                fg_color=PRIMARY if row == 1 else "transparent", hover_color=NAV_HOVER,
                text_color="#FFFFFF",
            ).grid(row=row, column=0, padx=10, pady=2, sticky="ew")

        sheets = ctk.CTkFrame(sidebar, fg_color="#1B2B44", corner_radius=10)
        sheets.grid(row=4, column=0, padx=10, pady=14, sticky="nsew")
        ctk.CTkLabel(
            sheets, text="Trang tính", text_color="#D3E0F0", font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w", padx=10, pady=(8, 1))
        self.sheet_scroll = ctk.CTkScrollableFrame(sheets, fg_color="transparent")
        self.sheet_scroll.pack(fill="both", expand=True, padx=3, pady=(0, 6))
        row = ctk.CTkFrame(sheets, fg_color="transparent")
        row.pack(fill="x", padx=7, pady=(0, 7))
        for text, command, width in (
            ("+", self.add_sheet, 24), ("Đổi tên", self.rename_sheet, 62), ("Xóa", self.delete_sheet, 43),
        ):
            ctk.CTkButton(
                row, text=text, width=width, height=28, fg_color=NAV_HOVER,
                hover_color="#3C5370", command=command,
            ).pack(side="left", padx=2)
        self.mode_switch = ctk.CTkSwitch(
            sidebar, text="Giao diện tối", text_color="#D3E0F0", command=self.toggle_mode,
        )
        self.mode_switch.grid(row=5, column=0, padx=16, pady=(0, 14), sticky="w")
        if ctk.get_appearance_mode() == "Dark":
            self.mode_switch.select()

    def build_header(self):
        header = ctk.CTkFrame(self.content, corner_radius=0, fg_color=SURFACE)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=16, pady=9, sticky="w")
        ctk.CTkLabel(
            title_box, text="Quản lý bàn giao", text_color=TEXT, font=ctk.CTkFont(size=21, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(title_box, textvariable=self.file_var, text_color=MUTED).pack(anchor="w")
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.grid(row=0, column=1, padx=16, sticky="e")
        ctk.CTkLabel(actions, text="Tìm kiếm", text_color=MUTED).pack(side="left", padx=(0, 6))
        self.search_entry = ctk.CTkEntry(actions, width=175, height=32)
        self.search_entry.pack(side="left", padx=(0, 8))
        self.search_entry.bind("<KeyRelease>", self.search)
        self.search_entry.bind("<Escape>", self.clear_search)
        self.receiver_search = ctk.CTkCheckBox(
            actions, text="Receiver", width=80, variable=self.receiver_only, command=self.search,
        )
        self.receiver_search.pack(side="left", padx=(0, 10))
        self.button(actions, "Lưu", self.save, 58, primary=True).pack(side="left")

    def build_metrics(self):
        metrics = ctk.CTkFrame(self.content, fg_color="transparent")
        metrics.grid(row=1, column=0, padx=16, pady=(10, 7), sticky="ew")
        self.metric_values = []
        for col, label in enumerate(("Dòng có dữ liệu", "Receiver", "Trang tính", "Lưu tệp")):
            metrics.grid_columnconfigure(col, weight=1, uniform="metric")
            card = ctk.CTkFrame(metrics, fg_color=SURFACE, border_color=BORDER, border_width=1, corner_radius=10)
            card.grid(row=0, column=col, padx=(0 if col == 0 else 4, 0 if col == 3 else 4), sticky="ew")
            ctk.CTkLabel(card, text=label, text_color=MUTED, height=17).pack(anchor="w", padx=11, pady=(5, 0))
            value = ctk.CTkLabel(card, text="—", text_color=TEXT, font=ctk.CTkFont(size=17, weight="bold"))
            value.pack(anchor="w", padx=11, pady=(0, 5))
            self.metric_values.append(value)

    def build_toolbar(self):
        bar = ctk.CTkFrame(self.content, fg_color=SURFACE, border_color=BORDER, border_width=1, corner_radius=10)
        bar.grid(row=2, column=0, padx=16, pady=(0, 7), sticky="ew")
        bar.grid_columnconfigure(0, weight=1)
        commands = ctk.CTkFrame(bar, fg_color="transparent")
        commands.grid(row=0, column=0, padx=8, pady=(6, 3), sticky="w")
        self.button(commands, "＋ Dòng", self.insert_row, 64, primary=True).pack(side="left", padx=(0, 5))
        menu_groups = (
            ("Dòng và cột", 99, (("Xóa dòng đã chọn", self.delete_rows), ("Chèn cột", self.insert_col))),
            ("Định dạng", 91, (
                ("In đậm", lambda: self.format_font("bold")), ("In nghiêng", lambda: self.format_font("italic")),
                ("Màu nền", self.set_fill), ("Canh trái", lambda: self.align("left")),
                ("Canh giữa", lambda: self.align("center")),
            )),
            ("Công cụ Excel", 111, (
                ("Bật/tắt bộ lọc Excel", self.toggle_filter),
                ("Bật/tắt cố định tiêu đề", self.freeze_header), ("Tạo bảng Excel", self.make_table),
            )),
        )
        for label, width, choices in menu_groups:
            menu = tk.Menu(self, tearoff=False)
            for text, command in choices:
                menu.add_command(label=text, command=command)
            button = self.button(commands, f"{label} ▾", None, width)
            button.configure(command=lambda m=menu, b=button: self.popup_menu(m, b.winfo_rootx(), b.winfo_rooty() + b.winfo_height()))
            button.pack(side="left", padx=(0, 5))

        receiver_tools = ctk.CTkFrame(bar, fg_color="transparent")
        receiver_tools.grid(row=1, column=0, padx=8, pady=(3, 6), sticky="ew")
        self.auto_switch = ctk.CTkSwitch(
            receiver_tools, text="Tự điền theo Receiver", variable=self.auto_fill, width=145,
        )
        self.auto_switch.pack(side="left", padx=(2, 9))
        self.fill_button = self.button(receiver_tools, "Điền dòng đã chọn", self.fill_selection, 125)
        self.fill_button.pack(side="left", padx=(0, 5))
        self.settings_button = self.button(receiver_tools, "Chọn cột", self.configure_autofill, 72)
        self.settings_button.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(receiver_tools, textvariable=self.fill_var, text_color=MUTED).pack(side="left")

    @staticmethod
    def popup_menu(menu, x, y):
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def build_formula_bar(self):
        bar = ctk.CTkFrame(self.content, fg_color="transparent")
        bar.grid(row=3, column=0, padx=16, pady=(0, 7), sticky="ew")
        bar.grid_columnconfigure(1, weight=1)
        address = ctk.CTkEntry(bar, textvariable=self.cell_var, width=66, height=31)
        address.grid(row=0, column=0, padx=(0, 6))
        address.bind("<Return>", self.go_to_address)
        self.formula_entry = ctk.CTkEntry(bar, textvariable=self.formula_var, height=31)
        self.formula_entry.grid(row=0, column=1, sticky="ew")
        self.formula_entry.bind("<Return>", self.commit_formula)
        self.button(bar, "Áp dụng", self.commit_formula, 72).grid(row=0, column=2, padx=(6, 0))

    def build_grid(self):
        wrap = ctk.CTkFrame(self.content, fg_color=SURFACE, border_color=BORDER, border_width=1, corner_radius=10)
        wrap.grid(row=4, column=0, padx=16, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, style="Modern.Treeview", show="headings", selectmode="extended")
        y = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        x = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=(2, 0))
        y.grid(row=0, column=1, sticky="ns", pady=2)
        x.grid(row=1, column=0, sticky="ew", padx=2, pady=(0, 2))
        self.tree.bind("<Button-1>", self.point_cell)
        self.tree.bind("<Double-1>", self.begin_edit)
        self.tree.bind("<<TreeviewSelect>>", self.select_cell)
        self.tree.bind("<Button-3>", self.context_menu)
        self.tree.bind("<F2>", self.begin_edit)
        self.tree.bind("<Control-c>", lambda e: self.copy_rows())
        self.tree.bind("<Control-v>", lambda e: self.paste_rows())
        self.tree.bind("<Delete>", lambda e: self.clear_rows())

    def build_statusbar(self):
        bar = ctk.CTkFrame(self.content, fg_color="transparent")
        bar.grid(row=5, column=0, padx=16, pady=6, sticky="ew")
        ctk.CTkLabel(bar, textvariable=self.status_var, text_color=MUTED).pack(side="left")
        self.next_button = self.button(bar, "Tiếp ›", lambda: self.change_page(1), 59)
        self.next_button.pack(side="right")
        ctk.CTkLabel(bar, textvariable=self.page_var, text_color=TEXT).pack(side="right", padx=7)
        self.prev_button = self.button(bar, "‹ Trước", lambda: self.change_page(-1), 59)
        self.prev_button.pack(side="right")
        ctk.CTkLabel(bar, textvariable=self.result_var, text_color=MUTED).pack(side="right", padx=9)

    def create_new(self, skip_confirm=False):
        if not skip_confirm and not self.confirm_changes():
            return
        self.wb = Workbook()
        self.ws.title = "Handover"
        self.ws.append([
            "Ngày", "Ca", "Người bàn giao", "Người nhận", "Nội dung công việc",
            "Trạng thái", "Ưu tiên", "Người phụ trách", "Hạn hoàn thành", "Ghi chú",
        ])
        self.ws.append([date.today(), "Ca 1", "", "", "", "Chưa xử lý", "Bình thường", "", "", ""])
        for col, width in enumerate((13, 10, 20, 20, 38, 16, 14, 20, 16, 38), 1):
            self.ws.column_dimensions[get_column_letter(col)].width = width
        for cell in self.ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1E4E8C")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.ws.row_dimensions[1].height = 30
        self.ws.freeze_panes = "A2"
        self.path, self.dirty = None, False
        self.fill_settings.clear()
        self.refresh_all()
        self.set_status("Đã tạo sổ bàn giao mới")

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if path:
            self.open_path(path)

    def open_path(self, path, silent=False):
        if not silent and not self.confirm_changes():
            return
        path = Path(path)
        try:
            workbook = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
        except Exception as error:
            messagebox.showerror("Không mở được tệp", str(error))
            return
        self.wb, self.path, self.dirty = workbook, path, False
        self.fill_settings.clear()
        self.refresh_all()
        self.write_config({"last_file": str(path)})
        self.set_status("Đã mở tệp")

    def save(self):
        self.finish_edit()
        return self.save_path(self.path) if self.path else self.save_as()

    def save_as(self):
        self.finish_edit()
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsm" if self.wb.vba_archive else ".xlsx",
            initialfile=self.path.name if self.path else "Handover.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Excel Macro", "*.xlsm")],
        )
        return self.save_path(Path(path)) if path else False

    def save_path(self, path):
        try:
            self.wb.save(path)
        except PermissionError:
            messagebox.showerror("Không thể lưu", "Hãy đóng tệp đang mở trong Microsoft Excel rồi thử lại.")
            return False
        except Exception as error:
            messagebox.showerror("Lỗi lưu tệp", str(error))
            return False
        self.path, self.dirty = path, False
        self.file_var.set(path.name)
        self.write_config({"last_file": str(path)})
        self.update_save_status()
        self.set_status(f"Đã lưu lúc {datetime.now():%H:%M:%S}")
        return True

    def refresh_all(self):
        self.file_var.set(self.path.name if self.path else "Sổ mới chưa lưu")
        self.search_entry.delete(0, "end")
        self.page = 0
        self.active_cell = (1, 1)
        self.tree.selection_remove(self.tree.selection())
        for widget in self.sheet_scroll.winfo_children():
            widget.destroy()
        for index, name in enumerate(self.wb.sheetnames):
            active = index == self.wb.index(self.ws)
            ctk.CTkButton(
                self.sheet_scroll, text=name, anchor="w", height=31, corner_radius=7,
                text_color="#FFFFFF", fg_color=NAV_HOVER if active else "transparent", hover_color=NAV_HOVER,
                command=lambda i=index: self.switch_sheet(i),
            ).pack(fill="x", pady=2)
        self.refresh_grid(reset_scroll=True)

    def switch_sheet(self, index):
        self.finish_edit()
        self.wb.active = index
        self.refresh_all()
        self.set_status(f"Trang tính {self.ws.title}")

    def search(self, _event=None):
        self.page = 0
        self.refresh_grid(reset_scroll=True)

    def clear_search(self, _event=None):
        self.search_entry.delete(0, "end")
        self.search()

    def change_page(self, offset):
        self.page += offset
        self.refresh_grid(reset_scroll=True)

    def refresh_grid(self, reset_scroll=False):
        selected = self.tree.selection()
        old_x = self.tree.xview()[0]
        old_y = self.tree.yview()[0]
        headers = sheet_headers(self.ws)
        receiver_col = receiver_column(headers)
        if not receiver_col:
            self.receiver_only.set(False)
        self.receiver_search.configure(state="normal" if receiver_col else "disabled")
        self.update_receiver_tools(headers, receiver_col)

        data = {}
        for (row, col), cell in self.ws._cells.items():
            if has_value(cell.value):
                data.setdefault(row, {})[col] = cell

        search = self.search_entry.get().strip().casefold()
        if search:
            visible_rows = []
            for row, cells in sorted(data.items()):
                candidates = (cells.get(receiver_col),) if self.receiver_only.get() else cells.values()
                if any(cell is not None and search in self.show(cell.value).casefold() for cell in candidates):
                    visible_rows.append(row)
        else:
            end = max(min(max(data, default=1) + 12, MAX_EXCEL_ROW), self.active_cell[0])
            visible_rows = range(1, end + 1)

        count = len(visible_rows)
        self.page_count = max(1, (count + PAGE_SIZE - 1) // PAGE_SIZE)
        self.page = min(max(0, self.page), self.page_count - 1)
        start = self.page * PAGE_SIZE
        self.page_var.set(f"{self.page + 1} / {self.page_count}")
        self.result_var.set(f"{count} dòng phù hợp" if search else "Trang")
        self.prev_button.configure(state="normal" if self.page else "disabled")
        self.next_button.configure(state="normal" if self.page + 1 < self.page_count else "disabled")

        last_col = max(self.ws.max_column, self.active_cell[1])
        columns = ["row"] + [get_column_letter(col) for col in range(1, last_col + 1)]
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns
        self.tree.heading("row", text="Dòng", anchor="center")
        self.tree.column("row", width=57, minwidth=50, stretch=False, anchor="center")
        for col, letter in enumerate(columns[1:], 1):
            label = f"{letter} · {self.show(headers[col].value)}" if col in headers else letter
            self.tree.heading(letter, text=label, anchor="w", command=lambda c=col: self.sort_column(c))
            dimension = self.ws.column_dimensions.get(letter)
            width = int(dimension.width * 8) if dimension and dimension.width else 110
            heading_width = min(self.heading_font.measure(label) + 18, 280)
            self.tree.column(letter, width=min(max(width, heading_width, 86), 360),
                             minwidth=55, stretch=False, anchor="w")

        dark = ctk.get_appearance_mode() == "Dark"
        self.tree.tag_configure("header", background="#294363" if dark else "#D9E7FA",
                                font=(self.grid_font, 10, "bold"))
        self.tree.tag_configure("stripe", background="#1B283B" if dark else "#F2F6FC")
        for row in visible_rows[start:start + PAGE_SIZE]:
            cells = data.get(row, {})
            values = [row] + [self.show(cells[c].value) if c in cells else "" for c in range(1, last_col + 1)]
            tags = ("header",) if row == 1 else (("stripe",) if row % 2 == 0 else ())
            self.tree.insert("", "end", iid=str(row), values=values, tags=tags)
        if not reset_scroll:
            self.tree.selection_set([item for item in selected if self.tree.exists(item)])
            self.tree.xview_moveto(old_x)
            self.tree.yview_moveto(old_y)
        else:
            self.tree.yview_moveto(0)

        receivers = {
            key for row, cells in data.items()
            if row > 1 and (key := receiver_key(cells.get(receiver_col)))
        }
        values = (len(data) - int(1 in data), len(receivers), len(self.wb.sheetnames))
        for label, value in zip(self.metric_values, values):
            label.configure(text=str(value))
        self.update_save_status()
        self.show_active_cell()

    def update_save_status(self):
        state = "Có thay đổi" if self.dirty else ("Đã lưu" if self.path else "Sổ mới")
        self.metric_values[3].configure(text=state)

    def update_receiver_tools(self, headers, receiver_col):
        columns = self.autofill_columns(headers, receiver_col)
        state = "normal" if receiver_col else "disabled"
        self.auto_switch.configure(state=state)
        self.settings_button.configure(state=state)
        self.fill_button.configure(state="normal" if columns else "disabled")
        if not receiver_col:
            text = "Thiếu cột Receiver ở dòng 1"
        elif not columns:
            text = "Chưa chọn cột"
        else:
            text = f"{get_column_letter(receiver_col)} · {len(columns)} cột tự điền"
        self.fill_var.set(text)

    def autofill_columns(self, headers, receiver_col):
        if not receiver_col:
            return []
        settings = self.fill_settings.get(self.ws)
        return [
            col for col, cell in headers.items() if col != receiver_col
            and (header_key(cell.value) in DEFAULT_FILL_HEADERS if settings is None else cell in settings)
        ]

    def configure_autofill(self):
        ws = self.ws
        headers = sheet_headers(ws)
        receiver_col = receiver_column(headers)
        if not receiver_col:
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("Cột tự điền theo Receiver")
        dialog.geometry("445x440")
        dialog.minsize(390, 340)
        dialog.transient(self)
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(
            dialog, text="Chọn cột tự điền", font=ctk.CTkFont(size=18, weight="bold"), text_color=TEXT,
        ).grid(row=0, column=0, padx=17, pady=(14, 2), sticky="w")
        ctk.CTkLabel(
            dialog, text="Điền ô trống từ dòng gần nhất phía trên có cùng Receiver.\nÁp dụng cho trang tính này khi sổ đang mở.",
            justify="left", text_color=MUTED,
        ).grid(row=1, column=0, padx=17, pady=(0, 8), sticky="w")
        items = ctk.CTkScrollableFrame(dialog, fg_color=SURFACE)
        items.grid(row=2, column=0, padx=16, sticky="nsew")
        active = self.autofill_columns(headers, receiver_col)
        variables = {}
        for col, cell in headers.items():
            if col == receiver_col:
                continue
            variables[col] = tk.BooleanVar(dialog, value=col in active)
            ctk.CTkCheckBox(
                items, text=f"{get_column_letter(col)} · {self.show(cell.value)}", variable=variables[col],
            ).pack(anchor="w", padx=5, pady=4)

        def defaults():
            for col, variable in variables.items():
                variable.set(header_key(headers[col].value) in DEFAULT_FILL_HEADERS)

        def apply():
            self.fill_settings[ws] = {headers[col] for col, variable in variables.items() if variable.get()}
            self.update_receiver_tools(headers, receiver_col)
            dialog.destroy()

        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.grid(row=3, column=0, padx=16, pady=11, sticky="ew")
        self.button(actions, "Dùng mặc định", defaults, 115).pack(side="left")
        self.button(actions, "Áp dụng", apply, 80, primary=True).pack(side="right")
        dialog.bind("<Escape>", lambda e: dialog.destroy())
        dialog.update_idletasks()
        dialog.grab_set()
        dialog.focus_set()

    def fill_rows(self, rows, headers, receiver_col):
        return fill_receiver_rows(self.ws, rows, receiver_col, self.autofill_columns(headers, receiver_col))

    def report_fill(self, count, fallback):
        rows, cells = count
        self.set_status(f"Đã tự điền {cells} ô trong {rows} dòng" if cells else fallback)

    def fill_selection(self):
        rows = self.selected_rows()
        if not rows:
            self.set_status("Chọn dòng có Receiver để tự điền")
            return
        headers = sheet_headers(self.ws)
        count = self.fill_rows(rows, headers, receiver_column(headers))
        if count[1]:
            self.dirty = True
            self.refresh_grid()
        self.report_fill(count, "Không có ô trống phù hợp để tự điền")

    @staticmethod
    def show(value):
        if value is None:
            return ""
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y %H:%M")
        if isinstance(value, date):
            return value.strftime("%d/%m/%Y")
        return str(value)

    def show_active_cell(self):
        row, col = self.active_cell
        self.cell_var.set(f"{get_column_letter(col)}{row}")
        cell = self.ws._cells.get(self.active_cell)
        self.formula_var.set(self.show(cell.value) if cell is not None else "")

    def hit_cell(self, event):
        row = self.tree.identify_row(event.y)
        if row and self.tree.identify_region(event.x, event.y) == "cell":
            col = int(self.tree.identify_column(event.x)[1:]) - 1
            if col > 0:
                return int(row), col
        return None

    def point_cell(self, event):
        coordinate = self.hit_cell(event)
        if coordinate:
            self.active_cell = coordinate
            self.show_active_cell()

    def select_cell(self, _event=None):
        selected = self.tree.selection()
        if selected and str(self.active_cell[0]) not in selected:
            focused = self.tree.focus()
            self.active_cell = (int(focused if focused in selected else selected[0]), self.active_cell[1])
            self.show_active_cell()

    def address(self):
        try:
            ref = self.cell_var.get().strip().upper()
            if not ref.isascii() or not ref.isalnum():
                raise ValueError
            row, col = coordinate_to_tuple(ref)
            if not (1 <= row <= MAX_EXCEL_ROW and 1 <= col <= MAX_EXCEL_COL):
                raise ValueError
            return row, col
        except (ValueError, KeyError):
            messagebox.showwarning("Địa chỉ ô không hợp lệ", "Ví dụ hợp lệ: A1, D25, AA10")
            return None

    def go_to_address(self, _event=None):
        coordinate = self.address()
        if coordinate:
            self.active_cell = coordinate
            row, col = coordinate
            self.search_entry.delete(0, "end")
            self.page = (row - 1) // PAGE_SIZE
            self.refresh_grid(reset_scroll=True)
            item = str(row)
            self.tree.selection_set(item)
            self.tree.focus(item)
            self.tree.see(item)
            widths = [self.tree.column(column, "width") for column in self.tree["columns"]]
            self.tree.xview_moveto(sum(widths[:col]) / sum(widths))
            self.tree.focus_set()

    def begin_edit(self, event=None):
        coordinate = self.active_cell if event is None or event.keysym == "F2" else self.hit_cell(event)
        if not coordinate:
            return
        self.finish_edit()
        row, col = coordinate
        cell = self.ws._cells.get(coordinate)
        if isinstance(cell, MergedCell):
            self.set_status("Hãy sửa ô đầu tiên của vùng ô gộp")
            return "break"
        box = self.tree.bbox(str(row), f"#{col + 1}")
        if not box:
            return
        self.active_cell = self.edit_cell = coordinate
        self.show_active_cell()
        x, y, width, height = box
        self.editor = ttk.Entry(self.tree, style="Handover.TEntry")
        self.editor.insert(0, self.formula_var.get())
        self.editor.select_range(0, "end")
        self.editor.place(x=x, y=y, width=min(max(width, 120), self.tree.winfo_width() - x), height=height)
        self.editor.focus_set()
        self.editor.bind("<Return>", lambda e: self.finish_edit())
        self.editor.bind("<Escape>", lambda e: self.cancel_edit())
        self.editor.bind("<FocusOut>", lambda e: self.finish_edit(restore_focus=False))
        return "break"

    def finish_edit(self, restore_focus=True):
        if self.editor is None:
            return
        value, coordinate = self.editor.get(), self.edit_cell
        self.cancel_edit(restore_focus)
        self.apply_cell_edit(coordinate, value)

    def cancel_edit(self, restore_focus=True):
        if self.editor is not None:
            editor, self.editor = self.editor, None
            self.edit_cell = None
            editor.destroy()
            if restore_focus:
                self.tree.focus_set()

    def apply_cell_edit(self, coordinate, value):
        row, col = coordinate
        cell = self.ws._cells.get(coordinate)
        if isinstance(cell, MergedCell):
            self.set_status("Hãy sửa ô đầu tiên của vùng ô gộp")
            return
        self.active_cell = coordinate
        headers = sheet_headers(self.ws)
        receiver_col = receiver_column(headers)
        if self.show(cell.value if cell is not None else None) != value:
            text = value.strip()
            stored = value
            if text.startswith("="):
                stored = text
            elif col != receiver_col and text and not (text.startswith("0") and len(text) > 1):
                try:
                    stored = int(text)
                except ValueError:
                    try:
                        numeric = float(text.replace(",", "."))
                    except ValueError:
                        pass
                    else:
                        if abs(numeric) < float("inf"):
                            stored = numeric
            self.ws.cell(row, col).value = stored
            self.dirty = True
        count = (0, 0)
        status = f"Ô {get_column_letter(col)}{row}"
        if row > 1 and col == receiver_col and self.auto_fill.get():
            count = self.fill_rows([row], headers, receiver_col)
            self.dirty |= bool(count[1])
            status = "Không có ô phù hợp để tự điền"
        self.refresh_grid()
        self.report_fill(count, status)

    def commit_formula(self, _event=None):
        coordinate = self.address()
        if coordinate:
            self.apply_cell_edit(coordinate, self.formula_var.get())

    def selected_rows(self):
        return sorted(int(item) for item in self.tree.selection())

    def insert_row(self):
        rows = self.selected_rows()
        row = max(2, rows[0]) if rows else self.ws.max_row + 1
        if row > MAX_EXCEL_ROW:
            self.set_status("Đã đến dòng cuối của Excel")
            return
        if row <= self.ws.max_row:
            self.ws.insert_rows(row)
        self.ws.cell(row, 1)
        self.active_cell = (row, self.active_cell[1])
        self.dirty = True
        self.go_to_active_row()

    def go_to_active_row(self):
        self.show_active_cell()
        self.go_to_address()

    def delete_rows(self):
        rows = self.selected_rows()
        if rows and messagebox.askyesno("Xóa dòng", f"Xóa {len(rows)} dòng đã chọn?"):
            for row in reversed(rows):
                self.ws.delete_rows(row)
            self.mark_dirty_and_refresh()

    def insert_col(self):
        self.ws.insert_cols(self.active_cell[1])
        self.mark_dirty_and_refresh()

    def clear_rows(self):
        rows = set(self.selected_rows())
        if not rows:
            return
        for (row, _), cell in self.ws._cells.items():
            if row in rows and not isinstance(cell, MergedCell):
                cell.value = None
        self.mark_dirty_and_refresh()

    def copy_rows(self):
        rows = self.selected_rows()
        if not rows:
            return
        text = "\n".join(
            "\t".join(self.show(cell.value) if cell is not None else "" for cell in (
                self.ws._cells.get((row, col)) for col in range(1, self.ws.max_column + 1)
            ))
            for row in rows
        )
        self.clipboard_clear()
        self.clipboard_append(text)
        self.set_status(f"Đã sao chép {len(rows)} dòng")

    def paste_rows(self):
        try:
            lines = self.clipboard_get().splitlines()
        except tk.TclError:
            return
        if not lines:
            return
        selected = self.selected_rows()
        start = selected[0] if selected else max(2, self.ws.max_row + 1)
        if start + len(lines) - 1 > MAX_EXCEL_ROW or any(line.count("\t") >= MAX_EXCEL_COL for line in lines):
            self.set_status("Vùng dán vượt giới hạn của Excel")
            return

        pasted_widths = {}
        for row, line in enumerate(lines, start):
            entries = line.split("\t")
            for col, value in enumerate(entries, 1):
                current = self.ws._cells.get((row, col))
                if not isinstance(current, MergedCell):
                    self.ws.cell(row, col).value = value or None
            pasted_widths[row] = len(entries)

        headers = sheet_headers(self.ws)
        receiver_col = receiver_column(headers)
        count = (0, 0)
        if receiver_col and self.auto_fill.get():
            rows = [row for row, width in pasted_widths.items() if width >= receiver_col]
            count = self.fill_rows(rows, headers, receiver_col)
        self.mark_dirty_and_refresh()
        self.report_fill(count, f"Đã dán {len(lines)} dòng")

    def selected_cells_iter(self):
        rows = self.selected_rows() or [self.active_cell[0]]
        for row in rows:
            for col in range(1, self.ws.max_column + 1):
                cell = self.ws.cell(row, col)
                if not isinstance(cell, MergedCell):
                    yield cell

    def format_font(self, kind):
        for cell in self.selected_cells_iter():
            font = copy(cell.font)
            setattr(font, kind, not getattr(font, kind))
            cell.font = font
        self.mark_dirty_and_refresh()

    def set_fill(self):
        color = colorchooser.askcolor()[1]
        if color:
            for cell in self.selected_cells_iter():
                cell.fill = PatternFill("solid", fgColor=color.removeprefix("#"))
            self.mark_dirty_and_refresh()

    def align(self, alignment):
        for cell in self.selected_cells_iter():
            cell.alignment = Alignment(horizontal=alignment, vertical="center", wrap_text=True)
        self.mark_dirty_and_refresh()

    def freeze_header(self):
        self.ws.freeze_panes = None if self.ws.freeze_panes else "A2"
        self.dirty = True
        self.update_save_status()
        self.set_status("Đã đặt đóng băng hàng tiêu đề" if self.ws.freeze_panes else "Đã bỏ đóng băng hàng tiêu đề")

    def toggle_filter(self):
        self.ws.auto_filter.ref = None if self.ws.auto_filter.ref else self.ws.dimensions
        self.dirty = True
        self.update_save_status()
        self.set_status("Đã bật bộ lọc Excel" if self.ws.auto_filter.ref else "Đã tắt bộ lọc Excel")

    def make_table(self):
        try:
            table = Table(displayName=f"HandoverTable{datetime.now():%H%M%S}", ref=self.ws.dimensions)
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False,
            )
            self.ws.add_table(table)
        except Exception as error:
            messagebox.showerror("Không thể tạo bảng Excel", str(error))
            return
        self.dirty = True
        self.update_save_status()
        self.set_status("Đã tạo bảng Excel")

    def sort_column(self, col):
        ws = self.ws
        if any(merged.max_row > 1 for merged in ws.merged_cells.ranges):
            self.set_status("Không thể sắp xếp dữ liệu có ô gộp")
            return
        data = {}
        for (row, column), cell in ws._cells.items():
            if row > 1 and has_value(cell.value):
                data.setdefault(row, {})[column] = (cell.value, cell.data_type)
        if len(data) < 2:
            return

        def sort_key(row):
            value = data[row].get(col, (None, None))[0]
            return value is None, str(value).casefold()

        row_order = sorted(data)
        for destination, source in zip(row_order, sorted(row_order, key=sort_key)):
            for column in data[destination].keys() - data[source].keys():
                ws.cell(destination, column).value = None
            for column, (value, data_type) in data[source].items():
                cell = ws.cell(destination, column)
                cell.value = value
                cell.data_type = data_type
        self.mark_dirty_and_refresh()
        self.set_status(f"Đã sắp xếp theo cột {get_column_letter(col)}")

    def add_sheet(self):
        self.finish_edit()
        name = simpledialog.askstring("Thêm trang tính", "Tên trang tính", initialvalue="Handover mới")
        if name:
            try:
                worksheet = self.wb.create_sheet(name)
                self.wb.active = worksheet
            except ValueError as error:
                messagebox.showerror("Lỗi", str(error))
                return
            self.dirty = True
            self.refresh_all()

    def rename_sheet(self):
        name = simpledialog.askstring("Đổi tên", "Tên mới", initialvalue=self.ws.title)
        if name:
            try:
                self.ws.title = name
            except ValueError as error:
                messagebox.showerror("Lỗi", str(error))
                return
            self.dirty = True
            self.refresh_all()

    def delete_sheet(self):
        if len(self.wb.worksheets) <= 1:
            messagebox.showwarning("Không thể xóa", "Sổ cần ít nhất một trang tính.")
            return
        worksheet = self.ws
        if messagebox.askyesno("Xóa trang tính", f"Xóa trang tính '{worksheet.title}'?"):
            index = self.wb.index(worksheet)
            self.wb.remove(worksheet)
            self.fill_settings.pop(worksheet, None)
            self.wb.active = min(index, len(self.wb.sheetnames) - 1)
            self.dirty = True
            self.refresh_all()

    def mark_dirty_and_refresh(self):
        self.dirty = True
        self.refresh_grid()

    def set_status(self, text):
        self.status_var.set(text)

    def toggle_mode(self):
        theme = "Dark" if self.mode_switch.get() else "Light"
        ctk.set_appearance_mode(theme)
        self.apply_tree_theme()
        self.refresh_grid()
        self.write_config({"theme": theme})

    def context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if iid not in self.tree.selection():
            self.tree.selection_set(iid)
        self.point_cell(event)
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(label="Sửa ô", command=self.begin_edit)
        menu.add_command(label="Tự điền theo Receiver", command=self.fill_selection, state=self.fill_button.cget("state"))
        menu.add_separator()
        menu.add_command(label="Sao chép dòng", command=self.copy_rows)
        menu.add_command(label="Dán", command=self.paste_rows)
        menu.add_command(label="Xóa nội dung", command=self.clear_rows)
        menu.add_separator()
        menu.add_command(label="Chèn dòng", command=self.insert_row)
        menu.add_command(label="Xóa dòng", command=self.delete_rows)
        self.popup_menu(menu, event.x_root, event.y_root)

    def confirm_changes(self):
        self.finish_edit()
        if not self.dirty:
            return True
        answer = messagebox.askyesnocancel("Có thay đổi chưa lưu", "Bạn muốn lưu thay đổi trước khi tiếp tục?")
        return self.save() if answer else answer is not None

    def autosave(self):
        try:
            if self.dirty:
                path = self.path or Path.home() / "Handover.xlsx"
                backup = path.with_name(path.stem + "_autosave" + path.suffix)
                self.wb.save(backup)
                self.set_status(f"Đã tự sao lưu {backup.name}")
        except OSError:
            self.set_status("Chưa thể tự sao lưu tệp")
        finally:
            self.after(300000, self.autosave)

    def read_config(self):
        try:
            return json.loads(CONFIG.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def write_config(self, changes):
        try:
            config = self.read_config()
            config.update(changes)
            CONFIG.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass

    def close_app(self):
        if self.confirm_changes():
            self.destroy()


if __name__ == "__main__":
    ModernHandover(sys.argv[1] if len(sys.argv) > 1 else None).mainloop()

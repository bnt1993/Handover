# -*- coding: utf-8 -*-
"""Handover Excel Manager Modern UI."""
import os, sys, json
from pathlib import Path
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
import customtkinter as ctk
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.table import Table, TableStyleInfo

APP_NAME = "Handover Workspace"
CONFIG = Path.home() / ".handover_workspace.json"
MAX_ROWS, MAX_COLS = 1500, 100

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

COLORS = {
    "nav": "#0B1220", "nav_hover": "#172033", "primary": "#2563EB",
    "primary_hover": "#1D4ED8", "success": "#16A34A", "warning": "#F59E0B",
    "danger": "#DC2626", "light_bg": "#F3F6FB", "dark_bg": "#0B1120",
    "light_card": "#FFFFFF", "dark_card": "#111827", "muted": "#64748B"
}

class ModernHandover(ctk.CTk):
    def __init__(self, initial_file=None):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1540x900")
        self.minsize(1120, 680)
        self.protocol("WM_DELETE_WINDOW", self.close_app)
        self.wb = None
        self.path = None
        self.dirty = False
        self.active_cell = (1, 1)
        self.editor = None
        self.search_var = tk.StringVar()
        self.formula_var = tk.StringVar()
        self.cell_var = tk.StringVar(value="A1")
        self.status_var = tk.StringVar(value="Sẵn sàng")
        self.file_var = tk.StringVar(value="Chưa mở file")
        self.build_styles()
        self.build_ui()
        self.create_new(skip_confirm=True)
        cfg = self.read_config()
        candidate = initial_file or cfg.get("last_file")
        if candidate and os.path.exists(candidate):
            self.open_path(candidate, silent=True)
        self.after(300000, self.autosave)

    def build_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.apply_tree_theme()

    def apply_tree_theme(self):
        dark = ctk.get_appearance_mode() == "Dark"
        bg = "#111827" if dark else "#FFFFFF"
        fg = "#F8FAFC" if dark else "#172033"
        heading = "#1F2937" if dark else "#E9EEF8"
        border = "#334155" if dark else "#D8E0EC"
        self.style.configure("Modern.Treeview", background=bg, fieldbackground=bg, foreground=fg,
                             rowheight=34, borderwidth=0, relief="flat", font=("Segoe UI", 10))
        self.style.configure("Modern.Treeview.Heading", background=heading, foreground=fg,
                             relief="flat", borderwidth=0, padding=(10, 9), font=("Segoe UI Semibold", 10))
        self.style.map("Modern.Treeview", background=[("selected", COLORS["primary"])], foreground=[("selected", "#FFFFFF")])
        self.style.map("Modern.Treeview.Heading", background=[("active", "#CBD5E1" if not dark else "#334155")])
        self.style.configure("Vertical.TScrollbar", gripcount=0, background=border, troughcolor=bg, borderwidth=0)
        self.style.configure("Horizontal.TScrollbar", gripcount=0, background=border, troughcolor=bg, borderwidth=0)

    def build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.build_sidebar()
        self.content = ctk.CTkFrame(self, corner_radius=0, fg_color=(COLORS["light_bg"], COLORS["dark_bg"]))
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
        self.bind_all("<Control-c>", lambda e: self.copy_rows())
        self.bind_all("<Control-v>", lambda e: self.paste_rows())
        self.bind_all("<Delete>", lambda e: self.clear_rows())

    def build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=COLORS["nav"])
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_rowconfigure(5, weight=1)
        brand = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand.grid(row=0, column=0, padx=20, pady=(24, 28), sticky="ew")
        ctk.CTkLabel(brand, text="H", width=42, height=42, corner_radius=12, fg_color=COLORS["primary"],
                     text_color="white", font=ctk.CTkFont(size=21, weight="bold")).pack(side="left")
        labels = ctk.CTkFrame(brand, fg_color="transparent")
        labels.pack(side="left", padx=12)
        ctk.CTkLabel(labels, text="HANDOVER", text_color="white", font=ctk.CTkFont(size=17, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(labels, text="Excel Workspace", text_color="#94A3B8", font=ctk.CTkFont(size=11)).pack(anchor="w")
        self.nav_button(1, "＋  Tạo sổ mới", self.create_new, primary=True)
        self.nav_button(2, "▣  Mở file Excel", self.open_file)
        self.nav_button(3, "↓  Lưu thay đổi", self.save)
        self.nav_button(4, "↗  Lưu thành file mới", self.save_as)
        sheet_box = ctk.CTkFrame(self.sidebar, fg_color="#111B2E", corner_radius=14)
        sheet_box.grid(row=5, column=0, padx=14, pady=18, sticky="nsew")
        ctk.CTkLabel(sheet_box, text="DANH SÁCH SHEET", text_color="#94A3B8", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w", padx=14, pady=(14,8))
        self.sheet_scroll = ctk.CTkScrollableFrame(sheet_box, fg_color="transparent")
        self.sheet_scroll.pack(fill="both", expand=True, padx=5, pady=(0,8))
        row = ctk.CTkFrame(sheet_box, fg_color="transparent")
        row.pack(fill="x", padx=10, pady=(0,10))
        ctk.CTkButton(row, text="+", width=36, height=32, command=self.add_sheet).pack(side="left")
        ctk.CTkButton(row, text="Đổi tên", width=76, height=32, fg_color="#263449", hover_color="#334155", command=self.rename_sheet).pack(side="left", padx=5)
        ctk.CTkButton(row, text="Xóa", width=55, height=32, fg_color="#3B1D28", hover_color="#5B2432", command=self.delete_sheet).pack(side="left")
        bottom = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        bottom.grid(row=6, column=0, padx=18, pady=18, sticky="ew")
        ctk.CTkLabel(bottom, text="Giao diện", text_color="#94A3B8").pack(side="left")
        self.mode_switch = ctk.CTkSwitch(bottom, text="", width=42, command=self.toggle_mode)
        self.mode_switch.pack(side="right")

    def nav_button(self, row, text, command, primary=False):
        btn = ctk.CTkButton(self.sidebar, text=text, command=command, anchor="w", height=42, corner_radius=10,
                            font=ctk.CTkFont(size=13, weight="bold" if primary else "normal"),
                            fg_color=COLORS["primary"] if primary else "transparent",
                            hover_color=COLORS["primary_hover"] if primary else COLORS["nav_hover"])
        btn.grid(row=row, column=0, padx=14, pady=4, sticky="ew")

    def build_header(self):
        header = ctk.CTkFrame(self.content, height=78, corner_radius=0, fg_color=("#FFFFFF", "#111827"))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=24, pady=15, sticky="w")
        ctk.CTkLabel(title_box, text="Quản lý bàn giao", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(title_box, textvariable=self.file_var, text_color=COLORS["muted"], font=ctk.CTkFont(size=11)).pack(anchor="w")
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.grid(row=0, column=1, padx=22, pady=15, sticky="e")
        self.search_entry = ctk.CTkEntry(actions, textvariable=self.search_var, placeholder_text="Tìm trong sheet...", width=270, height=38, corner_radius=10)
        self.search_entry.pack(side="left", padx=(0,10))
        self.search_entry.bind("<KeyRelease>", lambda e: self.refresh_grid())
        ctk.CTkButton(actions, text="Lưu", width=88, height=38, corner_radius=10, command=self.save).pack(side="left")

    def build_metrics(self):
        self.metrics = ctk.CTkFrame(self.content, fg_color="transparent")
        self.metrics.grid(row=1, column=0, padx=22, pady=(18,10), sticky="ew")
        for i in range(4): self.metrics.grid_columnconfigure(i, weight=1)
        self.metric_values = []
        cards = [("Tổng số dòng", "0", COLORS["primary"]), ("Số sheet", "0", "#7C3AED"), ("Ô có dữ liệu", "0", COLORS["success"]), ("Trạng thái", "Sẵn sàng", COLORS["warning"])]
        for i,(title,val,color) in enumerate(cards):
            card = ctk.CTkFrame(self.metrics, corner_radius=14, fg_color=("#FFFFFF", "#111827"), border_width=1, border_color=("#E2E8F0", "#243047"))
            card.grid(row=0, column=i, padx=(0 if i==0 else 6, 0 if i==3 else 6), sticky="ew")
            ctk.CTkFrame(card, width=5, height=48, corner_radius=3, fg_color=color).pack(side="left", padx=(14,12), pady=14)
            tx = ctk.CTkFrame(card, fg_color="transparent"); tx.pack(side="left", pady=12)
            ctk.CTkLabel(tx, text=title, text_color=COLORS["muted"], font=ctk.CTkFont(size=11)).pack(anchor="w")
            lbl = ctk.CTkLabel(tx, text=val, font=ctk.CTkFont(size=18, weight="bold")); lbl.pack(anchor="w")
            self.metric_values.append(lbl)

    def build_toolbar(self):
        bar = ctk.CTkFrame(self.content, corner_radius=12, fg_color=("#FFFFFF", "#111827"), border_width=1, border_color=("#E2E8F0", "#243047"))
        bar.grid(row=2, column=0, padx=22, pady=8, sticky="ew")
        commands = [
            ("B", lambda: self.format_font("bold")), ("I", lambda: self.format_font("italic")),
            ("Màu nền", self.set_fill), ("Canh trái", lambda: self.align("left")),
            ("Canh giữa", lambda: self.align("center")), ("+ Dòng", self.insert_row),
            ("− Dòng", self.delete_rows), ("+ Cột", self.insert_col),
            ("Bộ lọc", self.toggle_filter), ("Đóng băng", self.freeze_header), ("Tạo Table", self.make_table)
        ]
        for text,cmd in commands:
            ctk.CTkButton(bar, text=text, command=cmd, height=32, width=58 if len(text)<3 else 82, corner_radius=8,
                          fg_color=("#EEF2F7", "#1B2638"), text_color=("#172033", "#E5E7EB"), hover_color=("#DCE5F2", "#2A394F")).pack(side="left", padx=4, pady=8)

    def build_formula_bar(self):
        box = ctk.CTkFrame(self.content, fg_color="transparent")
        box.grid(row=3, column=0, padx=22, pady=(2,8), sticky="ew")
        box.grid_columnconfigure(1, weight=1)
        ctk.CTkEntry(box, textvariable=self.cell_var, width=92, height=36, corner_radius=9).grid(row=0,column=0,padx=(0,8))
        self.formula_entry = ctk.CTkEntry(box, textvariable=self.formula_var, height=36, corner_radius=9, placeholder_text="Nội dung hoặc công thức Excel (=SUM(...))")
        self.formula_entry.grid(row=0,column=1,sticky="ew")
        self.formula_entry.bind("<Return>", self.commit_formula)

    def build_grid(self):
        wrap = ctk.CTkFrame(self.content, corner_radius=14, fg_color=("#FFFFFF", "#111827"), border_width=1, border_color=("#E2E8F0", "#243047"))
        wrap.grid(row=4, column=0, padx=22, pady=(0,10), sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1); wrap.grid_rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, style="Modern.Treeview", show="headings", selectmode="extended")
        y = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        x = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.tree.grid(row=0,column=0,sticky="nsew",padx=(2,0),pady=(2,0)); y.grid(row=0,column=1,sticky="ns",pady=2); x.grid(row=1,column=0,sticky="ew",padx=2)
        self.tree.bind("<Double-1>", self.begin_edit)
        self.tree.bind("<<TreeviewSelect>>", self.select_cell)
        self.tree.bind("<Button-3>", self.context_menu)

    def build_statusbar(self):
        bar = ctk.CTkFrame(self.content, height=35, corner_radius=0, fg_color="transparent")
        bar.grid(row=5,column=0,padx=24,pady=(0,8),sticky="ew")
        ctk.CTkLabel(bar,textvariable=self.status_var,text_color=COLORS["muted"],font=ctk.CTkFont(size=11)).pack(side="left")
        ctk.CTkLabel(bar,text="Tự động sao lưu mỗi 5 phút",text_color=COLORS["muted"],font=ctk.CTkFont(size=11)).pack(side="right")

    @property
    def ws(self): return self.wb.active if self.wb else None

    def create_new(self, skip_confirm=False):
        if not skip_confirm and not self.confirm_changes(): return
        self.wb=Workbook(); ws=self.wb.active; ws.title="Handover"
        ws.append(["Ngày","Ca","Người bàn giao","Người nhận","Nội dung công việc","Trạng thái","Ưu tiên","Người phụ trách","Hạn hoàn thành","Ghi chú"])
        ws.append([date.today(),"Ca 1","","","","Chưa xử lý","Bình thường","","",""])
        widths=[13,10,20,20,38,16,14,20,16,38]
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
        self.style_header(); ws.freeze_panes="A2"; self.path=None; self.dirty=False
        self.refresh_all(); self.set_status("Đã tạo sổ bàn giao mới")

    def open_file(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm"),("Tất cả","*.*")])
        if p: self.open_path(p)

    def open_path(self,p,silent=False):
        if not silent and not self.confirm_changes(): return
        try:
            self.set_status("Đang mở file...")
            self.wb=load_workbook(p,keep_vba=p.lower().endswith(".xlsm"),data_only=False)
            self.path=p; self.dirty=False; self.refresh_all(); self.write_config({"last_file":p}); self.set_status("Đã mở file")
        except Exception as e: messagebox.showerror("Không mở được file",str(e))

    def save(self):
        if not self.path: return self.save_as()
        try:
            self.wb.save(self.path); self.dirty=False; self.file_var.set(os.path.basename(self.path)); self.set_status(f"Đã lưu lúc {datetime.now():%H:%M:%S}"); return True
        except PermissionError: messagebox.showerror("Không thể lưu","Hãy đóng file đang mở trong Microsoft Excel rồi thử lại.")
        except Exception as e: messagebox.showerror("Lỗi lưu file",str(e))
        return False

    def save_as(self):
        p=filedialog.asksaveasfilename(defaultextension=".xlsx",initialfile="Handover.xlsx",filetypes=[("Excel","*.xlsx"),("Excel Macro","*.xlsm")])
        if not p:return False
        self.path=p; self.write_config({"last_file":p}); return self.save()

    def refresh_all(self):
        self.file_var.set(os.path.basename(self.path) if self.path else "Sổ mới chưa lưu")
        for w in self.sheet_scroll.winfo_children(): w.destroy()
        for idx,name in enumerate(self.wb.sheetnames):
            active=idx==self.wb.index(self.wb.active)
            ctk.CTkButton(self.sheet_scroll,text=("●  " if active else "   ")+name,anchor="w",height=34,corner_radius=8,
                          fg_color="#23324A" if active else "transparent",hover_color="#23324A",
                          command=lambda i=idx:self.switch_sheet(i)).pack(fill="x",pady=2)
        self.refresh_grid(); self.update_metrics()

    def refresh_grid(self):
        if not self.ws:return
        self.tree.delete(*self.tree.get_children())
        max_col=max(1,min(self.ws.max_column,MAX_COLS)); columns=[get_column_letter(i) for i in range(1,max_col+1)]
        self.tree["columns"]=columns
        for i,col in enumerate(columns,1):
            self.tree.heading(col,text=col,command=lambda c=i:self.sort_column(c))
            width=int((self.ws.column_dimensions[col].width or 12)*8)
            self.tree.column(col,width=min(max(width,86),340),minwidth=55,stretch=False,anchor="w")
        q=self.search_var.get().strip().lower(); max_row=max(1,min(self.ws.max_row+25,MAX_ROWS))
        for r in range(1,max_row+1):
            vals=[self.show(self.ws.cell(r,c).value) for c in range(1,max_col+1)]
            if q and not any(q in str(v).lower() for v in vals):continue
            self.tree.insert("","end",iid=str(r),values=vals,tags=("head" if r==1 else "data",))
        dark=ctk.get_appearance_mode()=="Dark"
        self.tree.tag_configure("head",background="#DCE7F6" if not dark else "#1E3A5F",font=("Segoe UI Semibold",10))
        self.update_metrics()

    def update_metrics(self):
        if not self.ws:return
        populated=sum(1 for row in self.ws.iter_rows() for c in row if c.value not in (None,""))
        vals=[max(0,self.ws.max_row-1),len(self.wb.sheetnames),populated,"Đã sửa" if self.dirty else "Đã lưu"]
        for lbl,val in zip(self.metric_values,vals): lbl.configure(text=str(val))

    @staticmethod
    def show(v):
        if v is None:return ""
        if isinstance(v,datetime):return v.strftime("%d/%m/%Y %H:%M")
        if isinstance(v,date):return v.strftime("%d/%m/%Y")
        return str(v)

    def switch_sheet(self,index): self.wb.active=index; self.search_var.set(""); self.refresh_all()

    def select_cell(self,_=None):
        sel=self.tree.selection()
        if not sel:return
        r=int(sel[0]); self.active_cell=(r,1); self.cell_var.set(f"A{r}")
        self.formula_var.set(self.show(self.ws.cell(r,1).value)); self.set_status(f"Dòng {r} • {len(sel)} dòng được chọn",update_metric=False)

    def begin_edit(self,event):
        row=self.tree.identify_row(event.y); colid=self.tree.identify_column(event.x)
        if not row or not colid:return
        c=int(colid[1:]); box=self.tree.bbox(row,colid)
        if not box:return
        if self.editor:self.editor.destroy()
        x,y,w,h=box; val=self.show(self.ws.cell(int(row),c).value)
        self.editor=ttk.Entry(self.tree); self.editor.insert(0,val); self.editor.select_range(0,"end")
        self.editor.place(x=x,y=y,width=max(w,140),height=h); self.editor.focus_set()
        self.active_cell=(int(row),c); self.cell_var.set(f"{get_column_letter(c)}{row}"); self.formula_var.set(val)
        self.editor.bind("<Return>",lambda e:self.finish_edit())
        self.editor.bind("<Escape>",lambda e:self.cancel_edit())
        self.editor.bind("<FocusOut>",lambda e:self.finish_edit())

    def finish_edit(self):
        if not self.editor:return
        value=self.editor.get(); self.editor.destroy(); self.editor=None
        r,c=self.active_cell; self.assign_value(r,c,value); self.refresh_grid()

    def cancel_edit(self):
        if self.editor:self.editor.destroy(); self.editor=None

    def assign_value(self,r,c,value):
        text=value.strip() if isinstance(value,str) else value
        if isinstance(text,str) and text.startswith("="):v=text
        elif isinstance(text,str) and text and not(text.startswith("0") and len(text)>1):
            try:v=int(text)
            except ValueError:
                try:v=float(text.replace(",","."))
                except ValueError:v=value
        else:v=value
        self.ws.cell(r,c).value=v; self.mark_dirty()

    def commit_formula(self,_=None):
        try:
            ref=self.cell_var.get().upper(); letters="".join(x for x in ref if x.isalpha()); digits="".join(x for x in ref if x.isdigit())
            self.assign_value(int(digits),column_index_from_string(letters),self.formula_var.get()); self.refresh_grid()
        except Exception: messagebox.showwarning("Địa chỉ ô không hợp lệ","Ví dụ hợp lệ: A1, D25, AA10")

    def selected_rows(self):return sorted(int(x) for x in self.tree.selection())
    def insert_row(self):
        rows=self.selected_rows(); self.ws.insert_rows(rows[0] if rows else self.ws.max_row+1); self.mark_dirty(); self.refresh_grid()
    def delete_rows(self):
        rows=self.selected_rows()
        if rows and messagebox.askyesno("Xóa dòng",f"Xóa {len(rows)} dòng đã chọn?"):
            for r in reversed(rows):self.ws.delete_rows(r)
            self.mark_dirty(); self.refresh_grid()
    def insert_col(self):
        self.ws.insert_cols(self.active_cell[1]); self.mark_dirty(); self.refresh_grid()
    def clear_rows(self):
        for r in self.selected_rows():
            for c in range(1,self.ws.max_column+1):self.ws.cell(r,c).value=None
        self.mark_dirty(); self.refresh_grid()
    def copy_rows(self):
        rows=self.selected_rows()
        if not rows:return
        text="\n".join("\t".join(self.show(self.ws.cell(r,c).value) for c in range(1,self.ws.max_column+1)) for r in rows)
        self.clipboard_clear(); self.clipboard_append(text); self.set_status(f"Đã sao chép {len(rows)} dòng",False)
    def paste_rows(self):
        try:text=self.clipboard_get()
        except Exception:return
        rows=self.selected_rows(); start=rows[0] if rows else self.ws.max_row+1
        for i,line in enumerate(text.splitlines()):
            for j,val in enumerate(line.split("\t"),1):self.ws.cell(start+i,j).value=val
        self.mark_dirty(); self.refresh_grid()

    def selected_cells_iter(self):
        rows=self.selected_rows() or [self.active_cell[0]]
        for r in rows:
            for c in range(1,self.ws.max_column+1):yield self.ws.cell(r,c)
    def format_font(self,kind):
        for cell in self.selected_cells_iter():
            f=cell.font; cell.font=Font(name=f.name or "Calibri",size=f.sz or 11,bold=(not f.bold if kind=="bold" else f.bold),italic=(not f.italic if kind=="italic" else f.italic),color=f.color)
        self.mark_dirty(); self.refresh_grid()
    def set_fill(self):
        color=colorchooser.askcolor()[1]
        if color:
            for cell in self.selected_cells_iter():cell.fill=PatternFill("solid",fgColor=color.replace("#",""))
            self.mark_dirty()
    def align(self,a):
        for cell in self.selected_cells_iter():cell.alignment=Alignment(horizontal=a,vertical="center",wrap_text=True)
        self.mark_dirty()
    def freeze_header(self):self.ws.freeze_panes=None if self.ws.freeze_panes else "A2"; self.mark_dirty()
    def toggle_filter(self):self.ws.auto_filter.ref=None if self.ws.auto_filter.ref else self.ws.dimensions; self.mark_dirty()
    def make_table(self):
        try:
            table=Table(displayName=f"HandoverTable{datetime.now():%H%M%S}",ref=self.ws.dimensions)
            table.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2",showRowStripes=True,showColumnStripes=False)
            self.ws.add_table(table); self.mark_dirty(); self.set_status("Đã tạo Excel Table",False)
        except Exception as e:messagebox.showerror("Không thể tạo Table",str(e))
    def sort_column(self,col):
        if self.ws.max_row<3:return
        rows=[[self.ws.cell(r,c).value for c in range(1,self.ws.max_column+1)] for r in range(2,self.ws.max_row+1)]
        rows.sort(key=lambda x:(x[col-1] is None,str(x[col-1]).lower()))
        for r,row in enumerate(rows,2):
            for c,v in enumerate(row,1):self.ws.cell(r,c).value=v
        self.mark_dirty(); self.refresh_grid()

    def add_sheet(self):
        name=simpledialog.askstring("Thêm sheet","Tên sheet:",initialvalue="Handover mới")
        if name:
            try:self.wb.create_sheet(name);self.wb.active=len(self.wb.sheetnames)-1;self.mark_dirty();self.refresh_all()
            except Exception as e:messagebox.showerror("Lỗi",str(e))
    def rename_sheet(self):
        name=simpledialog.askstring("Đổi tên","Tên mới:",initialvalue=self.ws.title)
        if name:
            try:self.ws.title=name;self.mark_dirty();self.refresh_all()
            except Exception as e:messagebox.showerror("Lỗi",str(e))
    def delete_sheet(self):
        if len(self.wb.sheetnames)<=1:return messagebox.showwarning("Không thể xóa","Workbook cần ít nhất một sheet.")
        if messagebox.askyesno("Xóa sheet",f"Xóa sheet '{self.ws.title}'?"):self.wb.remove(self.ws);self.mark_dirty();self.refresh_all()

    def style_header(self):
        for cell in self.ws[1]:
            cell.font=Font(bold=True,color="FFFFFF");cell.fill=PatternFill("solid",fgColor="1E4E8C");cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        self.ws.row_dimensions[1].height=30
    def mark_dirty(self):self.dirty=True;self.update_metrics()
    def set_status(self,text,update_metric=True):
        self.status_var.set(text)
        if update_metric and self.metric_values:self.metric_values[3].configure(text=text[:14])
    def toggle_mode(self):
        ctk.set_appearance_mode("Dark" if self.mode_switch.get() else "Light");self.apply_tree_theme();self.refresh_grid()
    def context_menu(self,event):
        iid=self.tree.identify_row(event.y)
        if iid:self.tree.selection_set(iid)
        m=tk.Menu(self,tearoff=False);m.add_command(label="Sửa ô",command=lambda:self.begin_edit(event));m.add_separator();m.add_command(label="Sao chép dòng",command=self.copy_rows);m.add_command(label="Dán",command=self.paste_rows);m.add_command(label="Xóa nội dung",command=self.clear_rows);m.add_separator();m.add_command(label="Chèn dòng",command=self.insert_row);m.add_command(label="Xóa dòng",command=self.delete_rows);m.tk_popup(event.x_root,event.y_root)
    def confirm_changes(self):
        if not self.dirty:return True
        ans=messagebox.askyesnocancel("Có thay đổi chưa lưu","Bạn muốn lưu thay đổi trước khi tiếp tục?")
        if ans is None:return False
        return self.save() if ans else True
    def autosave(self):
        try:
            if self.dirty:
                p=Path(self.path) if self.path else Path.home()/"Handover.xlsx"
                backup=p.with_name(p.stem+"_autosave"+p.suffix);self.wb.save(backup);self.set_status(f"Đã tự lưu {backup.name}",False)
        finally:self.after(300000,self.autosave)
    def read_config(self):
        try:return json.loads(CONFIG.read_text(encoding="utf-8"))
        except Exception:return {}
    def write_config(self,d):
        try:
            cfg=self.read_config();cfg.update(d);CONFIG.write_text(json.dumps(cfg,ensure_ascii=False),encoding="utf-8")
        except Exception:pass
    def close_app(self):
        if self.confirm_changes():self.destroy()

if __name__=="__main__":
    ModernHandover(sys.argv[1] if len(sys.argv)>1 else None).mainloop()

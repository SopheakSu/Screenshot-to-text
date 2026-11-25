"""
ScreenToExcel - paste or open a screenshot of the alarm list, read it with OCR,
and copy the rows in your report's column order so you can paste into Excel.

Hotkeys
  Ctrl+V       (in this window)       paste a screenshot from the clipboard
  Ctrl+Shift+C (in this window)       copy all rows for Excel
  Delete       (in this window)       remove selected rows
  Double-click a cell to edit it.
"""
import ctypes
import json
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk, ImageGrab

import ocr_core

APP = "ScreenToExcel"
CFG_PATH = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP, "settings.json")
DEFAULT_CFG = {
    "engine": "auto",
    "type_text": "Intruder",
    "columns": {"Time": False, "Code": True, "Name": True, "Contact": True,
                "Area": True, "Zone": True, "Type": True},
    "append": False,
    "header": False,
    "auto_copy": True,
    "skip_duplicates": False,
}

# Sharp coordinates on high-DPI screens (must happen before Tk starts)
if sys.platform == "win32":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def load_cfg():
    cfg = json.loads(json.dumps(DEFAULT_CFG))
    try:
        with open(CFG_PATH, encoding="utf-8") as f:
            saved = json.load(f)
        cfg.update({k: v for k, v in saved.items() if k != "columns"})
        cfg["columns"].update(saved.get("columns", {}))
    except Exception:
        pass
    return cfg


def save_cfg(cfg):
    try:
        os.makedirs(os.path.dirname(CFG_PATH), exist_ok=True)
        with open(CFG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


# ------------------------------------------------------------------- app
class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_cfg()
        self.q = queue.Queue()
        self.busy = False
        root.title("Screen → Excel (OCR)")
        root.geometry("1150x640")
        root.minsize(850, 450)
        try:
            ttk.Style().theme_use("vista")
        except Exception:
            pass
        self.build_ui()
        self.refresh_columns()
        root.bind("<Control-v>", lambda e: self.paste())
        root.bind("<Control-V>", lambda e: self.paste())
        root.bind("<Control-Shift-C>", lambda e: self.copy_all())
        root.bind("<Delete>", lambda e: self.delete_selected())
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(100, self.poll)

    # ---------- UI
    def build_ui(self):
        top = ttk.Frame(self.root, padding=6)
        top.pack(fill="x")
        ttk.Button(top, text="📋  Paste image  (Ctrl+V)", command=self.paste).pack(side="left", padx=2)
        ttk.Button(top, text="📂  Open image…", command=self.open_file).pack(side="left", padx=2)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(top, text="⧉  Copy all for Excel", command=self.copy_all).pack(side="left", padx=2)
        ttk.Button(top, text="Copy selected", command=self.copy_selected).pack(side="left", padx=2)
        ttk.Button(top, text="Clear", command=self.clear).pack(side="left", padx=2)

        opt = ttk.LabelFrame(self.root, text="Options", padding=6)
        opt.pack(fill="x", padx=6)
        ttk.Label(opt, text="Columns to copy:").grid(row=0, column=0, sticky="w")
        self.col_vars = {}
        for i, name in enumerate(ocr_core.FIELDS):
            v = tk.BooleanVar(value=self.cfg["columns"].get(name, False))
            label = name + (" (blank)" if name == "Contact" else "")
            ttk.Checkbutton(opt, text=label, variable=v, command=self.refresh_columns).grid(
                row=0, column=i + 1, sticky="w", padx=4)
            self.col_vars[name] = v

        ttk.Label(opt, text="Type text:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.type_var = tk.StringVar(value=self.cfg["type_text"])
        ttk.Entry(opt, textvariable=self.type_var, width=12).grid(row=1, column=1, sticky="w", pady=(6, 0))
        ttk.Label(opt, text="OCR engine:").grid(row=1, column=2, sticky="e", pady=(6, 0))
        self.engine_var = tk.StringVar(value=self.cfg["engine"])
        ttk.Combobox(opt, textvariable=self.engine_var, values=["auto", "windows", "tesseract"],
                     width=10, state="readonly").grid(row=1, column=3, sticky="w", pady=(6, 0))
        self.append_var = tk.BooleanVar(value=self.cfg["append"])
        self.header_var = tk.BooleanVar(value=self.cfg["header"])
        self.autocopy_var = tk.BooleanVar(value=self.cfg["auto_copy"])
        self.dedupe_var = tk.BooleanVar(value=self.cfg["skip_duplicates"])
        ttk.Checkbutton(opt, text="Keep old rows (add new screenshot below)", variable=self.append_var).grid(row=1, column=4, columnspan=2, sticky="w")
        ttk.Checkbutton(opt, text="Skip duplicates", variable=self.dedupe_var).grid(row=1, column=6, sticky="w")
        ttk.Checkbutton(opt, text="Auto-copy after scan", variable=self.autocopy_var).grid(row=1, column=7, sticky="w")
        ttk.Checkbutton(opt, text="Include header row", variable=self.header_var).grid(row=1, column=8, sticky="w")

        body = ttk.PanedWindow(self.root, orient="vertical")
        body.pack(fill="both", expand=True, padx=6, pady=6)

        tf = ttk.Frame(body)
        self.tree = ttk.Treeview(tf, columns=ocr_core.FIELDS, show="headings", selectmode="extended")
        widths = {"Time": 95, "Code": 95, "Name": 330, "Contact": 90, "Area": 50, "Zone": 240, "Type": 80}
        for c in ocr_core.FIELDS:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths[c], anchor="w", stretch=c in ("Name", "Zone"))
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self.edit_cell)
        body.add(tf, weight=4)

        pf = ttk.LabelFrame(body, text="Last screenshot", padding=2)
        self.preview = ttk.Label(pf)
        self.preview.pack(fill="both", expand=True)
        body.add(pf, weight=1)

        self.status = tk.StringVar(value="Copy a screenshot (from any app), then press Ctrl+V here — or click Open image.")
        ttk.Label(self.root, textvariable=self.status, relief="sunken", anchor="w", padding=3).pack(fill="x", side="bottom")

    def refresh_columns(self):
        shown = [c for c in ocr_core.FIELDS if self.col_vars[c].get()]
        self.tree["displaycolumns"] = shown or ocr_core.FIELDS

    # ---------- input
    def paste(self):
        img = ImageGrab.grabclipboard()
        if isinstance(img, list) and img:  # a copied file
            try:
                img = Image.open(img[0])
            except Exception:
                img = None
        if not isinstance(img, Image.Image):
            self.status.set("Clipboard has no image. Copy a screenshot first, then press Ctrl+V.")
            return
        self.process(img)

    def open_file(self):
        p = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff"), ("All", "*.*")])
        if p:
            self.process(Image.open(p))

    # ---------- OCR (background thread)
    def process(self, img):
        if self.busy:
            return
        self.busy = True
        self.show_preview(img)
        self.status.set("Reading text…")
        self.root.config(cursor="watch")
        engine, type_text = self.engine_var.get(), self.type_var.get()

        def work():
            try:
                rows, used = ocr_core.image_to_rows(img.copy(), engine, type_text)
                self.q.put(("rows", (rows, used)))
            except Exception as e:
                self.q.put(("error", str(e)))

        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "status":
                    self.status.set(data)
                elif kind == "error":
                    self.done()
                    messagebox.showerror(APP, data)
                    self.status.set("OCR failed.")
                elif kind == "rows":
                    self.done()
                    self.add_rows(*data)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def done(self):
        self.busy = False
        self.root.config(cursor="")

    def add_rows(self, rows, used):
        if not self.append_var.get():
            self.tree.delete(*self.tree.get_children())
        existing = {tuple(self.tree.set(i, c) for c in ("Code", "Time", "Zone")) for i in self.tree.get_children()}
        added = skipped = 0
        for r in rows:
            key = (r.fields["Code"], r.fields["Time"], r.fields["Zone"])
            if self.dedupe_var.get() and key in existing:
                skipped += 1
                continue
            existing.add(key)
            self.tree.insert("", "end", values=[r.fields[c] for c in ocr_core.FIELDS])
            added += 1
        msg = f"{added} row(s) added via {used} OCR"
        if skipped:
            msg += f", {skipped} duplicate(s) skipped"
        if added and self.autocopy_var.get():
            self.copy_all(quiet=True)
            msg += " — copied, go to Excel and press Ctrl+V"
        if not rows:
            msg = "No rows found. Include the code column (01-xxxxxx) in your selection."
        self.status.set(msg + ".  Double-click a cell to fix it.")

    def show_preview(self, img):
        im = img.copy()
        im.thumbnail((1100, 140))
        self._pv = ImageTk.PhotoImage(im)
        self.preview.configure(image=self._pv)

    # ---------- table editing
    def edit_cell(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return
        name = self.tree["displaycolumns"][int(col[1:]) - 1]
        if name == "#all":
            name = ocr_core.FIELDS[int(col[1:]) - 1]
        x, y, w, h = self.tree.bbox(item, col)
        e = ttk.Entry(self.tree)
        e.insert(0, self.tree.set(item, name))
        e.select_range(0, "end")
        e.place(x=x, y=y, width=w, height=h)
        e.focus()

        def save(_=None):
            self.tree.set(item, name, e.get())
            e.destroy()
        e.bind("<Return>", save)
        e.bind("<FocusOut>", save)
        e.bind("<Escape>", lambda _: e.destroy())

    def delete_selected(self):
        if isinstance(self.root.focus_get(), (ttk.Entry, tk.Entry)):
            return
        for i in self.tree.selection():
            self.tree.delete(i)

    def clear(self):
        self.tree.delete(*self.tree.get_children())
        self.status.set("Cleared.")

    # ---------- copy
    def _tsv(self, items):
        cols = [c for c in ocr_core.FIELDS if self.col_vars[c].get()]
        lines = []
        if self.header_var.get():
            lines.append("\t".join(cols))
        for i in items:
            vals = []
            for c in cols:
                v = "" if c == "Contact" and not self.tree.set(i, c) else self.tree.set(i, c)
                vals.append(str(v).replace("\t", " "))
            lines.append("\t".join(vals))
        return "\r\n".join(lines)

    def _to_clipboard(self, text, n, quiet):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        if not quiet:
            self.status.set(f"Copied {n} row(s). Click the first cell in Excel and press Ctrl+V.")

    def copy_all(self, quiet=False):
        items = self.tree.get_children()
        if items:
            self._to_clipboard(self._tsv(items), len(items), quiet)

    def copy_selected(self):
        items = self.tree.selection()
        if items:
            self._to_clipboard(self._tsv(items), len(items), False)

    def on_close(self):
        self.cfg.update({
            "engine": self.engine_var.get(),
            "type_text": self.type_var.get(),
            "columns": {c: v.get() for c, v in self.col_vars.items()},
            "append": self.append_var.get(),
            "header": self.header_var.get(),
            "auto_copy": self.autocopy_var.get(),
            "skip_duplicates": self.dedupe_var.get(),
        })
        save_cfg(self.cfg)
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()

"""
ScreenToExcel - paste (Ctrl+V) or open a picture, read the text with OCR,
and copy it so it pastes into Excel: one picture row = one Excel row,
separate columns in the picture = separate Excel columns.
"""
import ctypes
import queue
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk, ImageGrab

import ocr_core

APP = "ScreenToExcel"

if sys.platform == "win32":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.busy = False
        root.title("Screen → Excel (OCR)")
        root.geometry("760x620")
        root.minsize(500, 400)

        bar = ttk.Frame(root, padding=6)
        bar.pack(fill="x")
        ttk.Button(bar, text="📋 Paste image (Ctrl+V)", command=self.paste).pack(side="left", padx=2)
        ttk.Button(bar, text="📂 Open image…", command=self.open_file).pack(side="left", padx=2)
        ttk.Button(bar, text="⧉ Copy text", command=self.copy).pack(side="left", padx=2)
        ttk.Button(bar, text="Clear", command=self.clear).pack(side="left", padx=2)

        opt = ttk.Frame(root, padding=(6, 0))
        opt.pack(fill="x")
        self.autocopy = tk.BooleanVar(value=True)
        self.split = tk.BooleanVar(value=True)
        self.engine = tk.StringVar(value="auto")
        ttk.Checkbutton(opt, text="Copy automatically", variable=self.autocopy).pack(side="left")
        ttk.Checkbutton(opt, text="Split columns into Excel cells", variable=self.split).pack(side="left", padx=10)
        ttk.Label(opt, text="OCR:").pack(side="left")
        ttk.Combobox(opt, textvariable=self.engine, values=["auto", "windows", "tesseract"],
                     width=9, state="readonly").pack(side="left", padx=4)

        panes = ttk.PanedWindow(root, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=6, pady=6)
        left = ttk.LabelFrame(panes, text="Picture", padding=2)
        self.preview = ttk.Label(left, anchor="n")
        self.preview.pack(fill="both", expand=True)
        panes.add(left, weight=1)
        right = ttk.LabelFrame(panes, text="Result (you can edit it)", padding=2)
        self.text = tk.Text(right, wrap="none", font=("Consolas", 11), undo=True)
        ys = ttk.Scrollbar(right, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=ys.set)
        self.text.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        panes.add(right, weight=2)

        self.status = tk.StringVar(value="Copy a picture/screenshot, then press Ctrl+V here.")
        ttk.Label(root, textvariable=self.status, relief="sunken", anchor="w", padding=3).pack(fill="x", side="bottom")

        # Ctrl+V pastes an image when the clipboard holds one; otherwise normal text paste
        root.bind_all("<Control-v>", self.on_ctrl_v)
        root.bind_all("<Control-V>", self.on_ctrl_v)
        root.after(100, self.poll)

    def on_ctrl_v(self, event):
        img = self.clipboard_image()
        if img is not None:
            self.process(img)
            return "break"
        return None  # let the text box paste text normally

    @staticmethod
    def clipboard_image():
        try:
            img = ImageGrab.grabclipboard()
        except Exception:
            return None
        if isinstance(img, list) and img:
            try:
                return Image.open(img[0])
            except Exception:
                return None
        return img if isinstance(img, Image.Image) else None

    def paste(self):
        img = self.clipboard_image()
        if img is None:
            self.status.set("No picture in the clipboard. Copy a screenshot first.")
        else:
            self.process(img)

    def open_file(self):
        p = filedialog.askopenfilename(filetypes=[("Pictures", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp"), ("All files", "*.*")])
        if p:
            self.process(Image.open(p))

    def process(self, img):
        if self.busy:
            return
        self.busy = True
        img = img.convert("RGB")
        im = img.copy()
        im.thumbnail((300, 520))
        self._pv = ImageTk.PhotoImage(im)
        self.preview.configure(image=self._pv)
        self.status.set("Reading text…")
        self.root.config(cursor="watch")
        engine, split = self.engine.get(), self.split.get()

        def work():
            try:
                self.q.put(("ok", ocr_core.image_to_table(img, engine, split)))
            except Exception as e:
                self.q.put(("err", str(e)))
        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                self.busy = False
                self.root.config(cursor="")
                if kind == "err":
                    self.status.set("OCR failed.")
                    messagebox.showerror(APP, data)
                    continue
                rows, used = data
                self.text.delete("1.0", "end")
                self.text.insert("1.0", "\n".join("\t".join(r) for r in rows))
                msg = f"{len(rows)} row(s) read ({used} OCR)."
                if not rows:
                    msg = "No text found in the picture."
                elif self.autocopy.get():
                    self.copy(quiet=True)
                    msg += " Copied — click a cell in Excel and press Ctrl+V."
                self.status.set(msg)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def copy(self, quiet=False):
        txt = self.text.get("1.0", "end").strip("\n")
        if not txt:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(txt.replace("\n", "\r\n"))
        self.root.update()
        if not quiet:
            self.status.set(f"Copied {txt.count(chr(10)) + 1} row(s). Click a cell in Excel and press Ctrl+V.")

    def clear(self):
        self.text.delete("1.0", "end")
        self.preview.configure(image="")
        self.status.set("Cleared.")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()

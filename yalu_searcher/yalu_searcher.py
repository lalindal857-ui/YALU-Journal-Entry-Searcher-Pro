import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import re, sys
import openpyxl

APP_TITLE = "YALU Searcher Pro"

def asset_path(name):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name

class YaluSearcher:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("430x166")
        self.root.resizable(False, False)
        self.root.configure(bg="#081226")
        self.rows, self.col, self.last_copy = [], {}, ""
        try:
            self.root.iconbitmap(str(asset_path("yalu_searcher.ico")))
        except Exception:
            pass
        self._ui()
        self.root.bind("<Return>", lambda e: self.search_copy())
        self.root.bind("<Control-o>", lambda e: self.open_excel())

    def _ui(self):
        header = tk.Frame(self.root, bg="#081226")
        header.pack(fill="x", padx=10, pady=(7, 4))
        try:
            self.logo = tk.PhotoImage(file=str(asset_path("yalu_logo.png")))
            tk.Label(header, image=self.logo, bg="#081226", bd=0).pack(side="left", padx=(0, 8))
        except Exception:
            tk.Label(header, text="JE", bg="#0756d8", fg="white", font=("Segoe UI", 16, "bold"), width=4).pack(side="left", padx=(0, 8))
        titlebox = tk.Frame(header, bg="#081226")
        titlebox.pack(side="left", fill="y", expand=True)
        tk.Label(titlebox, text="YALU SEARCHER PRO", bg="#081226", fg="white", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(7, 0))
        tk.Label(titlebox, text="Journal Entry Searcher", bg="#081226", fg="#91a9ce", font=("Segoe UI", 8)).pack(anchor="w")
        tk.Label(header, text="PRO", bg="#0756d8", fg="white", font=("Segoe UI", 8, "bold"), padx=7, pady=3).pack(side="right", pady=12)

        row = tk.Frame(self.root, bg="#081226")
        row.pack(fill="x", padx=10, pady=3)
        tk.Button(row, text="📂", command=self.open_excel, bg="#0756d8", fg="white", activebackground="#0646b3", relief="flat", width=4, font=("Segoe UI", 10, "bold"), cursor="hand2").pack(side="left")
        self.entry = tk.Entry(row, justify="center", bg="#172844", fg="white", insertbackground="white", relief="flat", font=("Segoe UI", 12), width=19)
        self.entry.pack(side="left", fill="x", expand=True, padx=7, ipady=6)
        self.entry.insert(0, "Entry No.")
        self.entry.bind("<FocusIn>", self._clear_hint)
        tk.Button(row, text="COPY", command=self.search_copy, bg="#087f68", fg="white", activebackground="#066653", relief="flat", width=7, font=("Segoe UI", 9, "bold"), cursor="hand2").pack(side="right")

        bottom = tk.Frame(self.root, bg="#081226")
        bottom.pack(fill="x", padx=11, pady=(5, 0))
        self.cost_var = tk.StringVar(value="Cost Center: —")
        tk.Label(bottom, textvariable=self.cost_var, bg="#081226", fg="#b5c9e8", font=("Segoe UI", 8), anchor="w").pack(side="left", fill="x", expand=True)
        self.status_var = tk.StringVar(value="Open Excel report")
        tk.Label(bottom, textvariable=self.status_var, bg="#081226", fg="#40c9a2", font=("Segoe UI", 8), anchor="e").pack(side="right")

        tk.Label(self.root, text="Prepared By Lakshan", bg="#081226", fg="#6684ad", font=("Segoe UI", 7)).pack(anchor="e", padx=12, pady=(2, 3))
        menu = tk.Menu(self.root)
        fm = tk.Menu(menu, tearoff=0)
        fm.add_command(label="Open Excel Report", command=self.open_excel, accelerator="Ctrl+O")
        fm.add_command(label="Show last copied text", command=self.show_last)
        menu.add_cascade(label="File", menu=fm)
        self.root.config(menu=menu)

    def _clear_hint(self, _):
        if self.entry.get() == "Entry No.": self.entry.delete(0, "end")

    @staticmethod
    def clean(v): return "" if v is None else str(v).strip()

    @staticmethod
    def is_header(v): return bool(re.match(r"^JE-", v.strip(), re.I))

    def open_excel(self):
        path = filedialog.askopenfilename(title="Select Journal Entries report", filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if not path: return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if "Journal Entries" not in wb.sheetnames: raise ValueError("Sheet 'Journal Entries' not found.")
            ws = wb["Journal Entries"]
            headers = [self.clean(v) for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
            required = ["Entry No. / Line", "Posting Date / Cost Center", "Description"]
            missing = [h for h in required if h not in headers]
            if missing: raise ValueError("Missing columns: " + ", ".join(missing))
            self.col = {h: headers.index(h) for h in required}
            self.rows = list(ws.iter_rows(min_row=2, values_only=True))
            wb.close()
            self.status_var.set(f"Loaded {len(self.rows):,} rows")
            self.entry.focus_set()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))

    def search_copy(self):
        q = self.entry.get().strip()
        if not self.rows:
            messagebox.showinfo(APP_TITLE, "Open an Excel report first."); return
        if not q or q == "Entry No.": return
        ei, ci, di = (self.col["Entry No. / Line"], self.col["Posting Date / Cost Center"], self.col["Description"])
        start, full_entry = None, ""
        for i, r in enumerate(self.rows):
            val = self.clean(r[ei])
            if self.is_header(val) and (val.casefold() == q.casefold() or val.casefold().endswith("-" + q.casefold()) or val.casefold().endswith(q.casefold())):
                start, full_entry = i, val; break
        if start is None:
            self.status_var.set("Not found"); self.cost_var.set("Cost Center: —"); return
        output, centers = [], []
        for r in self.rows[start+1:]:
            raw = self.clean(r[ei])
            if self.is_header(raw): break
            if not re.fullmatch(r"\d+", raw): continue
            desc, center = self.clean(r[di]), self.clean(r[ci])
            if desc: output.append(f"{desc} ({full_entry} / {raw})")
            if center and center not in centers: centers.append(center)
        if not output:
            self.status_var.set("No detail lines"); return
        # Keep the copied text clean; show the author credit in a separate mini notification.
        self.last_copy = "\n".join(output)
        self.root.clipboard_clear(); self.root.clipboard_append(self.last_copy); self.root.update()
        self.cost_var.set("Cost Center: " + (", ".join(centers) if centers else "—"))
        self.status_var.set("Copied ✓"); self.root.bell()
        self.show_prepared_notification()

    def show_prepared_notification(self):
        # Small, temporary Windows-style toast separate from the clipboard contents.
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.configure(bg="#0756d8")
        width, height = 238, 66
        x = self.root.winfo_screenwidth() - width - 22
        y = self.root.winfo_screenheight() - height - 70
        toast.geometry(f"{width}x{height}+{x}+{y}")
        panel = tk.Frame(toast, bg="#101d35", highlightbackground="#1677e8", highlightthickness=1)
        panel.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(panel, text="✓  COPIED SUCCESSFULLY", bg="#101d35", fg="#52d6b0",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=11, pady=(8, 1))
        tk.Label(panel, text="Prepared By Lakshan", bg="#101d35", fg="white",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=11)
        toast.after(2600, toast.destroy)

    def show_last(self):
        if self.last_copy: messagebox.showinfo("Copied Description", self.last_copy)
        else: messagebox.showinfo(APP_TITLE, "No search result yet.")

if __name__ == "__main__":
    root = tk.Tk(); YaluSearcher(root); root.mainloop()

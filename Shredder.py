import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES

def shred(path):
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names
    sheet = sheets[1] if len(sheets) > 1 else sheets[0]
    df = pd.read_excel(path, sheet_name=sheet)
    base = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join(os.path.expanduser("~"), "Desktop", base)
    os.makedirs(out, exist_ok=True)
    for i in range(0, len(df), 999):
        chunk = df.iloc[i : i + 999]
        start, end = i + 1, min(i + 999, len(df))
        name = f"{base}-{start}-{end}.csv"
        chunk.to_csv(os.path.join(out, name), index=False)
    print(f"{os.path.basename(path)} → split on '{sheet}'")

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Shredder")
        self.geometry("800x500")
        self.configure(bg="#001f3f")
        self.resizable(False, False)

        drop = tk.Frame(self, bg="#001f3f")
        drop.pack(fill="both", expand=True)
        drop.drop_target_register(DND_FILES)
        drop.dnd_bind("<<Drop>>", self._on_drop)

        btn_cfg = {
            "bg": "#001f3f",
            "fg": "white",
            "activebackground": "#001f3f",
            "activeforeground": "white",
            "bd": 0,
            "highlightthickness": 0,
            "relief": "flat",
            "takefocus": 0,
        }
        tk.Button(self, text="Browse", command=self._on_browse, **btn_cfg).pack(side="left", padx=10, pady=10)
        tk.Button(self, text="Exit",   command=self.destroy,    **btn_cfg).pack(side="right", padx=10, pady=10)

    def _on_drop(self, e):
        for f in self.tk.splitlist(e.data):
            shred(f)

    def _on_browse(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel", "*.xls;*.xlsx")])
        for f in files:
            shred(f)

if __name__ == "__main__":
    App().mainloop()

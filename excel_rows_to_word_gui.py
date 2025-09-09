# -*- coding: utf-8 -*-
"""
Excel → mikro-tabele w Wordzie (GUI, Tkinter) z mapowaniem pól, kolejnością i transformacjami.
Autor: GPT-5 Thinking (dla Jakuba)

Funkcje:
- wybór pliku Excel (.xlsx / .xls), folderu i nazwy .docx
- edytowalna mapa pól: etykieta (Nowa nazwa), źródło (Befintlig), transformacja, stała wartość, włączenie
- zmiana kolejności wierszy (↑/↓)
- zakładka "Pozostałe kolumny": checkboxy (domyślnie wyłączone)
- ramka na zdjęcie + opcjonalny page break
- opcja użycia przecinka dziesiętnego
"""

# ---------- BOOTSTRAP PIP ----------
import sys, subprocess
def _pip_install(spec):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-U", spec])
    except Exception:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "-U", spec])

def _ensure_min(pkg, module, min_v):
    try:
        try:
            from importlib.metadata import version as _v
        except Exception:
            from importlib_metadata import version as _v  # backport
        cur = _v(module)
    except Exception:
        cur = None
    def _lt(a,b):
        pa=lambda s:[int(x) for x in str(s).split(".") if x.isdigit()]
        return pa(a or "0")<pa(b)
    if cur is None or _lt(cur, min_v):
        _pip_install(f"{pkg}>={min_v}")

def _ensure(pkg, module=None):
    import importlib.util
    if not importlib.util.find_spec(module or pkg.replace("-","_")):
        _pip_install(pkg)

_ensure("pandas","pandas")
_ensure("python-docx","docx")
_ensure("openpyxl","openpyxl")
_ensure_min("xlrd","xlrd","2.0.1")

# ---------- IMPORTY ----------
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

EMU_PER_CM = 360000  # Word EMU -> cm

# ---------- TRANSFORMACJE ----------
def tf_identity(val, row, comma=False):
    return "" if val is None else str(val)

def tf_m2_to_ha_round2(val, row, comma=False):
    try:
        ha = round(float(val)/10000.0, 2)
        s = f"{ha:.2f}"
        return s.replace(".", ",") if comma else s
    except Exception:
        return ""
def tf_prelim_to_bedomning(val, row, comma=False):
    v = str(val).strip().lower() if val is not None else ""
    return "Säker" if v in {"0","nej","no","false","0.0",""} or val==0 else "Preliminärt"
def tf_constant(val, row, comma=False, const_val=""):
    return const_val

TRANSFORMS = {
    "identity": ("Bez zmian", tf_identity),
    "m2_to_ha_round2": ("m² → ha (0,01)", tf_m2_to_ha_round2),
    "prelim_to_bedomning": ("0/nej → Säker, inne → Preliminärt", tf_prelim_to_bedomning),
    "constant": ("Stała wartość", tf_constant),
}

# ---------- DOCX HELPERS ----------
def _shade_cell(cell, fill_hex="F2F2F2"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)

def _set_cell_text(cell, text, bold=False, size_pt=10):
    cell.text = "" if text is None else str(text)
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.bold = bold
            r.font.size = Pt(size_pt)
    if cell.paragraphs:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

# ---------- CZYTANIE EXCEL ----------
def read_excel_any(path):
    ext = os.path.splitext(path.lower())[1]
    if ext == ".xls":
        return pd.read_excel(path, engine="xlrd")
    return pd.read_excel(path)

# ---------- DOMYŚLNA MAPA (Twoja tabela) ----------
def default_mapping(df_cols):
    # pomocnicza: sprawdź czy kolumna istnieje
    has = set(map(str, df_cols))
    def col(name): return name if name in has else ""  # brak → puste (dozwolone przy constant)
    return [
        # order wg Twojej kolumny "Kolejnosc"
        {"enabled": True, "label": "Naturvärdesbiotop",           "source": col("objektnummer"),             "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Naturvärdesklass",            "source": col("naturvardesklass"),         "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Areal (ha)",                  "source": col("Shape__Area"),              "transform": "m2_to_ha_round2",      "const": ""},
        {"enabled": True, "label": "Naturtyp",                    "source": col("naturtyp"),                 "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Biotoptyp",                   "source": "",                              "transform": "constant",             "const": ""},
        {"enabled": True, "label": "Hydrologisk huvudgrupp",      "source": col("hydromorfologiskTyp"),      "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Natura 2000-naturtyp",        "source": "",                              "transform": "constant",             "const": ""},
        {"enabled": True, "label": "Beskrivning",                 "source": col("objektbeskrivning"),        "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Biotopvärde",                 "source": col("biotopvarden"),             "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Tidigare kända värdearter",   "source": col("vardearterKandaTidigare"),  "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Inventerade värdearter",      "source": col("vardearterObserverade"),    "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Invasiva främmande arter",    "source": col("invasivaFrammandeArter"),   "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Artvärde",                    "source": col("artvarden"),                "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Motivering till naturvärdesklass","source":"",                           "transform": "constant",             "const": ""},
        {"enabled": True, "label": "Datum för fältbesök",         "source": col("datumForObjektavgr"),       "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Inventerare",                 "source": col("utforare"),                 "transform": "identity",             "const": ""},
        {"enabled": True, "label": "Bedömning",                   "source": col("preliminarAvgransning"),    "transform": "prelim_to_bedomning",  "const": ""},
    ]

# ---------- GENERACJA DOCX Z MAPĄ ----------
def build_docx(df, mapping_rows, extra_cols, out_path, out_name,
               add_photo=True, photo_h_cm=6.0, page_break=False,
               use_decimal_comma=True):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    content_w_cm = (section.page_width - section.left_margin - section.right_margin) / EMU_PER_CM
    left_w_cm, right_w_cm = 6.0, max(6.0, content_w_cm - 6.0)

    # zmapuj wybrane "pozostałe kolumny" jako mapping identity (na końcu)
    for col_name in extra_cols:
        mapping_rows.append({"enabled": True, "label": col_name, "source": col_name, "transform": "identity", "const": ""})

    # generuj
    for _, row in df.iterrows():
        # tabela właściwa
        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.style = "Table Grid"
        table.autofit = False

        for item in mapping_rows:
            if not item.get("enabled", False):
                continue
            lbl = item.get("label","")
            src = item.get("source","")
            tf_key = item.get("transform","identity")
            const_val = item.get("const","")

            # wylicz wartość
            val_raw = row.get(src, "") if src else ""
            # wybierz funkcję
            tf = TRANSFORMS.get(tf_key, TRANSFORMS["identity"])[1]
            if tf_key == "constant":
                val = tf(val_raw, row, use_decimal_comma, const_val=const_val)
            else:
                val = tf(val_raw, row, use_decimal_comma)

            # dodaj wiersz
            tr = table.add_row().cells
            _set_cell_text(tr[0], lbl, bold=True)
            _shade_cell(tr[0], "EAEAEA")
            _set_cell_text(tr[1], val)

        # szerokości kolumn
        for r in table.rows:
            r.cells[0].width = Cm(left_w_cm)
            r.cells[1].width = Cm(right_w_cm)

        doc.add_paragraph("")
        if add_photo:
            p = doc.add_paragraph("Representativt foto (infoga här):")
            if p.runs:
                p.runs[0].font.size = Pt(10)
            ph = doc.add_table(rows=1, cols=1)
            ph.style = "Table Grid"
            ph.autofit = False
            ph.rows[0].height = Cm(photo_h_cm)
            ph.cell(0, 0).width = Cm(content_w_cm)
            _set_cell_text(ph.cell(0, 0), " ")
            doc.add_paragraph("")
        if page_break:
            doc.add_page_break()

    out_file = os.path.join(out_path, out_name if out_name.lower().endswith(".docx") else out_name + ".docx")
    doc.save(out_file)
    return out_file

# ---------- GUI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel → Mikro-tabele Word (mapowanie)")
        self.geometry("980x720")
        self.resizable(True, True)
        self.df = None
        self.mapping = []      # lista dictów
        self.extra_vars = {}   # nazwa kolumny -> tk.BooleanVar()
        self._build_ui()

    # UI
    def _build_ui(self):
        pad = {"padx":8, "pady":6}
        top = ttk.Frame(self); top.pack(fill="both", expand=True)
        # Plik / folder / nazwa
        ttk.Label(top, text="Plik Excel (.xlsx / .xls):").grid(row=0, column=0, sticky="w", **pad)
        self.in_entry = ttk.Entry(top, width=80); self.in_entry.grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(top, text="Wybierz…", command=self.pick_excel).grid(row=0, column=2, **pad)
        ttk.Label(top, text="Folder wyjściowy:").grid(row=1, column=0, sticky="w", **pad)
        self.out_entry = ttk.Entry(top, width=80); self.out_entry.grid(row=1, column=1, sticky="we", **pad)
        ttk.Button(top, text="Wybierz…", command=self.pick_outdir).grid(row=1, column=2, **pad)
        ttk.Label(top, text="Nazwa pliku .docx:").grid(row=2, column=0, sticky="w", **pad)
        self.name_entry = ttk.Entry(top, width=40); self.name_entry.insert(0, "wynik.docx"); self.name_entry.grid(row=2, column=1, sticky="w", **pad)

        # Notebook: Mapowanie / Pozostałe kolumny / Opcje
        nb = ttk.Notebook(top); nb.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=8, pady=8)
        top.rowconfigure(3, weight=1); top.columnconfigure(1, weight=1)

        # --- Tab 1: Mapowanie
        self.tab_map = ttk.Frame(nb); nb.add(self.tab_map, text="Mapowanie pól")
        self._build_map_tab(self.tab_map)

        # --- Tab 2: Pozostałe kolumny
        self.tab_extra = ttk.Frame(nb); nb.add(self.tab_extra, text="Pozostałe kolumny")
        self.extra_box = ttk.LabelFrame(self.tab_extra, text="Kolumny (domyślnie wyłączone)")
        self.extra_box.pack(fill="both", expand=True, padx=8, pady=8)

        # --- Opcje
        self.tab_opt = ttk.Frame(nb); nb.add(self.tab_opt, text="Opcje")
        self._build_opt_tab(self.tab_opt)

        # Start
        ttk.Button(top, text="Wczytaj i załaduj mapę", command=self.load_columns).grid(row=4, column=0, sticky="w", **pad)
        ttk.Button(top, text="Generuj DOCX", command=self.run).grid(row=4, column=1, sticky="e", **pad)

    def _build_map_tab(self, parent):
        # tabela
        cols = ("order","#on","label","source","transform","const")
        self.tree = ttk.Treeview(parent, columns=cols, show="headings", height=12)
        for c, txt, w in [
            ("order","Kolej.",60),
            ("#on","Włącz",60),
            ("label","Nowa nazwa",250),
            ("source","Źródło",220),
            ("transform","Transformacja",220),
            ("const","Stała",180),
        ]:
            self.tree.heading(c, text=txt)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=8, pady=8)

        # edytor wybranego wiersza
        ed = ttk.LabelFrame(parent, text="Edytuj wiersz")
        ed.pack(fill="x", padx=8, pady=(0,8))
        self.var_on = tk.BooleanVar(value=True)
        self.var_label = tk.StringVar()
        self.var_source = tk.StringVar()
        self.var_transform = tk.StringVar(value="identity")
        self.var_const = tk.StringVar()

        ttk.Checkbutton(ed, text="Włączony", variable=self.var_on).grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(ed, text="Nowa nazwa:").grid(row=0, column=1, sticky="e", padx=6)
        self.ent_label = ttk.Entry(ed, textvariable=self.var_label, width=34); self.ent_label.grid(row=0, column=2, sticky="w", padx=6)

        ttk.Label(ed, text="Źródło:").grid(row=1, column=1, sticky="e", padx=6)
        self.cmb_source = ttk.Combobox(ed, textvariable=self.var_source, width=32, values=[], state="readonly")
        self.cmb_source.grid(row=1, column=2, sticky="w", padx=6)

        ttk.Label(ed, text="Transformacja:").grid(row=0, column=3, sticky="e", padx=6)
        self.cmb_transform = ttk.Combobox(ed, textvariable=self.var_transform, width=28,
                                          values=[k for k in TRANSFORMS], state="readonly")
        self.cmb_transform.grid(row=0, column=4, sticky="w", padx=6)

        ttk.Label(ed, text="Stała:").grid(row=1, column=3, sticky="e", padx=6)
        self.ent_const = ttk.Entry(ed, textvariable=self.var_const, width=30); self.ent_const.grid(row=1, column=4, sticky="w", padx=6)

        ttk.Button(ed, text="Zastosuj", command=self.apply_edit).grid(row=0, column=5, rowspan=2, sticky="nsw", padx=6, pady=4)

        # przyciski listy
        btns = ttk.Frame(parent); btns.pack(fill="x", padx=8, pady=(0,8))
        ttk.Button(btns, text="↑", width=3, command=lambda: self.move_selected(-1)).pack(side="left", padx=2)
        ttk.Button(btns, text="↓", width=3, command=lambda: self.move_selected(1)).pack(side="left", padx=2)
        ttk.Button(btns, text="Usuń", command=self.delete_selected).pack(side="left", padx=8)
        ttk.Button(btns, text="Dodaj z kolumny…", command=self.add_from_column).pack(side="left", padx=8)

        self.tree.bind("<<TreeviewSelect>>", self.on_select_row)

    def _build_opt_tab(self, parent):
        opt = ttk.LabelFrame(parent, text="Opcje dokumentu")
        opt.pack(fill="x", padx=8, pady=8)
        self.var_photo = tk.BooleanVar(value=True)
        self.var_photo_h = tk.DoubleVar(value=6.0)
        self.var_break = tk.BooleanVar(value=False)
        self.var_comma = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="Dodaj ramkę na zdjęcie po każdym rekordzie", variable=self.var_photo).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Label(opt, text="Wysokość ramki [cm]:").grid(row=0, column=1, sticky="e")
        ttk.Entry(opt, textvariable=self.var_photo_h, width=6).grid(row=0, column=2, sticky="w", padx=6)
        ttk.Checkbutton(opt, text="Podział strony po każdym rekordzie", variable=self.var_break).grid(row=1, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(opt, text="Użyj przecinka dziesiętnego", variable=self.var_comma).grid(row=1, column=1, sticky="w", padx=8, pady=4)

    # Handlery
    def pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("All","*.*")])
        if p:
            self.in_entry.delete(0, tk.END); self.in_entry.insert(0, p)

    def pick_outdir(self):
        d = filedialog.askdirectory()
        if d:
            self.out_entry.delete(0, tk.END); self.out_entry.insert(0, d)

    def load_columns(self):
        path = self.in_entry.get().strip()
        if not path:
            messagebox.showwarning("Brak pliku","Wskaż plik Excel.")
            return
        try:
            self.df = read_excel_any(path)
        except Exception as e:
            messagebox.showerror("Błąd wczytywania", f"Nie udało się wczytać Excela:\n{e}\n\nJeśli to .xls, wymagany xlrd>=2.0.1.")
            return

        # mapa domyślna
        self.mapping = default_mapping(self.df.columns)
        self.refresh_tree()

        # pozostałe kolumny
        for w in self.extra_box.winfo_children(): w.destroy()
        mapped_sources = {m["source"] for m in self.mapping if m.get("source")}
        self.extra_vars = {}
        for c in self.df.columns:
            if c in mapped_sources:  # już w mapie
                continue
            v = tk.BooleanVar(value=False)
            ttk.Checkbutton(self.extra_box, text=c, variable=v).pack(anchor="w", padx=8, pady=2)
            self.extra_vars[c] = v

        # edytor: wartości listy źródeł
        self.cmb_source.configure(values=[""] + list(self.df.columns))
        messagebox.showinfo("OK", "Załadowano kolumny i domyślną mapę.")

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, m in enumerate(self.mapping):
            self.tree.insert("", "end", iid=str(i),
                             values=(i, "✓" if m.get("enabled",False) else "–",
                                     m.get("label",""), m.get("source",""),
                                     m.get("transform","identity"),
                                     m.get("const","")))

    def on_select_row(self, *_):
        sel = self.tree.selection()
        if not sel: return
        i = int(sel[0])
        m = self.mapping[i]
        self.var_on.set(bool(m.get("enabled",False)))
        self.var_label.set(m.get("label",""))
        self.var_source.set(m.get("source",""))
        self.var_transform.set(m.get("transform","identity"))
        self.var_const.set(m.get("const",""))

    def apply_edit(self):
        sel = self.tree.selection()
        if not sel: return
        i = int(sel[0])
        self.mapping[i] = {
            "enabled": bool(self.var_on.get()),
            "label": self.var_label.get().strip(),
            "source": self.var_source.get().strip(),
            "transform": self.var_transform.get().strip(),
            "const": self.var_const.get(),
        }
        self.refresh_tree()
        self.tree.selection_set(str(i))

    def move_selected(self, delta):
        sel = self.tree.selection()
        if not sel: return
        i = int(sel[0])
        j = i + delta
        if j < 0 or j >= len(self.mapping): return
        self.mapping[i], self.mapping[j] = self.mapping[j], self.mapping[i]
        self.refresh_tree()
        self.tree.selection_set(str(j))

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel: return
        i = int(sel[0])
        del self.mapping[i]
        self.refresh_tree()

    def add_from_column(self):
        # dodaj wiersz z aktualnie wybraną kolumną (z listy źródeł w edytorze)
        src = self.var_source.get().strip()
        if not src:
            messagebox.showinfo("Info","Wybierz źródło w polu 'Źródło', potem kliknij 'Dodaj z kolumny…'.")
            return
        self.mapping.append({"enabled": False, "label": src, "source": src, "transform":"identity", "const":""})
        self.refresh_tree()
        self.tree.selection_set(str(len(self.mapping)-1))

    def run(self):
        if self.df is None:
            messagebox.showwarning("Brak danych","Najpierw wczytaj kolumny z Excela.")
            return
        out_dir = self.out_entry.get().strip() or os.path.dirname(self.in_entry.get()) or os.getcwd()
        if not os.path.isdir(out_dir):
            messagebox.showwarning("Folder","Podany folder wyjściowy nie istnieje.")
            return
        out_name = self.name_entry.get().strip() or "wynik.docx"

        # ekstra kolumny zaznaczone
        extra = [k for k,v in self.extra_vars.items() if v.get()]

        try:
            out_file = build_docx(
                self.df,
                mapping_rows=list(self.mapping),  # kopia
                extra_cols=extra,
                out_path=out_dir,
                out_name=out_name,
                add_photo=self.var_photo.get(),
                photo_h_cm=float(self.var_photo_h.get()),
                page_break=self.var_break.get(),
                use_decimal_comma=self.var_comma.get(),
            )
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wygenerować DOCX:\n{e}")
            return

        messagebox.showinfo("Gotowe", f"Zapisano:\n{out_file}")

if __name__ == "__main__":
    App().mainloop()

from __future__ import annotations

import platform
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from line_repair import collect_docx_from_folder, repair_file
from pdf_to_word import collect_pdf_from_folder, convert_pdf_to_word

APP_NAME = "Gülşen'in Word Araçları"
APP_VERSION = "0.2.2"
BG = "#F8F3F7"
CARD = "#FFFDFE"
PINK = "#B45C82"
PINK_DARK = "#884362"
PINK_SOFT = "#F2DDE7"
LILAC = "#856AA9"
LILAC_SOFT = "#EEE8F5"
TEXT = "#342D32"
MUTED = "#756B71"
BORDER = "#E6DCE2"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("980x750")
        self.minsize(860, 650)
        self.configure(bg=BG)
        self.repair_files: list[Path] = []
        self.pdf_files: list[Path] = []
        self.repair_recursive = tk.BooleanVar(value=False)
        self.pdf_recursive = tk.BooleanVar(value=False)
        self.include_images = tk.BooleanVar(value=False)
        self.repair_status = tk.StringVar(value="Hazır")
        self.pdf_status = tk.StringVar(value="Hazır")
        self.q: queue.Queue = queue.Queue()
        self._styles()
        self._ui()
        self.after(120, self._poll)

    def ff(self):
        return "Segoe UI" if platform.system() == "Windows" else "Helvetica Neue"

    def _styles(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except Exception:
            pass
        s.configure("TNotebook", background=BG, borderwidth=0)
        s.configure("TNotebook.Tab", padding=(18, 11), font=(self.ff(), 10, "bold"))
        s.configure("Treeview", rowheight=30, background=CARD, fieldbackground=CARD, foreground=TEXT)
        s.configure("Treeview.Heading", font=(self.ff(), 9, "bold"))
        s.configure("Pink.Horizontal.TProgressbar", troughcolor=PINK_SOFT, background=PINK)
        s.configure("Lilac.Horizontal.TProgressbar", troughcolor=LILAC_SOFT, background=LILAC)

    def _button(self, parent, text, cmd, primary=False, lilac=False):
        accent = LILAC if lilac else PINK
        return tk.Button(parent, text=text, command=cmd, bg=accent if primary else "white",
                         fg="white" if primary else TEXT, activebackground=accent,
                         activeforeground="white" if primary else TEXT, relief="flat", bd=0,
                         padx=17, pady=10, cursor="hand2", font=(self.ff(), 10, "bold"),
                         highlightbackground=BORDER, highlightthickness=0 if primary else 1)

    def _ui(self):
        head = tk.Frame(self, bg=BG)
        head.pack(fill="x", padx=28, pady=(22, 12))
        tk.Label(head, text="✿", font=("Georgia", 34), fg=PINK, bg=BG).pack(side="left", padx=(0, 12))
        box = tk.Frame(head, bg=BG); box.pack(side="left", fill="x", expand=True)
        tk.Label(box, text=APP_NAME, font=(self.ff(), 24, "bold"), fg=TEXT, bg=BG).pack(anchor="w")
        tk.Label(box, text="Word belgeleri için sade, güvenli ve kolay araçlar", font=(self.ff(), 10), fg=MUTED, bg=BG).pack(anchor="w")
        tk.Label(head, text=f"v{APP_VERSION}", font=(self.ff(), 9), fg=MUTED, bg=BG).pack(side="right", anchor="n", pady=6)

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        a = tk.Frame(nb, bg=CARD); b = tk.Frame(nb, bg=CARD); c = tk.Frame(nb, bg=CARD)
        nb.add(a, text="Satır Onarıcı"); nb.add(b, text="PDF'den Word'e Aktar"); nb.add(c, text="Hakkında")
        self._repair_tab(a); self._pdf_tab(b); self._about(c)

    def _title(self, parent, title, desc):
        tk.Label(parent, text=title, font=(self.ff(), 18, "bold"), fg=TEXT, bg=CARD).pack(anchor="w")
        tk.Label(parent, text=desc, font=(self.ff(), 10), fg=MUTED, bg=CARD, justify="left", wraplength=850).pack(anchor="w", pady=(4, 14))

    def _tree(self, parent):
        f = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        f.pack(fill="both", expand=True)
        t = ttk.Treeview(f, columns=("file", "folder"), show="headings", height=9)
        t.heading("file", text="Dosya"); t.heading("folder", text="Klasör")
        t.column("file", width=330); t.column("folder", width=500)
        y = ttk.Scrollbar(f, orient="vertical", command=t.yview); t.configure(yscrollcommand=y.set)
        y.pack(side="right", fill="y"); t.pack(fill="both", expand=True)
        return t

    def _repair_tab(self, parent):
        w = tk.Frame(parent, bg=CARD); w.pack(fill="both", expand=True, padx=26, pady=24)
        self._title(w, "Satır Onarıcı", "Kopyalama sırasında cümlenin ortasında oluşan yanlış paragraf sonlarını bulur ve birleştirir.")
        n = tk.Frame(w, bg=PINK_SOFT, padx=13, pady=10); n.pack(fill="x", pady=(0, 14))
        tk.Label(n, text="Orijinal dosyaya dokunulmaz.  Sonuç: DosyaAdı_Düzelti.docx", bg=PINK_SOFT, fg=PINK_DARK, font=(self.ff(), 10, "bold")).pack(anchor="w")
        c = tk.Frame(w, bg=CARD); c.pack(fill="x", pady=(0, 10))
        self._button(c, "Dosya Seç", self.choose_repair_files).pack(side="left", padx=(0, 8))
        self._button(c, "Klasör Seç", self.choose_repair_folder).pack(side="left", padx=(0, 8))
        self._button(c, "Listeyi Temizle", self.clear_repair).pack(side="left")
        tk.Checkbutton(c, text="Alt klasörleri de tara", variable=self.repair_recursive, bg=CARD, fg=MUTED, selectcolor=CARD).pack(side="right")
        self.repair_count = tk.Label(w, text="Henüz dosya seçilmedi", bg=CARD, fg=MUTED, font=(self.ff(), 9, "bold")); self.repair_count.pack(anchor="w", pady=(0, 6))
        self.repair_tree = self._tree(w)
        self.repair_progress = ttk.Progressbar(w, style="Pink.Horizontal.TProgressbar"); self.repair_progress.pack(fill="x", pady=(14, 8))
        r = tk.Frame(w, bg=CARD); r.pack(fill="x")
        tk.Label(r, textvariable=self.repair_status, bg=CARD, fg=MUTED).pack(side="left")
        self.repair_btn = self._button(r, "Onar ve Kopyala", self.start_repair, primary=True); self.repair_btn.pack(side="right")

    def _pdf_tab(self, parent):
        w = tk.Frame(parent, bg=CARD); w.pack(fill="both", expand=True, padx=26, pady=24)
        self._title(w, "PDF'den Word'e Aktar", "Metin tabakası olan PDF'leri düzenlenebilir Word belgesine aktarır; kalın, italik, yazı boyutu, renk ve hizalamayı mümkün olduğunca korur.")
        n = tk.Frame(w, bg=LILAC_SOFT, padx=13, pady=10); n.pack(fill="x", pady=(0, 10))
        tk.Label(n, text="PDF'ye dokunulmaz.  Sonuç: DosyaAdı_Aktarma.docx", bg=LILAC_SOFT, fg=TEXT, font=(self.ff(), 10, "bold")).pack(anchor="w")
        o = tk.Frame(w, bg="#FBF9FD", highlightbackground=BORDER, highlightthickness=1, padx=10, pady=9); o.pack(fill="x", pady=(0, 10))
        tk.Checkbutton(o, text="Görselleri de aktar", variable=self.include_images, bg="#FBF9FD", fg=TEXT, selectcolor="#FBF9FD", font=(self.ff(), 10, "bold")).pack(side="left")
        tk.Label(o, text="Varsayılan olarak kapalıdır.", bg="#FBF9FD", fg=MUTED).pack(side="left", padx=(10, 0))
        c = tk.Frame(w, bg=CARD); c.pack(fill="x", pady=(0, 10))
        self._button(c, "PDF Seç", self.choose_pdf_files, lilac=True).pack(side="left", padx=(0, 8))
        self._button(c, "Klasör Seç", self.choose_pdf_folder, lilac=True).pack(side="left", padx=(0, 8))
        self._button(c, "Listeyi Temizle", self.clear_pdf, lilac=True).pack(side="left")
        tk.Checkbutton(c, text="Alt klasörleri de tara", variable=self.pdf_recursive, bg=CARD, fg=MUTED, selectcolor=CARD).pack(side="right")
        self.pdf_count = tk.Label(w, text="Henüz PDF seçilmedi", bg=CARD, fg=MUTED, font=(self.ff(), 9, "bold")); self.pdf_count.pack(anchor="w", pady=(0, 6))
        self.pdf_tree = self._tree(w)
        self.pdf_progress = ttk.Progressbar(w, style="Lilac.Horizontal.TProgressbar"); self.pdf_progress.pack(fill="x", pady=(14, 8))
        r = tk.Frame(w, bg=CARD); r.pack(fill="x")
        tk.Label(r, textvariable=self.pdf_status, bg=CARD, fg=MUTED).pack(side="left")
        self.pdf_btn = self._button(r, "Word'e Aktar", self.start_pdf, primary=True, lilac=True); self.pdf_btn.pack(side="right")

    def _about(self, parent):
        w = tk.Frame(parent, bg=CARD); w.pack(fill="both", expand=True, padx=30, pady=30)
        tk.Label(w, text="✿", font=("Georgia", 46), fg=PINK, bg=CARD).pack(pady=(20, 8))
        tk.Label(w, text=APP_NAME, font=(self.ff(), 21, "bold"), fg=TEXT, bg=CARD).pack()
        tk.Label(w, text="Satır Onarıcı ve PDF'den Word'e Aktar", font=(self.ff(), 11), fg=MUTED, bg=CARD).pack(pady=8)
        tk.Label(w, text="Kaynak dosyalar değiştirilmez; sonuçlar yeni dosya olarak oluşturulur.", font=(self.ff(), 10), fg=MUTED, bg=CARD).pack(pady=4)

    def _fill(self, tree, files):
        tree.delete(*tree.get_children())
        for p in files:
            tree.insert("", "end", values=(p.name, str(p.parent)))

    def choose_repair_files(self):
        xs = filedialog.askopenfilenames(title="Word dosyalarını seç", filetypes=[("Word belgeleri", "*.docx")])
        if xs:
            self.repair_files = [Path(x) for x in xs]; self._refresh_repair()

    def choose_repair_folder(self):
        x = filedialog.askdirectory(title="Word dosyalarının klasörünü seç")
        if x:
            self.repair_files = collect_docx_from_folder(x, self.repair_recursive.get()); self._refresh_repair()

    def _refresh_repair(self):
        self._fill(self.repair_tree, self.repair_files)
        self.repair_count.config(text=f"{len(self.repair_files)} dosya seçildi" if self.repair_files else "Henüz dosya seçilmedi")

    def clear_repair(self):
        self.repair_files = []; self._refresh_repair()

    def choose_pdf_files(self):
        xs = filedialog.askopenfilenames(title="PDF dosyalarını seç", filetypes=[("PDF dosyaları", "*.pdf")])
        if xs:
            self.pdf_files = [Path(x) for x in xs]; self._refresh_pdf()

    def choose_pdf_folder(self):
        x = filedialog.askdirectory(title="PDF dosyalarının klasörünü seç")
        if x:
            self.pdf_files = collect_pdf_from_folder(x, self.pdf_recursive.get()); self._refresh_pdf()

    def _refresh_pdf(self):
        self._fill(self.pdf_tree, self.pdf_files)
        self.pdf_count.config(text=f"{len(self.pdf_files)} PDF seçildi" if self.pdf_files else "Henüz PDF seçilmedi")

    def clear_pdf(self):
        self.pdf_files = []; self._refresh_pdf()

    def start_repair(self):
        if not self.repair_files:
            messagebox.showwarning(APP_NAME, "Önce en az bir Word dosyası seçin."); return
        self.repair_btn.config(state="disabled")
        self.repair_progress["maximum"] = len(self.repair_files); self.repair_progress["value"] = 0
        threading.Thread(target=self._repair_worker, daemon=True).start()

    def _repair_worker(self):
        done = repairs = 0; errors = []
        for i, p in enumerate(self.repair_files, 1):
            self.q.put(("rs", f"İşleniyor: {p.name}"))
            r = repair_file(p)
            if r.skipped: errors.append(f"{p.name}: {r.message}")
            else: done += 1; repairs += r.repaired_breaks
            self.q.put(("rp", i))
        self.q.put(("rd", done, repairs, errors))

    def start_pdf(self):
        if not self.pdf_files:
            messagebox.showwarning(APP_NAME, "Önce en az bir PDF seçin."); return
        self.pdf_btn.config(state="disabled")
        self.pdf_progress["maximum"] = len(self.pdf_files); self.pdf_progress["value"] = 0
        threading.Thread(target=self._pdf_worker, daemon=True).start()

    def _pdf_worker(self):
        done = pages = images = 0; errors = []; warnings = []
        img = self.include_images.get()
        for i, p in enumerate(self.pdf_files, 1):
            self.q.put(("ps", f"Aktarılıyor: {p.name}"))
            r = convert_pdf_to_word(p, img)
            if r.skipped: errors.append(f"{p.name}: {r.message}")
            else: done += 1; pages += r.pages; images += r.images
            if r.warning: warnings.append(f"{p.name}: {r.warning}")
            self.q.put(("pp", i))
        self.q.put(("pd", done, pages, images, errors, warnings))

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait(); k = m[0]
                if k == "rs": self.repair_status.set(m[1])
                elif k == "rp": self.repair_progress["value"] = m[1]
                elif k == "rd":
                    _, done, repairs, errors = m; self.repair_btn.config(state="normal")
                    self.repair_status.set(f"Tamamlandı • {done} dosya • {repairs} onarım")
                    text = f"{done} düzeltilmiş kopya oluşturuldu.\n{repairs} satır sonu onarıldı."
                    if errors: text += "\n\n" + "\n".join(errors[:8])
                    (messagebox.showwarning if errors else messagebox.showinfo)(APP_NAME, text)
                elif k == "ps": self.pdf_status.set(m[1])
                elif k == "pp": self.pdf_progress["value"] = m[1]
                elif k == "pd":
                    _, done, pages, images, errors, warnings = m; self.pdf_btn.config(state="normal")
                    self.pdf_status.set(f"Tamamlandı • {done} belge • {pages} sayfa")
                    text = f"{done} Word belgesi oluşturuldu.\n{pages} PDF sayfası işlendi."
                    if self.include_images.get(): text += f"\n{images} görsel aktarıldı."
                    if warnings: text += "\n\nUyarılar:\n" + "\n".join(warnings[:5])
                    if errors: text += "\n\nSorunlar:\n" + "\n".join(errors[:5])
                    (messagebox.showwarning if errors else messagebox.showinfo)(APP_NAME, text)
        except queue.Empty:
            pass
        self.after(120, self._poll)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()

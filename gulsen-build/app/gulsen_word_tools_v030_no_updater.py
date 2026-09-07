from __future__ import annotations

import platform
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None

from line_repair import collect_docx_from_folder, repair_file
from pdf_to_word import collect_media_from_folder, convert_media_to_word, SUPPORTED_IMAGE_EXTS

APP_NAME = "Gülşen'in Word Araçları"
APP_VERSION = "0.3.0"

BG = "#F8F4F7"
CARD = "#FFFDFE"
PINK = "#B65C82"
PINK_D = "#8E3F64"
PINK_S = "#F2DCE6"
TEXT = "#332B31"
MUTED = "#756A72"
BORDER = "#E8DDE3"

IMG_EXTS = tuple(sorted(SUPPORTED_IMAGE_EXTS))
ROOT = TkinterDnD.Tk if TkinterDnD else tk.Tk


class App(ROOT):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.configure(bg=BG)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        width = min(760, max(600, int(sw * 0.64)))
        height = min(470, max(360, int(sh * 0.58)))
        x = max(0, (sw - width) // 2)
        y = max(0, (sh - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(590, 350)

        self.items: list[Path] = []
        self.action = tk.StringVar(value="")
        self.include_images = tk.BooleanVar(value=False)
        self.recursive = tk.BooleanVar(value=False)
        self.auto_start = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Hazır")
        self.summary = tk.StringVar(value="Liste boş")
        self.progress_pct = tk.StringVar(value="0%")

        self.q: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.action_dialog: tk.Toplevel | None = None
        self._force_start_after_choice = False

        self._style()
        self._ui()
        self.after(120, self._poll)

    def ff(self):
        return "Segoe UI" if platform.system() == "Windows" else "Helvetica Neue"

    def _style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except Exception:
            pass
        s.configure("Treeview", rowheight=22, background=CARD, fieldbackground=CARD, foreground=TEXT)
        s.configure("Treeview.Heading", font=(self.ff(), 9, "bold"))
        s.configure("Work.Horizontal.TProgressbar", troughcolor=PINK_S, background=PINK)

    def _btn(self, parent, text, cmd, primary=False, width=None):
        bg, fg = ((PINK, "white") if primary else ("white", TEXT))
        return tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=bg,
            fg=fg,
            disabledforeground=fg,
            activebackground=PINK_D if primary else "#F5F0F3",
            activeforeground="white" if primary else TEXT,
            relief="flat",
            bd=0,
            padx=9,
            pady=5,
            width=width,
            font=(self.ff(), 9, "bold"),
            highlightbackground=PINK if primary else BORDER,
            highlightthickness=1,
            cursor="hand2",
        )

    def _ui(self):
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=12, pady=(8, 4))
        tk.Label(header, text="✿", font=("Georgia", 18), fg=PINK, bg=BG).pack(side="left", padx=(0, 6))

        title_box = tk.Frame(header, bg=BG)
        title_box.pack(side="left", fill="x", expand=True)
        tk.Label(title_box, text=APP_NAME, font=(self.ff(), 15, "bold"), fg=TEXT, bg=BG).pack(anchor="w")
        tk.Label(title_box, text="Dosyayı bırak • işlemi seç • yeni kopyanı oluştur", font=(self.ff(), 8), fg=MUTED, bg=BG).pack(anchor="w")
        tk.Label(header, text="v" + APP_VERSION, font=(self.ff(), 8), fg=MUTED, bg=BG).pack(side="right", anchor="n", pady=(3, 0))

        body = tk.Frame(self, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        top = tk.Frame(body, bg=CARD)
        top.pack(fill="x", padx=12, pady=(10, 7))
        self.drop = tk.Label(
            top,
            text="Dosya veya klasörü buraya sürükleyip bırak\nDOCX • PDF • JPG • PNG • TIFF • BMP • WEBP",
            justify="center",
            bg="#FBF8FA",
            fg=MUTED,
            font=(self.ff(), 10, "bold"),
            highlightbackground=BORDER,
            highlightthickness=1,
            pady=11,
        )
        self.drop.pack(fill="x")
        if DND_FILES:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self._on_drop)

        opts = tk.Frame(body, bg=CARD)
        opts.pack(fill="x", padx=12, pady=(0, 5))
        tk.Checkbutton(opts, text="Alt klasörleri tara", variable=self.recursive, bg=CARD, fg=MUTED, selectcolor=CARD, font=(self.ff(), 8)).pack(side="left")
        tk.Checkbutton(opts, text="Word'e aktarırken görselleri de ekle", variable=self.include_images, bg=CARD, fg=MUTED, selectcolor=CARD, font=(self.ff(), 8)).pack(side="left", padx=(12, 0))
        tk.Label(opts, text="OCR + yamuk tarama düzeltme", bg=CARD, fg=MUTED, font=(self.ff(), 8)).pack(side="right")

        buttons = tk.Frame(body, bg=CARD)
        buttons.pack(fill="x", padx=12, pady=(0, 6))
        left = tk.Frame(buttons, bg=CARD)
        left.pack(side="left")
        self.file_btn = self._btn(left, "Dosya Seç", self.pick_files, width=10)
        self.file_btn.pack(side="left", padx=(0, 5))
        self.folder_btn = self._btn(left, "Klasör Seç", self.pick_folder, width=10)
        self.folder_btn.pack(side="left", padx=(0, 5))
        self.clear_btn = self._btn(left, "Listeyi Temizle", self.clear_items, width=12)
        self.clear_btn.pack(side="left")

        right = tk.Frame(buttons, bg=CARD)
        right.pack(side="right")
        self.cancel_btn = self._btn(right, "İptal", self.cancel_work, width=7)
        self.cancel_btn.pack(side="right", padx=(6, 0))
        self.run_btn = self._btn(right, "İşlemi Başlat", self.start, primary=True, width=13)
        self.run_btn.pack(side="right")
        self.cancel_btn.config(state="disabled")

        self.action_label = tk.Label(body, text="", bg=CARD, fg=PINK_D, font=(self.ff(), 8, "bold"))
        self.action_label.pack(anchor="w", padx=12, pady=(0, 2))
        tk.Label(body, textvariable=self.summary, bg=CARD, fg=MUTED, font=(self.ff(), 8)).pack(anchor="w", padx=12, pady=(0, 3))

        table_frame = tk.Frame(body, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 7))
        self.tree = ttk.Treeview(table_frame, columns=("tip", "dosya", "klasor"), show="headings", height=5)
        self.tree.heading("tip", text="Tür")
        self.tree.heading("dosya", text="Dosya")
        self.tree.heading("klasor", text="Klasör")
        self.tree.column("tip", width=66, stretch=False, minwidth=60)
        self.tree.column("dosya", width=210, minwidth=120)
        self.tree.column("klasor", width=310, minwidth=130)
        sy = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sy.set)
        sy.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        bottom = tk.Frame(body, bg=CARD)
        bottom.pack(fill="x", padx=12, pady=(0, 9))
        self.prog = ttk.Progressbar(bottom, style="Work.Horizontal.TProgressbar", mode="determinate", maximum=100)
        self.prog.pack(fill="x", pady=(0, 3))
        status_row = tk.Frame(bottom, bg=CARD)
        status_row.pack(fill="x")
        tk.Label(status_row, textvariable=self.status, bg=CARD, fg=MUTED, font=(self.ff(), 8)).pack(side="left")
        tk.Label(status_row, textvariable=self.progress_pct, bg=CARD, fg=PINK_D, font=(self.ff(), 8, "bold")).pack(side="right")

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.run_btn.config(state=state, text="İşlem sürüyor" if busy else "İşlemi Başlat")
        self.file_btn.config(state=state)
        self.folder_btn.config(state=state)
        self.clear_btn.config(state=state)
        self.cancel_btn.config(state="normal" if busy else "disabled")

    def _center_dialog(self, dialog: tk.Toplevel):
        self.update_idletasks()
        dialog.update_idletasks()
        px, py = self.winfo_rootx(), self.winfo_rooty()
        pw, ph = max(1, self.winfo_width()), max(1, self.winfo_height())
        dw, dh = dialog.winfo_reqwidth(), dialog.winfo_reqheight()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        margin = 12
        x = min(max(margin, px + (pw - dw) // 2), max(margin, sw - dw - margin))
        y = min(max(margin, py + (ph - dh) // 2), max(margin, sh - dh - margin))
        dialog.geometry(f"+{x}+{y}")

    def _close_action_dialog(self):
        d = self.action_dialog
        self.action_dialog = None
        self._force_start_after_choice = False
        if d and d.winfo_exists():
            d.destroy()

    def _prompt_action(self, force_start_after_choice=False):
        self._force_start_after_choice = bool(force_start_after_choice)
        if self.action_dialog and self.action_dialog.winfo_exists():
            self._center_dialog(self.action_dialog)
            self.action_dialog.lift()
            self.action_dialog.focus_force()
            return

        dialog = tk.Toplevel(self)
        self.action_dialog = dialog
        dialog.title("İşlem Seç")
        dialog.configure(bg=BG)
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.protocol("WM_DELETE_WINDOW", self._close_action_dialog)

        fr = tk.Frame(dialog, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        fr.pack(fill="both", expand=True, padx=8, pady=8)
        tk.Label(fr, text="Ne yapmak istiyorsun?", bg=CARD, fg=TEXT, font=(self.ff(), 10, "bold")).pack(anchor="w", padx=10, pady=(9, 3))
        tk.Label(fr, text="Word: satır düzeltme • PDF/görsel: Word'e aktarım", bg=CARD, fg=MUTED, font=(self.ff(), 8)).pack(anchor="w", padx=10, pady=(0, 5))
        tk.Checkbutton(
            fr,
            text="Seçimden sonra otomatik başlat",
            variable=self.auto_start,
            bg=CARD,
            fg=TEXT,
            activebackground=CARD,
            selectcolor=CARD,
            font=(self.ff(), 8, "bold"),
        ).pack(anchor="w", padx=10, pady=(0, 8))
        row = tk.Frame(fr, bg=CARD)
        row.pack(fill="x", padx=10, pady=(0, 9))

        def choose(value: str):
            should_start = self._force_start_after_choice or self.auto_start.get()
            self.action.set(value)
            self._force_start_after_choice = False
            if dialog.winfo_exists():
                dialog.destroy()
            self.action_dialog = None
            self._refresh()
            if should_start:
                self.after(20, self.start)

        self._btn(row, "Satırları Düzelt", lambda: choose("repair"), width=15).pack(side="left", padx=(0, 7))
        self._btn(row, "Word'e Aktar", lambda: choose("convert"), primary=True, width=15).pack(side="left")
        dialog.update_idletasks()
        self._center_dialog(dialog)
        dialog.lift()
        dialog.focus_force()

    def clear_items(self):
        if self.worker and self.worker.is_alive():
            return
        self.items = []
        self.action.set("")
        self._refresh()
        self.status.set("Hazır")
        self.prog["value"] = 0
        self.progress_pct.set("0%")

    def _split_drop(self, data: str):
        try:
            return list(self.tk.splitlist(data))
        except Exception:
            return [p.strip("{}") for p in data.split() if p]

    def _on_drop(self, event):
        if not (self.worker and self.worker.is_alive()):
            self.add_paths(self._split_drop(event.data), ask_action=True)

    def pick_files(self):
        if self.worker and self.worker.is_alive():
            return
        sel = filedialog.askopenfilenames(filetypes=[
            ("Tüm desteklenenler", "*.docx *.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
            ("Word", "*.docx"),
            ("PDF ve Görseller", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
        ])
        if sel:
            self.add_paths(sel, ask_action=True)

    def pick_folder(self):
        if self.worker and self.worker.is_alive():
            return
        d = filedialog.askdirectory()
        if d:
            self.add_paths([d], ask_action=True)

    def add_paths(self, paths, ask_action=False):
        new: list[Path] = []
        for raw in paths:
            p = Path(str(raw).strip().strip("{}")).expanduser()
            if not p.exists():
                continue
            if p.is_dir():
                found = collect_docx_from_folder(p, self.recursive.get()) + collect_media_from_folder(p, self.recursive.get())
                for x in found:
                    if x not in new and x not in self.items:
                        new.append(x)
            else:
                suf = p.suffix.lower()
                allowed = (suf == ".docx" and not p.stem.endswith("_Düzelti")) or suf == ".pdf" or suf in IMG_EXTS
                if allowed and p not in new and p not in self.items:
                    new.append(p)
        if not new:
            messagebox.showinfo(APP_NAME, "Desteklenen bir dosya bulunamadı.")
            return
        self.items.extend(new)
        self._refresh()
        if ask_action:
            self._prompt_action(force_start_after_choice=False)

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for p in self.items:
            kind = "Word" if p.suffix.lower() == ".docx" else ("PDF" if p.suffix.lower() == ".pdf" else "Görsel")
            self.tree.insert("", "end", values=(kind, p.name, str(p.parent)))
        action_text = {"repair": "Seçili işlem: Satırları Düzelt", "convert": "Seçili işlem: Word'e Aktar"}.get(self.action.get(), "")
        self.action_label.config(text=action_text)
        self.summary.set(f"{len(self.items)} öge listede" if self.items else "Liste boş")

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.items:
            messagebox.showinfo(APP_NAME, "Önce dosya veya klasör seçin.")
            return
        action = self.action.get()
        if action not in {"repair", "convert"}:
            self._prompt_action(force_start_after_choice=True)
            return

        if action == "repair":
            targets = [p for p in self.items if p.suffix.lower() == ".docx" and not p.stem.endswith("_Düzelti")]
            if not targets:
                messagebox.showinfo(APP_NAME, "Satır düzeltme için en az bir .docx dosyası seçin.")
                return
        else:
            targets = [p for p in self.items if p.suffix.lower() == ".pdf" or p.suffix.lower() in IMG_EXTS]
            if not targets:
                messagebox.showinfo(APP_NAME, "Word'e aktarım için PDF veya görsel seçin.")
                return

        self.cancel_event.clear()
        self.prog["value"] = 0
        self.progress_pct.set("0%")
        self._set_busy(True)
        self.status.set("İşlem başladı…")
        self.worker = threading.Thread(target=self._work, args=(action, targets, self.include_images.get()), daemon=True)
        self.worker.start()

    def cancel_work(self):
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
            self.status.set("İptal ediliyor…")
            self.cancel_btn.config(state="disabled")

    def _work(self, action, targets, include_images):
        done = repairs = pages = images = ocr_pages = 0
        errors = []
        cancelled = False
        total_targets = max(1, len(targets))

        for index, p in enumerate(targets, 1):
            if self.cancel_event.is_set():
                cancelled = True
                break

            base = ((index - 1) / total_targets) * 100.0
            span = 100.0 / total_targets

            def file_progress(local_pct, message=""):
                overall = base + span * (max(0.0, min(100.0, float(local_pct))) / 100.0)
                self.q.put(("progress", overall, f"{index}/{total_targets} • {message or p.name}"))

            file_progress(1, f"Başlıyor: {p.name}")
            success = False
            try:
                if action == "repair":
                    file_progress(15, "Satırlar inceleniyor…")
                    r = repair_file(p)
                    if r.skipped:
                        errors.append(f"{p.name}: {r.message}")
                    else:
                        done += 1
                        repairs += r.repaired_breaks
                        success = True
                    file_progress(100, "Dosya tamamlandı")
                else:
                    r = convert_media_to_word(
                        p,
                        include_images=include_images,
                        cancel_event=self.cancel_event,
                        progress_callback=file_progress,
                    )
                    if getattr(r, "cancelled", False) or self.cancel_event.is_set():
                        cancelled = True
                        break
                    if r.skipped:
                        errors.append(f"{p.name}: {r.message}")
                    else:
                        done += 1
                        pages += r.pages
                        images += r.images
                        ocr_pages += getattr(r, "ocr_pages", 0)
                        success = True
                    file_progress(100, "Dosya tamamlandı")
            except Exception as exc:
                errors.append(f"{p.name}: {exc}")
                file_progress(100, "Dosya atlandı")

            if success:
                self.q.put(("file_done", p))

        if not cancelled and not self.cancel_event.is_set():
            self.q.put(("progress", 100.0, "İşlem tamamlandı"))
        self.q.put(("done", action, done, repairs, pages, images, ocr_pages, errors, cancelled))

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                kind = m[0]
                if kind == "progress":
                    _, value, text = m
                    value = max(0.0, min(100.0, float(value)))
                    self.prog["value"] = value
                    self.progress_pct.set(f"{int(round(value))}%")
                    if text:
                        self.status.set(text)
                elif kind == "file_done":
                    _, path = m
                    self.items = [item for item in self.items if item != path]
                    self._refresh()
                elif kind == "done":
                    self._set_busy(False)
                    _, action, done, repairs, pages, images, ocr_pages, errors, cancelled = m
                    if cancelled:
                        msg = "İşlem iptal edildi."
                    elif action == "repair":
                        msg = f"Tamamlandı • {done} dosya • {repairs} onarım"
                    else:
                        msg = f"Tamamlandı • {done} Word dosyası • {pages} sayfa • {ocr_pages} OCR"
                        if self.include_images.get():
                            msg += f" • {images} görsel"
                    if errors:
                        msg += f"\n\nListede kalan/atlanan dosya: {len(errors)}"
                    self.status.set(msg.split("\n")[0])
                    messagebox.showinfo(APP_NAME, msg)
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _on_close(self):
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
        try:
            if self.action_dialog and self.action_dialog.winfo_exists():
                self.action_dialog.destroy()
        except Exception:
            pass
        self.destroy()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()

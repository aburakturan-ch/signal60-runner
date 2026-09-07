from __future__ import annotations

import os
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
from updater import check_for_update, download_and_stage, install_staged

APP_NAME = "Gülşen'in Word Araçları"
APP_VERSION = "0.2.7"

BG = "#F8F4F7"
CARD = "#FFFDFE"
PINK = "#B65C82"
PINK_D = "#8E3F64"
PINK_S = "#F2DCE6"
TEXT = "#332B31"
MUTED = "#756A72"
BORDER = "#E8DDE3"
GREEN = "#4F7D66"
GREEN_D = "#365B49"
GREEN_S = "#E4F0E9"

IMG_EXTS = tuple(sorted(SUPPORTED_IMAGE_EXTS))
ROOT = TkinterDnD.Tk if TkinterDnD else tk.Tk


class App(ROOT):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.configure(bg=BG)

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        width = min(760, max(600, int(sw * 0.62)))
        height = min(465, max(350, int(sh * 0.56)))
        x = max(0, (sw - width) // 2)
        y = max(0, (sh - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(590, 340)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.items: list[Path] = []
        self.action = tk.StringVar(value="")
        self.include_images = tk.BooleanVar(value=False)
        self.recursive = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Hazır")
        self.summary = tk.StringVar(value="Henüz dosya veya klasör seçilmedi")
        self.q: queue.Queue = queue.Queue()
        self.update_info = None
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.action_dialog: tk.Toplevel | None = None
        self.closing = False

        self._style()
        self._ui()
        self.after(120, self._poll)
        self.after(1400, lambda: threading.Thread(target=self._check_update_worker, args=(False,), daemon=True).start())

    def ff(self):
        return "Segoe UI" if platform.system() == "Windows" else "Helvetica Neue"

    def _style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("aqua" if platform.system() == "Darwin" else "clam")
        except Exception:
            pass
        s.configure("Treeview", rowheight=21, background=CARD, fieldbackground=CARD, foreground=TEXT)
        s.configure("Treeview.Heading", font=(self.ff(), 8, "bold"))
        s.configure("P.Horizontal.TProgressbar", troughcolor=PINK_S, background=PINK)

    def _btn(self, parent, text, command, *, primary=False, green=False, width=None):
        mac = platform.system() == "Darwin"
        accent = GREEN if green else PINK
        dark = GREEN_D if green else PINK_D
        soft = GREEN_S if green else PINK_S
        if primary:
            bg, fg = (soft, dark) if mac else (accent, "white")
            active_fg = dark if mac else "white"
        else:
            bg, fg = "white", TEXT
            active_fg = TEXT
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=soft if not primary else dark,
            activeforeground=active_fg,
            relief="flat",
            bd=0,
            padx=9,
            pady=5,
            width=width,
            font=(self.ff(), 8, "bold"),
            highlightbackground=accent if primary else BORDER,
            highlightthickness=1,
            cursor="hand2",
        )

    def _ui(self):
        h = tk.Frame(self, bg=BG)
        h.pack(fill="x", padx=12, pady=(7, 4))
        tk.Label(h, text="✿", font=("Georgia", 18), fg=PINK, bg=BG).pack(side="left", padx=(0, 6))

        hb = tk.Frame(h, bg=BG)
        hb.pack(side="left", fill="x", expand=True)
        tk.Label(hb, text=APP_NAME, font=(self.ff(), 14, "bold"), fg=TEXT, bg=BG).pack(anchor="w")
        tk.Label(hb, text="Dosyayı bırak • işlemi seç • yeni kopyanı oluştur", font=(self.ff(), 8), fg=MUTED, bg=BG).pack(anchor="w")

        hr = tk.Frame(h, bg=BG)
        hr.pack(side="right")
        tk.Label(hr, text="v" + APP_VERSION, font=(self.ff(), 7), fg=MUTED, bg=BG).pack(anchor="e", pady=(0, 2))
        self.update_btn = self._btn(hr, "Güncelle", self.update_clicked, green=True, width=10)
        self.update_btn.pack(anchor="e")

        body = tk.Frame(self, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 9))

        top = tk.Frame(body, bg=CARD)
        top.pack(fill="x", padx=11, pady=(9, 6))
        self.drop = tk.Label(
            top,
            text="Dosya veya klasörü buraya sürükleyip bırak\nDOCX • PDF • JPG • PNG • TIFF • BMP • WEBP",
            justify="center",
            bg="#FBF8FA",
            fg=MUTED,
            font=(self.ff(), 9, "bold"),
            highlightbackground=BORDER,
            highlightthickness=1,
            pady=10,
        )
        self.drop.pack(fill="x")
        if DND_FILES:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self._on_drop)

        opts = tk.Frame(body, bg=CARD)
        opts.pack(fill="x", padx=11, pady=(0, 5))
        tk.Checkbutton(opts, text="Alt klasörleri tara", variable=self.recursive, bg=CARD, fg=MUTED, selectcolor=CARD, font=(self.ff(), 8)).pack(side="left")
        tk.Checkbutton(opts, text="Word'e aktarırken görselleri de ekle", variable=self.include_images, bg=CARD, fg=MUTED, selectcolor=CARD, font=(self.ff(), 8)).pack(side="left", padx=(12, 0))
        tk.Label(opts, text="OCR + yamuk tarama düzeltme", bg=CARD, fg=MUTED, font=(self.ff(), 7)).pack(side="right")

        note = tk.Frame(body, bg="#FBF9FD", highlightbackground=BORDER, highlightthickness=1)
        note.pack(fill="x", padx=11, pady=(0, 6))
        tk.Label(
            note,
            text="Kaynak değişmez.  Satır: _Düzelti.docx  •  Aktarım: _Aktarma.docx",
            bg="#FBF9FD",
            fg=MUTED,
            font=(self.ff(), 8),
            padx=8,
            pady=5,
        ).pack(anchor="w")

        buttons = tk.Frame(body, bg=CARD)
        buttons.pack(fill="x", padx=11, pady=(0, 5))
        left = tk.Frame(buttons, bg=CARD)
        left.pack(side="left")
        self.file_btn = self._btn(left, "Dosya Seç", self.pick_files, width=10)
        self.file_btn.pack(side="left", padx=(0, 5))
        self.folder_btn = self._btn(left, "Klasör Seç", self.pick_folder, width=10)
        self.folder_btn.pack(side="left", padx=(0, 5))
        self.clear_btn = self._btn(left, "Temizle", self.clear_items, width=8)
        self.clear_btn.pack(side="left")

        right = tk.Frame(buttons, bg=CARD)
        right.pack(side="right")
        self.cancel_btn = self._btn(right, "İptal", self.cancel_work, width=7)
        self.cancel_btn.pack(side="right", padx=(5, 0))
        self.cancel_btn.config(state="disabled")
        self.run_btn = self._btn(right, "İşlemi Başlat", self.start, primary=True, width=13)
        self.run_btn.pack(side="right")

        tk.Label(body, textvariable=self.summary, bg=CARD, fg=MUTED, font=(self.ff(), 7, "bold")).pack(anchor="w", padx=11, pady=(0, 3))

        tf = tk.Frame(body, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        tf.pack(fill="both", expand=True, padx=11, pady=(0, 7))
        self.tree = ttk.Treeview(tf, columns=("tip", "dosya", "klasor"), show="headings", height=5)
        for key, title, width in [("tip", "Tür", 65), ("dosya", "Dosya", 220), ("klasor", "Klasör", 340)]:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=60 if key == "tip" else 110, stretch=(key != "tip"))
        sy = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sy.set)
        sy.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        bottom = tk.Frame(body, bg=CARD)
        bottom.pack(fill="x", padx=11, pady=(0, 8))
        self.prog = ttk.Progressbar(bottom, style="P.Horizontal.TProgressbar")
        self.prog.pack(fill="x", pady=(0, 3))
        tk.Label(bottom, textvariable=self.status, bg=CARD, fg=MUTED, font=(self.ff(), 7)).pack(side="left")

    def _center_dialog(self, dialog: tk.Toplevel):
        self.update_idletasks()
        dialog.update_idletasks()
        pw = max(1, self.winfo_width())
        ph = max(1, self.winfo_height())
        px = self.winfo_rootx()
        py = self.winfo_rooty()
        dw = dialog.winfo_reqwidth()
        dh = dialog.winfo_reqheight()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        x = max(8, min(x, sw - dw - 8))
        y = max(8, min(y, sh - dh - 8))
        dialog.geometry(f"+{x}+{y}")

    def _prompt_action(self):
        if self.action_dialog is not None:
            try:
                if self.action_dialog.winfo_exists():
                    self._center_dialog(self.action_dialog)
                    self.action_dialog.lift()
                    self.action_dialog.focus_force()
                    return
            except Exception:
                pass

        dialog = tk.Toplevel(self)
        self.action_dialog = dialog
        dialog.title("İşlem Seç")
        dialog.configure(bg=BG)
        dialog.transient(self)
        dialog.resizable(False, False)

        fr = tk.Frame(dialog, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        fr.pack(fill="both", expand=True, padx=8, pady=8)
        tk.Label(fr, text="Bu dosyalarla ne yapmak istiyorsun?", bg=CARD, fg=TEXT, font=(self.ff(), 10, "bold")).pack(anchor="w", padx=10, pady=(9, 3))
        tk.Label(
            fr,
            text="Word: satır düzeltme • PDF/görsel: Word'e aktarım",
            bg=CARD,
            fg=MUTED,
            font=(self.ff(), 8),
            wraplength=330,
            justify="left",
        ).pack(anchor="w", padx=10, pady=(0, 7))
        row = tk.Frame(fr, bg=CARD)
        row.pack(fill="x", padx=10, pady=(0, 9))

        def close_dialog():
            try:
                dialog.grab_release()
            except Exception:
                pass
            try:
                dialog.destroy()
            except Exception:
                pass
            self.action_dialog = None

        def choose(value: str):
            self.action.set(value)
            close_dialog()
            self._refresh()

        self._btn(row, "Satırları Düzelt", lambda: choose("repair"), width=15).pack(side="left", padx=(0, 6))
        self._btn(row, "Word'e Aktar", lambda: choose("convert"), primary=True, width=15).pack(side="left")
        dialog.protocol("WM_DELETE_WINDOW", close_dialog)
        dialog.after_idle(lambda: self._center_dialog(dialog))
        dialog.grab_set()
        dialog.focus_force()
        self.wait_window(dialog)
        self.action_dialog = None

    def _busy(self):
        return self.worker is not None and self.worker.is_alive()

    def _set_busy(self, busy: bool):
        self.run_btn.config(state="disabled" if busy else "normal")
        self.cancel_btn.config(state="normal" if busy else "disabled")
        state = "disabled" if busy else "normal"
        self.file_btn.config(state=state)
        self.folder_btn.config(state=state)
        self.clear_btn.config(state=state)

    def clear_items(self):
        if self._busy():
            return
        self.items = []
        self.action.set("")
        self._refresh()
        self.status.set("Hazır")
        self.prog["value"] = 0

    def _split_drop(self, data):
        try:
            return list(self.tk.splitlist(data))
        except Exception:
            return [p.strip("{}") for p in data.split() if p]

    def _on_drop(self, event):
        if not self._busy():
            self.add_paths(self._split_drop(event.data), ask_action=True)

    def pick_files(self):
        if self._busy():
            return
        sel = filedialog.askopenfilenames(filetypes=[
            ("Tüm desteklenenler", "*.docx *.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
            ("Word", "*.docx"),
            ("PDF ve Görseller", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
        ])
        if sel:
            self.add_paths(sel, ask_action=True)

    def pick_folder(self):
        if self._busy():
            return
        d = filedialog.askdirectory()
        if d:
            self.add_paths([d], ask_action=True)

    def add_paths(self, paths, ask_action=False):
        if self._busy():
            return
        new = []
        for raw in paths:
            p = Path(str(raw).strip().strip("{}")).expanduser()
            if not p.exists():
                continue
            if p.is_dir():
                candidates = collect_docx_from_folder(p, self.recursive.get()) + collect_media_from_folder(p, self.recursive.get())
                for x in candidates:
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
            self._prompt_action()

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for p in self.items:
            kind = "Word" if p.suffix.lower() == ".docx" else ("PDF" if p.suffix.lower() == ".pdf" else "Görsel")
            self.tree.insert("", "end", values=(kind, p.name, str(p.parent)))
        action_label = {"repair": " • Satırları Düzelt", "convert": " • Word'e Aktar"}.get(self.action.get(), "")
        self.summary.set((f"{len(self.items)} öge seçildi" + action_label) if self.items else "Henüz dosya veya klasör seçilmedi")

    def start(self):
        if self._busy():
            return
        if not self.items:
            return messagebox.showinfo(APP_NAME, "Önce dosya veya klasör seçin.")
        action = self.action.get()
        if action not in {"repair", "convert"}:
            self._prompt_action()
            action = self.action.get()
        if action not in {"repair", "convert"}:
            return

        if action == "repair":
            targets = [p for p in self.items if p.suffix.lower() == ".docx" and not p.stem.endswith("_Düzelti")]
            if not targets:
                return messagebox.showinfo(APP_NAME, "Satır düzeltme için en az bir .docx dosyası seçin.")
        else:
            targets = [p for p in self.items if p.suffix.lower() == ".pdf" or p.suffix.lower() in IMG_EXTS]
            if not targets:
                return messagebox.showinfo(APP_NAME, "Word'e aktarım için PDF veya görsel seçin.")

        self.cancel_event.clear()
        self.prog["value"] = 0
        self.prog["maximum"] = max(1, len(targets))
        self.status.set("İşlem başlatılıyor…")
        self._set_busy(True)
        self.worker = threading.Thread(target=self._work, args=(action, targets, self.include_images.get()), daemon=True)
        self.worker.start()

    def cancel_work(self):
        if self._busy():
            self.cancel_event.set()
            self.cancel_btn.config(state="disabled")
            self.status.set("İptal ediliyor…")

    def _work(self, action, targets, include_images):
        done = repairs = pages = images = ocr_pages = 0
        errors = []
        cancelled = False

        for i, p in enumerate(targets, 1):
            if self.cancel_event.is_set():
                cancelled = True
                break
            self.q.put(("status", f"İşleniyor: {p.name}"))
            try:
                if action == "repair":
                    r = repair_file(p)
                    if r.skipped:
                        errors.append(f"{p.name}: {r.message}")
                    else:
                        done += 1
                        repairs += r.repaired_breaks
                else:
                    r = convert_media_to_word(p, include_images=include_images, cancel_event=self.cancel_event)
                    if self.cancel_event.is_set() or getattr(r, "cancelled", False):
                        cancelled = True
                        break
                    if r.skipped:
                        errors.append(f"{p.name}: {r.message}")
                    else:
                        done += 1
                        pages += r.pages
                        images += r.images
                        ocr_pages += getattr(r, "ocr_pages", 0)
            except Exception as exc:
                if self.cancel_event.is_set():
                    cancelled = True
                    break
                errors.append(f"{p.name}: {exc}")
            self.q.put(("prog", i))

        self.q.put(("done", action, done, repairs, pages, images, ocr_pages, errors, cancelled))

    def _check_update_worker(self, user_requested):
        try:
            self.q.put(("update_check", check_for_update(APP_VERSION), user_requested, None))
        except Exception as exc:
            self.q.put(("update_check", None, user_requested, str(exc)))

    def update_clicked(self):
        if self._busy():
            return messagebox.showinfo(APP_NAME, "Önce devam eden işlemi tamamlayın veya iptal edin.")
        if self.update_info:
            if not messagebox.askyesno(
                APP_NAME,
                f"v{self.update_info.version} sürümü indirilsin ve kurulsun mu?\n\nProgram güncelleme sırasında kapanıp yeniden açılacak.",
            ):
                return
            self.update_btn.config(state="disabled", text="İndiriliyor…")
            threading.Thread(target=self._download_update_worker, args=(self.update_info,), daemon=True).start()
        else:
            self.update_btn.config(state="disabled", text="Kontrol…")
            threading.Thread(target=self._check_update_worker, args=(True,), daemon=True).start()

    def _download_update_worker(self, info):
        try:
            self.q.put(("update_ready", download_and_stage(info), None))
        except Exception as exc:
            self.q.put(("update_ready", None, str(exc)))

    def on_close(self):
        if self.closing:
            return
        self.closing = True
        if self._busy():
            self.cancel_event.set()
            self.status.set("Kapatılıyor…")
            self.after(450, self.destroy)
        else:
            self.destroy()

    def _poll(self):
        if self.closing:
            try:
                self.update_idletasks()
            except Exception:
                return
        try:
            while True:
                m = self.q.get_nowait()
                k = m[0]
                if k == "status":
                    self.status.set(m[1])
                elif k == "prog":
                    self.prog["value"] = m[1]
                elif k == "done":
                    self._set_busy(False)
                    _, action, done, repairs, pages, images, ocr_pages, errors, cancelled = m
                    if cancelled:
                        msg = f"İşlem iptal edildi. • {done} dosya tamamlandı"
                    elif action == "repair":
                        msg = f"Tamamlandı • {done} dosya • {repairs} onarım"
                    else:
                        msg = f"Tamamlandı • {done} Word dosyası • {pages} sayfa • {ocr_pages} OCR"
                        if self.include_images.get():
                            msg += f" • {images} görsel"
                    if errors:
                        msg += f"\n\nAtlanan dosya: {len(errors)}"
                    self.status.set(msg.split("\n")[0])
                    if not self.closing:
                        messagebox.showinfo(APP_NAME, msg)
                elif k == "update_check":
                    _, info, user_requested, error = m
                    self.update_btn.config(state="normal")
                    if error:
                        self.update_btn.config(text="Güncelle")
                        if user_requested:
                            messagebox.showwarning(APP_NAME, "Güncelleme kontrolü yapılamadı.\n\n" + error)
                    elif info:
                        self.update_info = info
                        self.update_btn.config(text=f"Güncelle v{info.version}")
                        if user_requested:
                            self.update_clicked()
                    else:
                        self.update_info = None
                        self.update_btn.config(text="Güncel")
                        if user_requested:
                            messagebox.showinfo(APP_NAME, "Kullanılan sürüm güncel.")
                elif k == "update_ready":
                    _, asset, error = m
                    if error:
                        self.update_btn.config(state="normal", text="Güncelle")
                        messagebox.showerror(APP_NAME, "Güncelleme indirilemedi.\n\n" + error)
                    else:
                        messagebox.showinfo(APP_NAME, "Güncelleme indirildi. Program kapanıp yeni sürümle yeniden açılacak.")
                        try:
                            install_staged(Path(asset))
                            self.after(400, lambda: os._exit(0))
                        except Exception as exc:
                            self.update_btn.config(state="normal", text="Güncelle")
                            messagebox.showerror(APP_NAME, "Güncelleme kurulamadı.\n\n" + str(exc))
        except queue.Empty:
            pass
        except tk.TclError:
            return
        if not self.closing:
            self.after(120, self._poll)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()

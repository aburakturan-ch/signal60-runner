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
APP_VERSION = "0.2.6"
BG="#F8F4F7"; CARD="#FFFDFE"; PINK="#B65C82"; PINK_D="#8E3F64"; PINK_S="#F2DCE6"; TEXT="#332B31"; MUTED="#756A72"; BORDER="#E8DDE3"; GREEN="#4F7D66"
IMG_EXTS=tuple(sorted(SUPPORTED_IMAGE_EXTS))
ROOT = TkinterDnD.Tk if TkinterDnD else tk.Tk

class App(ROOT):
    def __init__(self):
        super().__init__(); self.title(APP_NAME); self.configure(bg=BG)
        sw,sh=self.winfo_screenwidth(),self.winfo_screenheight(); self.geometry(f"{min(840,max(640,int(sw*.74)))}x{min(540,max(390,int(sh*.72)))}"); self.minsize(620,380)
        self.items=[]; self.action=tk.StringVar(value=""); self.include_images=tk.BooleanVar(value=False); self.recursive=tk.BooleanVar(value=False); self.status=tk.StringVar(value="Hazır"); self.summary=tk.StringVar(value="Henüz dosya veya klasör seçilmedi"); self.q=queue.Queue(); self.update_info=None
        self._style(); self._ui(); self.after(120,self._poll); self.after(1200,lambda:threading.Thread(target=self._check_update_worker,args=(False,),daemon=True).start())

    def ff(self): return "Segoe UI" if platform.system()=="Windows" else "Helvetica Neue"

    def _style(self):
        s=ttk.Style(self)
        try: s.theme_use("aqua" if platform.system()=="Darwin" else "clam")
        except: pass
        s.configure("Treeview",rowheight=24,background=CARD,fieldbackground=CARD,foreground=TEXT); s.configure("Treeview.Heading",font=(self.ff(),9,"bold")); s.configure("P.Horizontal.TProgressbar",troughcolor=PINK_S,background=PINK)

    def _btn(self,p,t,c,primary=False,green=False):
        mac=platform.system()=="Darwin"; accent=GREEN if green else PINK; dark="#365B49" if green else PINK_D; soft="#E4F0E9" if green else PINK_S
        bg,fg=((soft,dark) if primary and mac else ((accent,"white") if primary else ("white",TEXT)))
        return tk.Button(p,text=t,command=c,bg=bg,fg=fg,activebackground=dark,activeforeground=(fg if mac else "white"),relief="flat",bd=0,padx=10,pady=6,font=(self.ff(),9,"bold"),highlightbackground=(accent if primary else BORDER),highlightthickness=1,cursor="hand2")

    def _ui(self):
        h=tk.Frame(self,bg=BG); h.pack(fill="x",padx=16,pady=(10,6)); tk.Label(h,text="✿",font=("Georgia",22),fg=PINK,bg=BG).pack(side="left",padx=(0,8))
        hb=tk.Frame(h,bg=BG); hb.pack(side="left",fill="x",expand=True); tk.Label(hb,text=APP_NAME,font=(self.ff(),17,"bold"),fg=TEXT,bg=BG).pack(anchor="w"); tk.Label(hb,text="Dosyayı bırak • işlemi seç • yeni kopyanı oluştur",font=(self.ff(),9),fg=MUTED,bg=BG).pack(anchor="w")
        right=tk.Frame(h,bg=BG); right.pack(side="right"); tk.Label(right,text="v"+APP_VERSION,font=(self.ff(),8),fg=MUTED,bg=BG).pack(anchor="e",pady=(0,3)); self.update_btn=self._btn(right,"Güncelle",self.update_clicked,green=True); self.update_btn.pack(anchor="e")
        body=tk.Frame(self,bg=CARD,highlightbackground=BORDER,highlightthickness=1); body.pack(fill="both",expand=True,padx=14,pady=(0,12))
        top=tk.Frame(body,bg=CARD); top.pack(fill="x",padx=14,pady=(12,9)); self.drop=tk.Label(top,text="Dosya veya klasörü buraya sürükleyip bırak\nDOCX • PDF • JPG • PNG • TIFF • BMP • WEBP",justify="center",bg="#FBF8FA",fg=MUTED,font=(self.ff(),11,"bold"),highlightbackground=BORDER,highlightthickness=2,pady=16); self.drop.pack(fill="x")
        if DND_FILES:
            self.drop.drop_target_register(DND_FILES); self.drop.dnd_bind('<<Drop>>',self._on_drop)
        opts=tk.Frame(body,bg=CARD); opts.pack(fill="x",padx=14,pady=(0,6)); tk.Checkbutton(opts,text="Alt klasörleri tara",variable=self.recursive,bg=CARD,fg=MUTED,selectcolor=CARD,font=(self.ff(),9)).pack(side="left"); tk.Checkbutton(opts,text="Word'e aktarırken görselleri de ekle",variable=self.include_images,bg=CARD,fg=MUTED,selectcolor=CARD,font=(self.ff(),9)).pack(side="left",padx=(16,0)); tk.Label(opts,text="OCR: basılı metin + el yazısı denemesi • yamuk taramada otomatik düzeltme",bg=CARD,fg=MUTED,font=(self.ff(),8)).pack(side="right")
        note=tk.Frame(body,bg="#FBF9FD",highlightbackground=BORDER,highlightthickness=1); note.pack(fill="x",padx=14,pady=(0,8)); tk.Label(note,text="Kaynak dosya değişmez.  Satır düzeltme: _Düzelti.docx  •  Aktarım: _Aktarma.docx",bg="#FBF9FD",fg=MUTED,font=(self.ff(),9),padx=10,pady=6).pack(anchor="w")
        buttons=tk.Frame(body,bg=CARD); buttons.pack(fill="x",padx=14,pady=(0,6)); self._btn(buttons,"Dosya Seç",self.pick_files).pack(side="left",padx=(0,6)); self._btn(buttons,"Klasör Seç",self.pick_folder).pack(side="left",padx=(0,6)); self._btn(buttons,"Temizle",self.clear_items).pack(side="left"); self.run_btn=self._btn(buttons,"İşlemi Başlat",self.start,primary=True); self.run_btn.pack(side="right")
        tk.Label(body,textvariable=self.summary,bg=CARD,fg=MUTED,font=(self.ff(),8,"bold")).pack(anchor="w",padx=14,pady=(0,3))
        tf=tk.Frame(body,bg=CARD,highlightbackground=BORDER,highlightthickness=1); tf.pack(fill="both",expand=True,padx=14,pady=(0,9)); self.tree=ttk.Treeview(tf,columns=("tip","dosya","klasor"),show="headings",height=6)
        for key,title,width in [("tip","Tür",80),("dosya","Dosya",250),("klasor","Klasör",430)]: self.tree.heading(key,text=title); self.tree.column(key,width=width,minwidth=70 if key=='tip' else 140,stretch=(key!='tip'))
        sy=ttk.Scrollbar(tf,orient="vertical",command=self.tree.yview); self.tree.configure(yscrollcommand=sy.set); sy.pack(side="right",fill="y"); self.tree.pack(fill="both",expand=True)
        bottom=tk.Frame(body,bg=CARD); bottom.pack(fill="x",padx=14,pady=(0,11)); self.prog=ttk.Progressbar(bottom,style="P.Horizontal.TProgressbar"); self.prog.pack(fill="x",pady=(0,4)); tk.Label(bottom,textvariable=self.status,bg=CARD,fg=MUTED,font=(self.ff(),8)).pack(side="left")

    def _prompt_action(self):
        dialog=tk.Toplevel(self); dialog.title("İşlem Seç"); dialog.configure(bg=BG); dialog.transient(self); dialog.grab_set(); dialog.resizable(False,False)
        fr=tk.Frame(dialog,bg=CARD,highlightbackground=BORDER,highlightthickness=1); fr.pack(fill='both',expand=True,padx=12,pady=12); tk.Label(fr,text="Bu dosyalarla ne yapmak istiyorsun?",bg=CARD,fg=TEXT,font=(self.ff(),12,'bold')).pack(anchor='w',padx=14,pady=(14,4)); tk.Label(fr,text="Word dosyalarında satır düzeltme; PDF ve görsellerde Word aktarımı kullanılabilir.",bg=CARD,fg=MUTED,font=(self.ff(),9),wraplength=430,justify='left').pack(anchor='w',padx=14,pady=(0,10))
        btns=tk.Frame(fr,bg=CARD); btns.pack(fill='x',padx=14,pady=(0,14))
        def choose(val): self.action.set(val); dialog.destroy()
        self._btn(btns,"Satırları Düzelt",lambda:choose('repair')).pack(side='left',padx=(0,8)); self._btn(btns,"Word'e Aktar",lambda:choose('convert'),primary=True).pack(side='left'); dialog.protocol("WM_DELETE_WINDOW",dialog.destroy); dialog.update_idletasks(); x=self.winfo_rootx()+max(20,(self.winfo_width()-dialog.winfo_width())//2); y=self.winfo_rooty()+max(20,(self.winfo_height()-dialog.winfo_height())//2); dialog.geometry(f"+{x}+{y}"); self.wait_window(dialog)

    def clear_items(self): self.items=[]; self.action.set(""); self._refresh(); self.status.set("Hazır"); self.prog['value']=0
    def _split_drop(self,data):
        try: return list(self.tk.splitlist(data))
        except: return [p.strip('{}') for p in data.split() if p]
    def _on_drop(self,event): self.add_paths(self._split_drop(event.data),ask_action=True)
    def pick_files(self):
        sel=filedialog.askopenfilenames(filetypes=[("Tüm desteklenenler","*.docx *.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),("Word","*.docx"),("PDF ve Görseller","*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp")]);
        if sel: self.add_paths(sel,ask_action=True)
    def pick_folder(self):
        d=filedialog.askdirectory();
        if d: self.add_paths([d],ask_action=True)

    def add_paths(self,paths,ask_action=False):
        new=[]
        for raw in paths:
            p=Path(str(raw).strip().strip('{}')).expanduser()
            if not p.exists(): continue
            if p.is_dir():
                for x in collect_docx_from_folder(p,self.recursive.get())+collect_media_from_folder(p,self.recursive.get()):
                    if x not in new and x not in self.items: new.append(x)
            else:
                suf=p.suffix.lower(); allowed=(suf=='.docx' and not p.stem.endswith('_Düzelti')) or suf=='.pdf' or suf in IMG_EXTS
                if allowed and p not in new and p not in self.items: new.append(p)
        if not new: messagebox.showinfo(APP_NAME,"Desteklenen bir dosya bulunamadı."); return
        self.items.extend(new); self._refresh();
        if ask_action: self._prompt_action()

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for p in self.items:
            kind="Word" if p.suffix.lower()=='.docx' else ("PDF" if p.suffix.lower()=='.pdf' else "Görsel"); self.tree.insert("","end",values=(kind,p.name,str(p.parent)))
        action_label={"repair":" • Satırları Düzelt","convert":" • Word'e Aktar"}.get(self.action.get(),""); self.summary.set((f"{len(self.items)} öge seçildi"+action_label) if self.items else "Henüz dosya veya klasör seçilmedi")

    def start(self):
        if not self.items: return messagebox.showinfo(APP_NAME,"Önce dosya veya klasör seçin.")
        action=self.action.get()
        if action not in {'repair','convert'}: self._prompt_action(); action=self.action.get()
        if action not in {'repair','convert'}: return
        if action=='repair':
            targets=[p for p in self.items if p.suffix.lower()=='.docx' and not p.stem.endswith('_Düzelti')]
            if not targets: return messagebox.showinfo(APP_NAME,"Satır düzeltme için en az bir .docx dosyası seçin.")
        else:
            targets=[p for p in self.items if p.suffix.lower()=='.pdf' or p.suffix.lower() in IMG_EXTS]
            if not targets: return messagebox.showinfo(APP_NAME,"Word'e aktarım için PDF veya görsel seçin.")
        self.run_btn.config(state='disabled'); self.prog['value']=0; self.prog['maximum']=len(targets); threading.Thread(target=self._work,args=(action,targets,self.include_images.get()),daemon=True).start()

    def _work(self,action,targets,include_images):
        done=repairs=pages=images=ocr_pages=0; errors=[]
        for i,p in enumerate(targets,1):
            self.q.put(('status',f"İşleniyor: {p.name}"))
            if action=='repair':
                r=repair_file(p); errors.append(f"{p.name}: {r.message}") if r.skipped else None; done+=0 if r.skipped else 1; repairs+=0 if r.skipped else r.repaired_breaks
            else:
                r=convert_media_to_word(p,include_images=include_images); errors.append(f"{p.name}: {r.message}") if r.skipped else None; done+=0 if r.skipped else 1; pages+=0 if r.skipped else r.pages; images+=0 if r.skipped else r.images; ocr_pages+=0 if r.skipped else getattr(r,'ocr_pages',0)
            self.q.put(('prog',i))
        self.q.put(('done',action,done,repairs,pages,images,ocr_pages,errors))

    def _check_update_worker(self,user_requested):
        try: self.q.put(('update_check',check_for_update(APP_VERSION),user_requested,None))
        except Exception as exc: self.q.put(('update_check',None,user_requested,str(exc)))

    def update_clicked(self):
        if self.update_info:
            if not messagebox.askyesno(APP_NAME,f"v{self.update_info.version} sürümü indirilsin ve kurulsun mu?\n\nProgram güncelleme sırasında kapanıp yeniden açılacak."): return
            self.update_btn.config(state='disabled',text='İndiriliyor…'); threading.Thread(target=self._download_update_worker,args=(self.update_info,),daemon=True).start()
        else:
            self.update_btn.config(state='disabled',text='Kontrol…'); threading.Thread(target=self._check_update_worker,args=(True,),daemon=True).start()

    def _download_update_worker(self,info):
        try: self.q.put(('update_ready',download_and_stage(info),None))
        except Exception as exc: self.q.put(('update_ready',None,str(exc)))

    def _poll(self):
        try:
            while True:
                m=self.q.get_nowait(); k=m[0]
                if k=='status': self.status.set(m[1])
                elif k=='prog': self.prog['value']=m[1]
                elif k=='done':
                    self.run_btn.config(state='normal'); _,action,done,repairs,pages,images,ocr_pages,errors=m; msg=(f"Tamamlandı • {done} dosya • {repairs} onarım" if action=='repair' else f"Tamamlandı • {done} Word dosyası • {pages} sayfa • {ocr_pages} OCR"); msg+=(f" • {images} görsel" if action=='convert' and self.include_images.get() else ""); msg+=(f"\n\nAtlanan dosya: {len(errors)}" if errors else ""); self.status.set(msg.split('\n')[0]); messagebox.showinfo(APP_NAME,msg)
                elif k=='update_check':
                    _,info,user_requested,error=m; self.update_btn.config(state='normal')
                    if error:
                        self.update_btn.config(text='Güncelle');
                        if user_requested: messagebox.showwarning(APP_NAME,"Güncelleme kontrolü yapılamadı.\n\n"+error)
                    elif info:
                        self.update_info=info; self.update_btn.config(text=f"Güncelle v{info.version}");
                        if user_requested: self.update_clicked()
                    else:
                        self.update_info=None; self.update_btn.config(text='Güncel');
                        if user_requested: messagebox.showinfo(APP_NAME,"Kullanılan sürüm güncel.")
                elif k=='update_ready':
                    _,asset,error=m
                    if error: self.update_btn.config(state='normal',text='Güncelle'); messagebox.showerror(APP_NAME,"Güncelleme indirilemedi.\n\n"+error)
                    else:
                        messagebox.showinfo(APP_NAME,"Güncelleme indirildi. Program şimdi kapanıp yeni sürümle yeniden açılacak.")
                        try: install_staged(Path(asset)); self.after(400,lambda:os._exit(0))
                        except Exception as exc: self.update_btn.config(state='normal',text='Güncelle'); messagebox.showerror(APP_NAME,"Güncelleme kurulamadı.\n\n"+str(exc))
        except queue.Empty: pass
        self.after(120,self._poll)

def main(): App().mainloop()
if __name__=='__main__': main()

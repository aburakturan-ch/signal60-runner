from __future__ import annotations
import platform, queue, threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from line_repair import collect_docx_from_folder, repair_file
from pdf_to_word import collect_pdf_from_folder, convert_pdf_to_word

APP_NAME="Gülşen'in Word Araçları"; APP_VERSION="0.2.3"
BG="#F8F4F7"; CARD="#FFFDFE"; PINK="#B65C82"; PINK_D="#8E3F64"; PINK_S="#F2DCE6"; LILAC="#8C6FB2"; LILAC_D="#674A8B"; LILAC_S="#EEE7F6"; TEXT="#332B31"; MUTED="#756A72"; BORDER="#E8DDE3"

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP_NAME); self.configure(bg=BG)
        sw,sh=self.winfo_screenwidth(),self.winfo_screenheight(); self.geometry(f"{min(860,max(650,int(sw*.76)))}x{min(500,max(370,int(sh*.68)))}"); self.minsize(620,350)
        self.rf=[]; self.pf=[]; self.rr=tk.BooleanVar(False); self.pr=tk.BooleanVar(False); self.img=tk.BooleanVar(False); self.rs=tk.StringVar(value="Hazır"); self.ps=tk.StringVar(value="Hazır"); self.q=queue.Queue()
        self._style(); self._ui(); self.after(100,self._poll)
    def ff(self): return "Segoe UI" if platform.system()=="Windows" else "Helvetica Neue"
    def _style(self):
        s=ttk.Style(self)
        try: s.theme_use("aqua" if platform.system()=="Darwin" else "clam")
        except: pass
        s.configure("TNotebook",background=BG,borderwidth=0); s.configure("TNotebook.Tab",padding=(12,6),font=(self.ff(),9,"bold")); s.configure("Treeview",rowheight=24,background=CARD,fieldbackground=CARD,foreground=TEXT); s.configure("Treeview.Heading",font=(self.ff(),9,"bold")); s.configure("P.Horizontal.TProgressbar",troughcolor=PINK_S,background=PINK); s.configure("L.Horizontal.TProgressbar",troughcolor=LILAC_S,background=LILAC)
    def _btn(self,p,t,c,primary=False,lilac=False):
        a,d,soft=(LILAC,LILAC_D,LILAC_S) if lilac else (PINK,PINK_D,PINK_S); mac=platform.system()=="Darwin"
        bg,fg=(soft,d) if primary and mac else ((a,"white") if primary else ("white",TEXT))
        return tk.Button(p,text=t,command=c,bg=bg,fg=fg,activebackground=d,activeforeground=(fg if mac else "white"),relief="flat",bd=0,padx=10,pady=5,font=(self.ff(),9,"bold"),highlightbackground=a if primary else BORDER,highlightthickness=1,cursor="hand2")
    def _ui(self):
        h=tk.Frame(self,bg=BG); h.pack(fill="x",padx=16,pady=(9,5)); tk.Label(h,text="✿",font=("Georgia",22),fg=PINK,bg=BG).pack(side="left",padx=(0,7)); b=tk.Frame(h,bg=BG); b.pack(side="left",fill="x",expand=True); tk.Label(b,text=APP_NAME,font=(self.ff(),17,"bold"),fg=TEXT,bg=BG).pack(anchor="w"); tk.Label(b,text="Word belgeleri için sade, güvenli ve kolay araçlar",font=(self.ff(),9),fg=MUTED,bg=BG).pack(anchor="w"); tk.Label(h,text="v"+APP_VERSION,font=(self.ff(),8),fg=MUTED,bg=BG).pack(side="right",anchor="n")
        self.nb=ttk.Notebook(self); self.nb.pack(fill="both",expand=True,padx=14,pady=(0,10)); a=tk.Frame(self.nb,bg=CARD); p=tk.Frame(self.nb,bg=CARD); x=tk.Frame(self.nb,bg=CARD); self.nb.add(a,text="Satır Onarıcı"); self.nb.add(p,text="PDF'den Word'e Aktar"); self.nb.add(x,text="Hakkında"); self._repair(a); self._pdf(p); self._about(x)
    def _head(self,w,title,desc):
        tk.Label(w,text=title,font=(self.ff(),14,"bold"),fg=TEXT,bg=CARD).pack(anchor="w"); tk.Label(w,text=desc,font=(self.ff(),9),fg=MUTED,bg=CARD,justify="left",wraplength=560).pack(anchor="w",pady=(2,7))
    def _tree(self,w):
        f=tk.Frame(w,bg=CARD,highlightbackground=BORDER,highlightthickness=1); f.pack(fill="both",expand=True); t=ttk.Treeview(f,columns=("f","d"),show="headings",height=4); t.heading("f",text="Dosya"); t.heading("d",text="Klasör"); t.column("f",width=300,minwidth=140); t.column("d",width=450,minwidth=180); y=ttk.Scrollbar(f,orient="vertical",command=t.yview); t.configure(yscrollcommand=y.set); y.pack(side="right",fill="y"); t.pack(fill="both",expand=True); return t
    def _repair(self,p):
        w=tk.Frame(p,bg=CARD); w.pack(fill="both",expand=True,padx=16,pady=11); self._head(w,"Satır Onarıcı","Cümlenin ortasında oluşan yanlış paragraf sonlarını bulur ve birleştirir."); n=tk.Frame(w,bg=PINK_S,padx=9,pady=5); n.pack(fill="x",pady=(0,7)); tk.Label(n,text="Orijinal korunur.  Sonuç: DosyaAdı_Düzelti.docx",bg=PINK_S,fg=PINK_D,font=(self.ff(),9,"bold")).pack(anchor="w"); c=tk.Frame(w,bg=CARD); c.pack(fill="x",pady=(0,5)); self._btn(c,"Dosya Seç",self.rfiles).pack(side="left",padx=(0,6)); self._btn(c,"Klasör Seç",self.rfolder).pack(side="left",padx=(0,6)); self._btn(c,"Temizle",self.rclear).pack(side="left"); tk.Checkbutton(c,text="Alt klasörleri tara",variable=self.rr,bg=CARD,fg=MUTED,selectcolor=CARD,font=(self.ff(),8)).pack(side="right"); self.rc=tk.Label(w,text="Henüz dosya seçilmedi",bg=CARD,fg=MUTED,font=(self.ff(),8,"bold")); self.rc.pack(anchor="w",pady=(0,3)); bottom=tk.Frame(w,bg=CARD); bottom.pack(side="bottom",fill="x",pady=(6,0)); self.rt=self._tree(w); self.rprog=ttk.Progressbar(bottom,style="P.Horizontal.TProgressbar"); self.rprog.pack(fill="x",pady=(0,4)); r=tk.Frame(bottom,bg=CARD); r.pack(fill="x"); tk.Label(r,textvariable=self.rs,bg=CARD,fg=MUTED,font=(self.ff(),8)).pack(side="left"); self.rb=self._btn(r,"Onar ve Kopyala",self.rstart,True); self.rb.pack(side="right")
    def _pdf(self,p):
        w=tk.Frame(p,bg=CARD); w.pack(fill="both",expand=True,padx=16,pady=11); self._head(w,"PDF'den Word'e Aktar","PDF metnini Word'e aktarır; kalın, italik, boyut, renk ve hizalamayı mümkün olduğunca korur."); n=tk.Frame(w,bg=LILAC_S,padx=9,pady=5); n.pack(fill="x",pady=(0,6)); tk.Label(n,text="PDF korunur.  Sonuç: DosyaAdı_Aktarma.docx",bg=LILAC_S,fg=LILAC_D,font=(self.ff(),9,"bold")).pack(anchor="w"); o=tk.Frame(w,bg="#FBF9FD",padx=8,pady=4,highlightbackground=BORDER,highlightthickness=1); o.pack(fill="x",pady=(0,6)); tk.Checkbutton(o,text="Görselleri de aktar",variable=self.img,bg="#FBF9FD",fg=TEXT,selectcolor="#FBF9FD",font=(self.ff(),9,"bold")).pack(side="left"); tk.Label(o,text="(varsayılan: kapalı)",bg="#FBF9FD",fg=MUTED,font=(self.ff(),8)).pack(side="left",padx=(6,0)); c=tk.Frame(w,bg=CARD); c.pack(fill="x",pady=(0,5)); self._btn(c,"PDF Seç",self.pfiles,lilac=True).pack(side="left",padx=(0,6)); self._btn(c,"Klasör Seç",self.pfolder,lilac=True).pack(side="left",padx=(0,6)); self._btn(c,"Temizle",self.pclear,lilac=True).pack(side="left"); tk.Checkbutton(c,text="Alt klasörleri tara",variable=self.pr,bg=CARD,fg=MUTED,selectcolor=CARD,font=(self.ff(),8)).pack(side="right"); self.pc=tk.Label(w,text="Henüz PDF seçilmedi",bg=CARD,fg=MUTED,font=(self.ff(),8,"bold")); self.pc.pack(anchor="w",pady=(0,3)); bottom=tk.Frame(w,bg=CARD); bottom.pack(side="bottom",fill="x",pady=(6,0)); self.pt=self._tree(w); self.pprog=ttk.Progressbar(bottom,style="L.Horizontal.TProgressbar"); self.pprog.pack(fill="x",pady=(0,4)); r=tk.Frame(bottom,bg=CARD); r.pack(fill="x"); tk.Label(r,textvariable=self.ps,bg=CARD,fg=MUTED,font=(self.ff(),8)).pack(side="left"); self.pb=self._btn(r,"Word'e Aktar",self.pstart,True,True); self.pb.pack(side="right")
    def _about(self,p):
        w=tk.Frame(p,bg=CARD); w.pack(fill="both",expand=True,padx=20,pady=18); tk.Label(w,text=APP_NAME,font=(self.ff(),16,"bold"),bg=CARD,fg=TEXT).pack(anchor="w"); tk.Label(w,text="Satır Onarıcı ve PDF'den Word'e Aktar. Kaynak dosyalar değiştirilmez; sonuçlar yeni dosya olarak oluşturulur.",font=(self.ff(),9),bg=CARD,fg=MUTED,wraplength=620,justify="left").pack(anchor="w",pady=(8,0))
    def _fill(self,t,fs): t.delete(*t.get_children()); [t.insert("","end",values=(p.name,str(p.parent))) for p in fs]
    def rfiles(self):
        x=filedialog.askopenfilenames(filetypes=[("Word","*.docx")]); self.rf=list(map(Path,x)) if x else self.rf; self._fill(self.rt,self.rf); self.rc.config(text=f"{len(self.rf)} dosya seçildi" if self.rf else "Henüz dosya seçilmedi")
    def rfolder(self):
        x=filedialog.askdirectory(); self.rf=collect_docx_from_folder(x,self.rr.get()) if x else self.rf; self._fill(self.rt,self.rf); self.rc.config(text=f"{len(self.rf)} dosya seçildi" if self.rf else "Henüz dosya seçilmedi")
    def rclear(self): self.rf=[]; self._fill(self.rt,self.rf); self.rc.config(text="Henüz dosya seçilmedi")
    def pfiles(self):
        x=filedialog.askopenfilenames(filetypes=[("PDF","*.pdf")]); self.pf=list(map(Path,x)) if x else self.pf; self._fill(self.pt,self.pf); self.pc.config(text=f"{len(self.pf)} PDF seçildi" if self.pf else "Henüz PDF seçilmedi")
    def pfolder(self):
        x=filedialog.askdirectory(); self.pf=collect_pdf_from_folder(x,self.pr.get()) if x else self.pf; self._fill(self.pt,self.pf); self.pc.config(text=f"{len(self.pf)} PDF seçildi" if self.pf else "Henüz PDF seçilmedi")
    def pclear(self): self.pf=[]; self._fill(self.pt,self.pf); self.pc.config(text="Henüz PDF seçilmedi")
    def rstart(self):
        if not self.rf: return messagebox.showinfo(APP_NAME,"Önce dosya veya klasör seçin.")
        self.rb.config(state="disabled"); self.rprog["maximum"]=len(self.rf); threading.Thread(target=self._rw,daemon=True).start()
    def _rw(self):
        done=rep=0; errs=[]
        for i,p in enumerate(self.rf,1):
            self.q.put(("rs",p.name)); r=repair_file(p); errs.append(r.message) if r.skipped else None; done+=0 if r.skipped else 1; rep+=0 if r.skipped else r.repaired_breaks; self.q.put(("rp",i))
        self.q.put(("rd",done,rep,errs))
    def pstart(self):
        if not self.pf: return messagebox.showinfo(APP_NAME,"Önce PDF veya klasör seçin.")
        self.pb.config(state="disabled"); self.pprog["maximum"]=len(self.pf); threading.Thread(target=self._pw,args=(self.img.get(),),daemon=True).start()
    def _pw(self,img):
        done=pages=ims=0; errs=[]
        for i,p in enumerate(self.pf,1):
            self.q.put(("ps",p.name)); r=convert_pdf_to_word(p,img); errs.append(r.message) if r.skipped else None; done+=0 if r.skipped else 1; pages+=0 if r.skipped else r.pages; ims+=0 if r.skipped else r.images; self.q.put(("pp",i))
        self.q.put(("pd",done,pages,ims,errs))
    def _poll(self):
        try:
            while 1:
                m=self.q.get_nowait(); k=m[0]
                if k=="rs": self.rs.set("İşleniyor: "+m[1])
                elif k=="rp": self.rprog["value"]=m[1]
                elif k=="rd": self.rb.config(state="normal"); self.rs.set(f"Tamamlandı • {m[1]} dosya • {m[2]} onarım"); messagebox.showinfo(APP_NAME,self.rs.get())
                elif k=="ps": self.ps.set("Aktarılıyor: "+m[1])
                elif k=="pp": self.pprog["value"]=m[1]
                elif k=="pd": self.pb.config(state="normal"); self.ps.set(f"Tamamlandı • {m[1]} Word • {m[2]} sayfa"+(f" • {m[3]} görsel" if self.img.get() else "")); messagebox.showinfo(APP_NAME,self.ps.get())
        except queue.Empty: pass
        self.after(100,self._poll)

def main(): App().mainloop()
if __name__=="__main__": main()

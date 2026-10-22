from pypdf import PdfReader
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
import threading
from ttkthemes import ThemedStyle


def rtf_olustur(metin):
    """Düz metni Türkçe karakterleri koruyan geçerli bir RTF belgesine çevirir."""
    parcalar = []
    for karakter in metin:
        if karakter in '\\{}':
            parcalar.append('\\' + karakter)
        elif karakter == '\n':
            parcalar.append('\\par\n')
        elif ord(karakter) > 127:
            kod = ord(karakter)
            # RTF \u parametresi 16 bit işaretli tamsayıdır
            parcalar.append(f'\\u{kod - 65536 if kod > 32767 else kod}?')
        else:
            parcalar.append(karakter)
    return '{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Calibri;}}\\f0\\fs22\n' + ''.join(parcalar) + '}'

class PDFMetinCekici:
    def __init__(self):
        self.pencere = TkinterDnD.Tk()
        self.pencere.title("PDF Metin Çekici")
        self.pencere.geometry("800x600")
        
        # Modern tema (isteğe bağlı, hata verirse kaldırılabilir)
        try:
            style = ThemedStyle(self.pencere)
            style.set_theme("arc")
        except Exception:
            pass
        
        # Ana frame
        self.ana_frame = ttk.Frame(self.pencere, padding="20")
        self.ana_frame.pack(expand=True, fill='both')
        
        # Başlık
        self.baslik = ttk.Label(self.ana_frame, text="PDF Metin Çekici", font=("Arial", 16, "bold"))
        self.baslik.pack(pady=10)
        
        # Sürükle-bırak alanı
        self.drop_frame = ttk.LabelFrame(self.ana_frame, text="PDF Dosyasını Buraya Sürükleyin", padding="20")
        self.drop_frame.pack(fill='x', pady=10)
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.dosya_surukle_birak)
        # Belirgin sürükle-bırak label'ı
        self.dnd_label = tk.Label(self.drop_frame, text="Buraya PDF dosyasını sürükleyin", bg="#e0e0e0", fg="#333", font=("Arial", 12), height=3)
        self.dnd_label.pack(expand=True, fill='both')
        self.dnd_label.drop_target_register(DND_FILES)
        self.dnd_label.dnd_bind('<<Drop>>', self.dosya_surukle_birak)
        
        # Dosya seçme butonu
        self.pdf_sec_buton = ttk.Button(self.ana_frame, text="PDF Dosyası Seç", command=self.pdf_sec)
        self.pdf_sec_buton.pack(pady=5)
        
        # Seçilen dosya yolu
        self.dosya_yolu_label = ttk.Label(self.ana_frame, text="Henüz dosya seçilmedi", wraplength=700)
        self.dosya_yolu_label.pack(pady=5)
        
        # İlerleme çubuğu
        self.ilerleme = ttk.Progressbar(self.ana_frame, mode='determinate')
        self.ilerleme.pack(fill='x', pady=5)
        
        # Metin çıkarma butonu
        self.metin_cek_buton = ttk.Button(self.ana_frame, text="Metni Çıkar", command=self.metin_cek, state='disabled')
        self.metin_cek_buton.pack(pady=5)
        
        # Kaydetme formatı seçimi
        self.format_frame = ttk.LabelFrame(self.ana_frame, text="Kaydetme Formatı", padding="10")
        self.format_frame.pack(fill='x', pady=5)
        
        self.format_var = tk.StringVar(value="txt")
        ttk.Radiobutton(self.format_frame, text="TXT", variable=self.format_var, value="txt").pack(side='left', padx=5)
        ttk.Radiobutton(self.format_frame, text="DOCX", variable=self.format_var, value="docx").pack(side='left', padx=5)
        ttk.Radiobutton(self.format_frame, text="RTF", variable=self.format_var, value="rtf").pack(side='left', padx=5)
        
        # Sonuç metin alanı
        self.sonuc_frame = ttk.LabelFrame(self.ana_frame, text="Çıkarılan Metin", padding="10")
        self.sonuc_frame.pack(fill='both', expand=True, pady=5)
        
        self.sonuc_text = tk.Text(self.sonuc_frame, height=10, width=70)
        self.sonuc_text.pack(side='left', fill='both', expand=True)
        
        # Scrollbar
        self.scrollbar = ttk.Scrollbar(self.sonuc_frame, orient='vertical', command=self.sonuc_text.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.sonuc_text.configure(yscrollcommand=self.scrollbar.set)
        
        self.secili_dosya = None
        
    def dosya_surukle_birak(self, event):
        dosya_yolu = event.data
        if dosya_yolu.startswith('{'):
            dosya_yolu = dosya_yolu[1:-1]
        if dosya_yolu.lower().endswith('.pdf'):
            self.secili_dosya = dosya_yolu
            self.dosya_yolu_label.config(text=dosya_yolu)
            self.metin_cek_buton.config(state='normal')
        else:
            messagebox.showerror("Hata", "Lütfen sadece PDF dosyası sürükleyin!")
    
    def pdf_sec(self):
        dosya_yolu = filedialog.askopenfilename(
            filetypes=[("PDF Dosyaları", "*.pdf")]
        )
        if dosya_yolu:
            self.secili_dosya = dosya_yolu
            self.dosya_yolu_label.config(text=dosya_yolu)
            self.metin_cek_buton.config(state='normal')
    
    def metin_cek(self):
        if not self.secili_dosya:
            messagebox.showerror("Hata", "Lütfen önce bir PDF dosyası seçin!")
            return
        
        self.metin_cek_buton.config(state='disabled')
        self.ilerleme['value'] = 0
        self.sonuc_text.delete(1.0, tk.END)
        
        threading.Thread(target=self._metin_cek_thread, daemon=True).start()
    
    def _ui(self, islev, *args):
        # Tkinter yalnızca ana iş parçacığından güncellenebilir
        self.pencere.after(0, islev, *args)

    def _ilerleme_ayarla(self, deger):
        self.ilerleme['value'] = deger

    def _metin_cek_thread(self):
        format_secimi = self.format_var.get()
        try:
            pdf_okuyucu = PdfReader(self.secili_dosya)
            toplam_sayfa = len(pdf_okuyucu.pages)
            sayfalar = []

            for i, sayfa in enumerate(pdf_okuyucu.pages, 1):
                sayfalar.append(sayfa.extract_text() or "")
                self._ui(self._ilerleme_ayarla, (i / toplam_sayfa) * 100)
            metin = "\n".join(sayfalar)

            kayit_yolu = os.path.splitext(self.secili_dosya)[0] + f"_metin.{format_secimi}"
            if format_secimi == "txt":
                with open(kayit_yolu, 'w', encoding='utf-8') as f:
                    f.write(metin)
            elif format_secimi == "docx":
                from docx import Document
                doc = Document()
                for paragraf in metin.split("\n"):
                    doc.add_paragraph(paragraf)
                doc.save(kayit_yolu)
            elif format_secimi == "rtf":
                with open(kayit_yolu, 'w', encoding='ascii') as f:
                    f.write(rtf_olustur(metin))
            self._ui(self._bitti, metin, kayit_yolu, None)
        except Exception as e:
            self._ui(self._bitti, None, None, e)

    def _bitti(self, metin, kayit_yolu, hata):
        self.metin_cek_buton.config(state='normal')
        self.ilerleme['value'] = 0
        if hata:
            messagebox.showerror("Hata", f"Metin çıkarılırken bir hata oluştu:\n{hata}")
            return
        self.sonuc_text.delete(1.0, tk.END)
        self.sonuc_text.insert(tk.END, metin)
        if not metin.strip():
            messagebox.showwarning("Uyarı", "PDF'de seçilebilir metin bulunamadı (taranmış bir belge olabilir).")
        messagebox.showinfo("Başarılı", f"Metin başarıyla çıkarıldı ve kaydedildi:\n{kayit_yolu}")
    
    def baslat(self):
        self.pencere.mainloop()

if __name__ == "__main__":
    uygulama = PDFMetinCekici()
    uygulama.baslat() 
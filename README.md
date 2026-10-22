<div align="center">

# 📄 PDF Metin Çekici

**PDF dosyalarındaki metni tek tıkla çıkarıp TXT, Word (DOCX) veya RTF olarak kaydeden masaüstü uygulaması.**
Dosyanı pencereye sürükle, **Metni Çıkar**'a bas, gerisini o halletsin. Dosyaların bilgisayarından hiç çıkmaz.

![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python&logoColor=white)
![Windows | macOS | Linux](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
[![Lisans: MIT](https://img.shields.io/badge/lisans-MIT-blue)](LICENSE)

</div>

## Neler yapabilir?

| | |
| --- | --- |
| **Sürükle-bırak** | PDF'i pencereye bırak ya da **PDF Dosyası Seç** ile aç. |
| **Metin çıkarma** | Tüm sayfaların metnini sırayla çıkarır ve pencerede gösterir. |
| **Üç biçim** | Sonucu **TXT**, **DOCX** (Word) veya **RTF** olarak kaydet; Türkçe karakterler korunur. |
| **Donmayan arayüz** | Büyük PDF'ler arka planda işlenir; ilerleme çubuğundan durumu takip edersin. |
| **Akıllı uyarı** | Taranmış (resim tabanlı) bir PDF'te seçilebilir metin yoksa seni uyarır. |
| **Gizlilik** | İnternet bağlantısı gerekmez; hiçbir dosya bir yere gönderilmez. |

## Hızlı başlangıç

```bash
git clone https://github.com/EfeKesgin/pdf-metin-cekici.git
cd pdf-metin-cekici
pip install -r requirements.txt
python pdf_metin_cekici.pyw
```

> Windows'ta `pdf_metin_cekici.pyw` dosyasına çift tıklayarak da açabilirsin; `.pyw` uzantısı
> sayesinde arkada konsol penceresi açılmaz.

## Nasıl kullanılır?

1. PDF dosyasını **"Buraya PDF dosyasını sürükleyin"** alanına bırak ya da **PDF Dosyası Seç**'e tıkla.
2. **Kaydetme Formatı**'ndan TXT, DOCX veya RTF'yi seç.
3. **Metni Çıkar**'a tıkla.

Çıkarılan metin pencerede görünür ve PDF ile aynı klasöre `<dosya-adı>_metin.<biçim>` adıyla
kaydedilir. Örneğin `rapor.pdf` → `rapor_metin.docx`.

> **Taranmış PDF'ler:** Tarayıcıdan ya da fotoğraftan oluşturulan PDF'lerde metin, resim
> olarak durur ve seçilemez. Bu tür dosyalar için önce bir OCR (metin tanıma) aracı gerekir.

## English

**PDF Metin Çekici** is a simple drag-and-drop desktop app (Python/Tkinter) that extracts text
from PDF files and saves it as TXT, DOCX or RTF, keeping Turkish/Unicode characters intact.
Processing runs in the background with a progress bar, works fully offline and warns you when a
PDF has no selectable text (scanned documents need OCR). The interface is in Turkish.

---

## Geliştiriciler için

### Gereksinimler

- Python 3.9 veya üzeri (Tkinter ile birlikte; Linux'ta `python3-tk` paketi gerekebilir)
- [pypdf](https://github.com/py-pdf/pypdf): PDF okuma
- [python-docx](https://github.com/python-openxml/python-docx): Word çıktısı
- [tkinterdnd2](https://github.com/Eliav2/tkinterdnd2): sürükle-bırak desteği
- [ttkthemes](https://github.com/TkinterEP/ttkthemes): modern "arc" teması

### Nasıl çalışır?

- Metin çıkarma ayrı bir iş parçacığında yapılır; arayüz güncellemeleri `after()` ile ana
  iş parçacığına aktarılır, bu yüzden büyük dosyalarda pencere donmaz.
- RTF çıktısı, ASCII dışındaki her karakteri `\uN` kaçışıyla yazar; böylece Türkçe karakterler
  Word ve WordPad'de doğru görünür.
- DOCX çıktısında her satır ayrı bir paragraf olarak eklenir.

## Lisans

PDF Metin Çekici [MIT Lisansı](LICENSE) ile dağıtılır.

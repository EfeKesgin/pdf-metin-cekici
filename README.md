# PDF Metin Çekici

![Python](https://img.shields.io/badge/python-3.9%2B-3776AB?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
[![Lisans: MIT](https://img.shields.io/badge/lisans-MIT-blue.svg)](LICENSE)

PDF dosyalarından metin çıkarmak için geliştirilmiş, modern ve kullanıcı dostu bir masaüstü uygulaması.

## 🚀 Özellikler

- 📄 PDF dosyası seçme veya sürükle-bırak ile yükleme
- 🔍 PDF'den metin çıkarma
- 👀 Çıkarılan metni anlık görüntüleme
- 💾 Metni TXT, DOCX veya RTF formatında kaydetme (Türkçe karakterler korunur)
- 🎯 Modern ve kullanıcı dostu arayüz
- ⚡ Hızlı ve verimli çalışma
- 📊 İlerleme çubuğu ile işlem durumu takibi
- 🧵 Büyük PDF’lerde arayüz donmaz (işlem arka planda yapılır)

## 🛠️ Teknolojiler

- Python 3.9+
- pypdf
- python-docx
- tkinterdnd2
- ttkthemes (modern tema)
- Tkinter (GUI)

## 📋 Gereksinimler

```bash
pip install -r requirements.txt
```

## 🚀 Kurulum

1. Projeyi klonlayın:
```bash
git clone https://github.com/EfeKesgin/pdf-metin-cekici.git
```

2. Gerekli kütüphaneleri yükleyin:
```bash
pip install -r requirements.txt
```

3. Uygulamayı çalıştırın:
```bash
python pdf_metin_cekici.pyw
```

## 💻 Kullanım

1. "PDF Dosyasını Buraya Sürükleyin" alanına PDF dosyasını bırakın veya "PDF Dosyası Seç" butonunu kullanın
2. "Metni Çıkar" butonuna tıklayın
3. Çıkarılan metin ekranda görüntülenir ve PDF ile aynı klasöre `<dosya-adı>_metin.<format>` adıyla kaydedilir

> **Not:** Taranmış (resim tabanlı) PDF’lerde seçilebilir metin olmadığı için çıktı boş olabilir; bu durumda OCR gerekir.

## 🤝 Katkıda Bulunma

1. Bu projeyi fork edin
2. Feature branch'i oluşturun (`git checkout -b feature/amazing-feature`)
3. Değişikliklerinizi commit edin (`git commit -m 'feat: Add some amazing feature'`)
4. Branch'inizi push edin (`git push origin feature/amazing-feature`)
5. Pull Request oluşturun

## 📝 Lisans

Bu proje MIT lisansı altında lisanslanmıştır. Daha fazla bilgi için `LICENSE` dosyasına bakın.

## 📧 İletişim

Efe Kesgin - [@EfeKesgin](https://github.com/EfeKesgin)

Proje Linki: [https://github.com/EfeKesgin/pdf-metin-cekici](https://github.com/EfeKesgin/pdf-metin-cekici)

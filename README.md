# PDF Aracı

iLovePDF benzeri işleri internete dosya yüklemeden, tamamen bilgisayarında yapan bir Windows uygulaması. Dosyaların bilgisayarından çıkmaz.

## Neler yapabilir

- **Sayfa düzenle:** sayfaları sırala, döndür, sil, başka PDF ya da resim ekle
- **Metin düzenle:** PDF'teki yazıyı tıklayıp değiştir; bilgisayarda aynı yazı tipi varsa onunla yazılır
- **Birleştir / Ayır**
- **Küçült:** üç seviye, kaydetmeden önce tahmini boyutu gösterir
- **JPG → PDF / PDF → JPG**
- **PDF → Word / Word → PDF** (Word → PDF için Microsoft Word ya da LibreOffice gerekir)

## Kurulum

1. [Python](https://www.python.org/downloads/) kur ("Add python.exe to PATH" kutusunu işaretle).
2. `kurulum.bat` dosyasına çift tıkla. Gerekli kütüphaneleri kurar, uygulamayı exe'ye çevirir ve masaüstüne kısayol koyar. Kurulum için internet gerekir; sonrasında uygulama internetsiz çalışır.

Kurmadan denemek için: `pip install pywebview pymupdf pdf2docx docx2pdf pywin32 pillow fonttools`, sonra `PDF Aracı.pyw` dosyasına çift tıkla.

Kaldırmak için `kaldir.bat`.

## Yapı

| Dosya | Görevi |
| --- | --- |
| `app.py` | Pencere (pywebview) ve arayüzün çağırdığı bütün PDF işlemleri (PyMuPDF) |
| `metin.py` | Metin düzenleme: yazıyı bulma, silme ve aynı yazı tipiyle yeniden yazma |
| `arayuz.html`, `stil.css` | Arayüz (tek sayfa, saf JS) |
| `kurulum.ps1` | Kütüphaneleri kurar, PyInstaller ile exe yapar, kısayolları oluşturur |

## Lisans

[GNU AGPL v3](LICENSE). Kullandığı PDF motoru PyMuPDF AGPL lisanslı olduğu için bu proje de AGPL ile yayınlanıyor.

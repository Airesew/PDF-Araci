"""
PDF Aracı — lokal, internetsiz, modern PDF yardımcısı.
Arayüz: pywebview (Windows'un kendi Edge WebView2 motoru)
PDF motoru: PyMuPDF · Resim: Pillow · PDF→Word: pdf2docx · Word→PDF: MS Word (docx2pdf) ya da LibreOffice
"""
import base64
import io
import json
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid

UYGULAMA = "PDF Aracı"
PDF_UZ = (".pdf",)
WORD_UZ = (".docx", ".doc", ".rtf", ".odt")
RESIM_UZ = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".gif", ".jfif")

SEVIYELER = {  # ad: (hedef dpi, jpeg kalitesi)
    "hafif": (150, 80),
    "orta": (110, 65),
    "guclu": (72, 45),
}

A4 = (595.28, 841.89)
KENAR_PT = 28.35  # 10 mm


# ------------------------------------------------------------------ yardımcılar

def kaynak_yolu(ad):
    """PyInstaller ile paketlendiğinde de dosyayı bulur."""
    taban = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(taban, ad)


def ayar_klasoru():
    taban = os.getenv("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    yol = os.path.join(taban, "PdfAraci")
    os.makedirs(yol, exist_ok=True)
    return yol


_GECICI = None


def gecici_klasor():
    """Her açılışa ayrı geçici klasör; kapanınca silinir. Çökmeden kalan eski klasörler bir gün sonra temizlenir."""
    global _GECICI
    if _GECICI is None:
        taban = os.path.join(tempfile.gettempdir(), "PdfAraci")
        os.makedirs(taban, exist_ok=True)
        for ad in os.listdir(taban):
            yol = os.path.join(taban, ad)
            try:
                if time.time() - os.path.getmtime(yol) > 86400:
                    shutil.rmtree(yol) if os.path.isdir(yol) else os.remove(yol)
            except OSError:
                pass
        _GECICI = tempfile.mkdtemp(dir=taban)
    return _GECICI


def gunlugu_ayarla():
    """Konsolsuz çalışırken (pythonw / exe) sys.stdout yok; bazı kütüphaneler oraya yazmaya
    çalışıp çöküyor. Çıktıları bir günlük dosyasına yönlendir."""
    if sys.stdout is not None and sys.stderr is not None:
        return
    try:
        f = open(os.path.join(ayar_klasoru(), "gunluk.txt"), "a", encoding="utf-8", buffering=1)
    except Exception:
        f = open(os.devnull, "w")
    if sys.stdout is None:
        sys.stdout = f
    if sys.stderr is None:
        sys.stderr = f


gunlugu_ayarla()

import webview  # noqa: E402

try:
    import pymupdf as fitz  # PyMuPDF (yeni adı)
except ImportError:  # pragma: no cover
    try:
        import fitz
    except ImportError:
        fitz = None

_FITZ = threading.RLock()  # PyMuPDF aynı anda tek iş parçacığından kullanılmalı


def benzersiz(yol):
    if not os.path.exists(yol):
        return yol
    kok, uz = os.path.splitext(yol)
    i = 2
    while os.path.exists(f"{kok} ({i}){uz}"):
        i += 1
    return f"{kok} ({i}){uz}"


def guvenli_kaydet(doc, hedef, **ek):
    """Önce geçici dosyaya yazar; çağıran sonra yerine taşır."""
    klasor = os.path.dirname(os.path.abspath(hedef)) or "."
    fd, gecici = tempfile.mkstemp(suffix=".pdf", dir=klasor)
    os.close(fd)
    # Çıktılar her zaman şifresiz ve kısıtlamasız kaydedilir
    secenek = dict(garbage=4, deflate=True, clean=True, encryption=getattr(fitz, "PDF_ENCRYPT_NONE", 1))
    secenek.update(ek)
    try:
        try:
            doc.save(gecici, use_objstms=1, **secenek)
        except TypeError:  # eski PyMuPDF
            doc.save(gecici, **secenek)
        return gecici
    except Exception:
        try:
            os.remove(gecici)
        except OSError:
            pass
        raise


def yerine_tasi(gecici, hedef):
    os.replace(gecici, hedef)


def sonuc(yol, **ek):
    d = {"path": yol, "name": os.path.basename(yol), "size": os.path.getsize(yol)}
    d.update(ek)
    return d


def kucult_doc(doc, seviye):
    dpi, kalite = SEVIYELER[seviye]
    if hasattr(doc, "rewrite_images"):  # PyMuPDF 1.25+
        doc.rewrite_images(dpi_threshold=dpi + 10, dpi_target=dpi, quality=kalite,
                           lossy=True, lossless=True, bitonal=False,
                           color=True, gray=True, set_to_gray=False)
    else:
        _eski_kucult(doc, dpi, kalite)
    try:
        doc.subset_fonts()
    except Exception:
        pass


def _eski_kucult(doc, dpi, kalite):
    gorulen = set()
    for sayfa in doc:
        for bilgi in sayfa.get_images(full=True):
            xref, smask = bilgi[0], bilgi[1]
            if xref in gorulen or smask:
                continue
            gorulen.add(xref)
            try:
                kutular = sayfa.get_image_rects(xref)
                if not kutular:
                    continue
                gen_inch = max(r.width for r in kutular) / 72
                pix = fitz.Pixmap(doc, xref)
                if gen_inch <= 0 or pix.width / gen_inch <= dpi + 10:
                    continue
                if pix.n - pix.alpha >= 4:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                olcek = dpi / (pix.width / gen_inch)
                pix = fitz.Pixmap(pix, max(1, int(pix.width * olcek)), max(1, int(pix.height * olcek)), None)
                sayfa.replace_image(xref, stream=pix.tobytes("jpeg", jpg_quality=kalite))
            except Exception:
                continue


def sayfalari_kur(kaynaklar, sayfalar):
    """sayfalar: [{"doc": id, "idx": n, "rot": 0/90/180/270}] → yeni fitz belgesi."""
    yeni = fitz.open()
    acik = {}
    try:
        gruplar = []
        for s in sayfalar:
            if gruplar and gruplar[-1]["doc"] == s["doc"] and gruplar[-1]["son"] + 1 == s["idx"]:
                gruplar[-1]["son"] = s["idx"]
                gruplar[-1]["rot"].append(s.get("rot", 0))
            else:
                gruplar.append({"doc": s["doc"], "bas": s["idx"], "son": s["idx"], "rot": [s.get("rot", 0)]})
        for g in gruplar:
            if g["doc"] not in acik:
                acik[g["doc"]] = fitz.open(kaynaklar[g["doc"]])
            src = acik[g["doc"]]
            once = yeni.page_count
            yeni.insert_pdf(src, from_page=g["bas"], to_page=g["son"])
            for k, ek_rot in enumerate(g["rot"]):
                if ek_rot % 360:
                    p = yeni[once + k]
                    p.set_rotation((p.rotation + ek_rot) % 360)
        return yeni
    finally:
        for d in acik.values():
            d.close()


# ------------------------------------------------------------------ resimler

def resim_baytlari(yol, dondur=0, kucult=False, uzun_kenar=2000):
    """Resmi PDF'e konacak hale getirir: telefon yönünü düzeltir, döndürür, isterse küçültür.
    Döner: (bayt, (genişlik, yükseklik))"""
    from PIL import Image, ImageOps
    with Image.open(yol) as kaynak:
        bicim = kaynak.format
        try:
            yon = kaynak.getexif().get(0x0112, 1)
        except Exception:
            yon = 1
        im = ImageOps.exif_transpose(kaynak) if yon not in (0, 1) else kaynak.copy()
    if dondur % 360:
        im = im.rotate(-(dondur % 360), expand=True)  # saat yönünde
    buyuk = kucult and max(im.size) > uzun_kenar
    if buyuk:
        im.thumbnail((uzun_kenar, uzun_kenar), Image.LANCZOS)
    # Değişiklik yoksa JPEG'i olduğu gibi göm (kalite kaybı olmaz)
    if bicim == "JPEG" and yon in (0, 1) and not dondur % 360 and not buyuk and im.mode in ("RGB", "L"):
        with open(yol, "rb") as f:
            return f.read(), im.size
    if im.mode in ("RGBA", "LA", "P", "PA"):
        im = im.convert("RGBA")
        zemin = Image.new("RGB", im.size, "white")
        zemin.paste(im, mask=im.split()[-1])
        im = zemin
    elif im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    tampon = io.BytesIO()
    if bicim in ("PNG", "BMP", "GIF") and not kucult:
        im.save(tampon, "PNG")
    else:
        im.save(tampon, "JPEG", quality=80 if kucult else 92, optimize=True)
    return tampon.getvalue(), im.size


def resim_sayfa_boyutu(w, h, boyut):
    if boyut == "a4":
        return (A4[1], A4[0]) if w > h else A4
    k = A4[1] / max(w, h)
    return w * k, h * k


# ------------------------------------------------------------------ JS'e açılan API
# Not: pywebview, '_' ile başlamayan her özelliği JS'e açar; iç durum hep '_' ile başlar.

class Api:
    def __init__(self):
        self._pencere = None
        self._belgeler = {}  # id -> pdf yolu (resimler için geçici pdf)
        self._resimler = {}  # id -> özgün resim yolu
        self._son_dosya = os.path.join(ayar_klasoru(), "son.json")
        self._kaydedilmemis = False
        self._tahmin_no = 0
        self._dizin = None
        if fitz is not None:
            from metin import FontDizini
            self._dizin = FontDizini(os.path.join(ayar_klasoru(), "yazitipleri.json"))
            self._dizin.baslat()

    # ---- pencere / diyaloglar
    def _diyalog(self, tur, **kw):
        sabitler = getattr(webview, "FileDialog", None)
        if sabitler is not None:
            t = {"ac": sabitler.OPEN, "kaydet": sabitler.SAVE, "klasor": sabitler.FOLDER}[tur]
        else:
            t = {"ac": webview.OPEN_DIALOG, "kaydet": webview.SAVE_DIALOG, "klasor": webview.FOLDER_DIALOG}[tur]
        return self._pencere.create_file_dialog(t, **kw) or None

    def dosya_sec(self, tur="pdf", coklu=False):
        turler = {
            "pdf": ("PDF dosyaları (*.pdf)",),
            "word": ("Word dosyaları (*.docx;*.doc;*.rtf;*.odt)",),
            "resim": ("Resimler (*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.tif;*.tiff;*.gif;*.jfif)",),
            "hepsi": ("PDF Word ve resimler (*.pdf;*.docx;*.doc;*.rtf;*.odt;*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.tif;*.tiff)",),
        }[tur]
        sonuc_ = self._diyalog("ac", allow_multiple=bool(coklu), file_types=turler)
        return list(sonuc_) if sonuc_ else []

    def _kayit_yeri(self, onerilen, klasor):
        s = self._diyalog("kaydet", save_filename=onerilen, directory=klasor or "",
                          file_types=("PDF dosyası (*.pdf)",))
        if not s:
            return None
        yol = s[0] if isinstance(s, (list, tuple)) else s
        if not yol.lower().endswith(".pdf"):
            yol += ".pdf"
        return yol

    def _ilk_klasor(self, bid):
        return os.path.dirname(self._asil_yol(bid))

    def _asil_yol(self, bid):
        return self._resimler.get(bid) or self._belgeler[bid]

    def goster(self, yol):
        """Dosyayı Gezgin'de seçili gösterir."""
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", "/select,", os.path.normpath(yol)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", yol])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(yol)])
        except Exception:
            pass

    def ac(self, yol):
        """Dosyayı bilgisayardaki varsayılan programla açar (Word, PDF okuyucu...)."""
        if sys.platform.startswith("win"):
            os.startfile(yol)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", yol])
        else:
            subprocess.Popen(["xdg-open", yol])

    # ---- ayarlar (tema vb.)
    def ayar_oku(self):
        return ayarlari_oku()

    def ayar_yaz(self, degisen):
        ayar = ayarlari_oku()
        ayar.update({k: v for k, v in (degisen or {}).items() if k in ("tema",)})
        try:
            with open(AYAR_DOSYASI, "w", encoding="utf-8") as f:
                json.dump(ayar, f, ensure_ascii=False)
        except Exception:
            pass
        return ayar

    def kaydedilmemis_ayarla(self, var):
        """Arayüz, kaydedilmemiş metin düzenlemesi olup olmadığını bildirir (kapatırken sormak için)."""
        self._kaydedilmemis = bool(var)
        return True

    def baslik_tema(self, koyu):
        """Windows'un başlık çubuğunu temaya göre koyu/açık yapar."""
        hwnd = _hwnd_bul(1)
        if hwnd:
            _baslik_rengi(hwnd, bool(koyu))
        return True

    # ---- son dosyalar
    def _son_oku(self):
        try:
            with open(self._son_dosya, encoding="utf-8") as f:
                ham = json.load(f)
        except Exception:
            return []
        liste = []
        for x in ham:  # eski sürüm yalnızca yolları tutuyordu
            if isinstance(x, str):
                x = {"p": x, "t": os.path.getmtime(x) if os.path.exists(x) else 0}
            if isinstance(x, dict) and x.get("p"):
                liste.append(x)
        return liste

    def _son_yaz(self, liste):
        try:
            with open(self._son_dosya, "w", encoding="utf-8") as f:
                json.dump(liste[:30], f, ensure_ascii=False)
        except Exception:
            pass

    def son_dosyalar(self, sinir=12):
        cikti = []
        for x in self._son_oku():
            y = x["p"]
            if os.path.exists(y):
                cikti.append({"path": y, "name": os.path.basename(y), "size": os.path.getsize(y),
                              "word": y.lower().endswith(WORD_UZ), "t": x.get("t", 0)})
            if len(cikti) >= sinir:
                break
        return cikti

    def son_kaldir(self, yol):
        self._son_yaz([x for x in self._son_oku() if x["p"] != yol])
        return True

    def _son_ekle(self, yol):
        liste = [x for x in self._son_oku() if x["p"] != yol]
        self._son_yaz([{"p": yol, "t": time.time()}] + liste)

    # ---- belge yükleme
    def yukle(self, yol):
        if not os.path.exists(yol):
            raise ValueError("Dosya bulunamadı.")
        ad = os.path.basename(yol)
        boyut = os.path.getsize(yol)
        dusuk = yol.lower()
        if dusuk.endswith(WORD_UZ):
            self._son_ekle(yol)
            return {"word": True, "path": yol, "name": ad, "size": boyut}
        if dusuk.endswith(RESIM_UZ):
            return self._resim_yukle(yol)
        if not dusuk.endswith(PDF_UZ):
            raise ValueError(f"'{ad}' desteklenen bir dosya değil (PDF, Word ya da resim olmalı).")
        with _FITZ:
            with fitz.open(yol) as d:
                if d.needs_pass:
                    raise ValueError(f"'{ad}' şifreyle korunuyor, şifreli PDF'ler desteklenmiyor.")
                sayfalar = [{"w": round(p.rect.width, 1), "h": round(p.rect.height, 1)} for p in d]
        bid = uuid.uuid4().hex[:10]
        self._belgeler[bid] = yol
        self._son_ekle(yol)
        return {"word": False, "id": bid, "path": yol, "name": ad, "size": boyut, "pages": sayfalar}

    def _resim_yukle(self, yol):
        try:
            veri, (w, h) = resim_baytlari(yol)
        except ImportError:
            raise RuntimeError("Pillow kurulu değil. kurulum.bat dosyasını tekrar çalıştır.")
        except Exception:
            raise ValueError(f"'{os.path.basename(yol)}' resmi açılamadı.")
        pw, ph = resim_sayfa_boyutu(w, h, "orijinal")
        bid = uuid.uuid4().hex[:10]
        gecici = os.path.join(gecici_klasor(), f"resim_{bid}.pdf")
        with _FITZ:
            d = fitz.open()
            s = d.new_page(width=pw, height=ph)
            s.insert_image(s.rect, stream=veri)
            d.save(gecici)
            d.close()
        self._belgeler[bid] = gecici
        self._resimler[bid] = yol
        return {"word": False, "resim": True, "id": bid, "path": yol, "name": os.path.basename(yol),
                "size": os.path.getsize(yol), "pages": [{"w": round(pw, 1), "h": round(ph, 1)}]}

    def yukle_veri(self, ad, b64):
        """Sürükle-bırak yedek yolu: dosya içeriği JS'ten gelir, geçici klasöre yazılır."""
        yol = benzersiz(os.path.join(gecici_klasor(), os.path.basename(ad)))
        with open(yol, "wb") as f:
            f.write(base64.b64decode(b64))
        return self.yukle(yol)

    def _render(self, sayfa, genislik):
        z = genislik / max(sayfa.rect.width, sayfa.rect.height)
        pix = sayfa.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
        try:
            veri, tur = pix.tobytes("jpeg", jpg_quality=80), "jpeg"
        except Exception:
            veri, tur = pix.tobytes("png"), "png"
        return f"data:image/{tur};base64," + base64.b64encode(veri).decode()

    def onizlemeler(self, bid, bas, adet, genislik=220):
        cikti = []
        with _FITZ:
            with fitz.open(self._belgeler[bid]) as d:
                for i in range(bas, min(bas + adet, d.page_count)):
                    cikti.append(self._render(d[i], genislik))
        return cikti

    def buyuk_onizleme(self, bid, idx):
        return self.onizlemeler(bid, idx, 1, 1400)[0]

    # ---- sayfa düzenleme / kaydetme
    def sayfalari_kaydet(self, sayfalar, onerilen):
        if not sayfalar:
            raise ValueError("Kaydedilecek sayfa yok.")
        hedef = self._kayit_yeri(onerilen, self._ilk_klasor(sayfalar[0]["doc"]))
        if not hedef:
            return None
        with _FITZ:
            doc = sayfalari_kur(self._belgeler, sayfalar)
            try:
                gecici = guvenli_kaydet(doc, hedef)
            finally:
                doc.close()
        yerine_tasi(gecici, hedef)
        return sonuc(hedef)

    def resimleri_kaydet(self, sayfalar, onerilen, ayar):
        """JPG → PDF: her resim bir sayfa. ayar = {"boyut": "a4"|"orijinal", "kenar": bool, "kucult": bool}"""
        if not sayfalar:
            raise ValueError("Önce resim ekle.")
        hedef = self._kayit_yeri(onerilen, self._ilk_klasor(sayfalar[0]["doc"]))
        if not hedef:
            return None
        boyut = ayar.get("boyut", "a4")
        kenar = KENAR_PT if (ayar.get("kenar") and boyut == "a4") else 0
        kucult = bool(ayar.get("kucult"))
        with _FITZ:
            doc = fitz.open()
            acik = {}
            try:
                for s in sayfalar:
                    bid, rot = s["doc"], s.get("rot", 0) % 360
                    if bid in self._resimler:
                        veri, (w, h) = resim_baytlari(self._resimler[bid], rot, kucult)
                        pw, ph = resim_sayfa_boyutu(w, h, boyut)
                        sayfa = doc.new_page(width=pw, height=ph)
                        sayfa.insert_image(fitz.Rect(kenar, kenar, pw - kenar, ph - kenar),
                                           stream=veri, keep_proportion=True)
                    else:
                        if bid not in acik:
                            acik[bid] = fitz.open(self._belgeler[bid])
                        doc.insert_pdf(acik[bid], from_page=s["idx"], to_page=s["idx"])
                        if rot:
                            p = doc[-1]
                            p.set_rotation((p.rotation + rot) % 360)
                gecici = guvenli_kaydet(doc, hedef, garbage=3)
                adet = doc.page_count
            finally:
                for d in acik.values():
                    d.close()
                doc.close()
        yerine_tasi(gecici, hedef)
        return sonuc(hedef, pages=adet)

    def parcalara_ayir(self, parcalar, taban_ad):
        """parcalar: [[sayfa, ...], ...] → seçilen klasöre her parça ayrı PDF."""
        s = self._diyalog("klasor", directory=self._ilk_klasor(parcalar[0][0]["doc"]))
        if not s:
            return None
        klasor = s[0] if isinstance(s, (list, tuple)) else s
        kok = os.path.splitext(taban_ad)[0]
        yollar = []
        with _FITZ:
            for parca in parcalar:
                bas, son = parca[0]["idx"] + 1, parca[-1]["idx"] + 1
                ad = f"{kok}_s{bas}" + (f"-{son}" if son != bas else "") + ".pdf"
                hedef = benzersiz(os.path.join(klasor, ad))
                doc = sayfalari_kur(self._belgeler, parca)
                try:
                    gecici = guvenli_kaydet(doc, hedef, garbage=3)
                finally:
                    doc.close()
                yerine_tasi(gecici, hedef)
                yollar.append(hedef)
        return {"count": len(yollar), "path": yollar[0], "folder": klasor}

    def birlestir(self, sayfalar, onerilen, seviye=None):
        """sayfalar: dosya sırasına ve her dosyanın kendi sayfa sırasına göre düz liste
        [{"doc": id, "idx": n, "rot": derece}, ...]"""
        if len({s["doc"] for s in sayfalar}) < 2:
            raise ValueError("Birleştirmek için en az 2 PDF ekle.")
        hedef = self._kayit_yeri(onerilen, self._ilk_klasor(sayfalar[0]["doc"]))
        if not hedef:
            return None
        with _FITZ:
            yeni = sayfalari_kur(self._belgeler, sayfalar)
            try:
                if seviye:
                    kucult_doc(yeni, seviye)
                gecici = guvenli_kaydet(yeni, hedef)
                adet = yeni.page_count
            finally:
                yeni.close()
        yerine_tasi(gecici, hedef)
        return sonuc(hedef, pages=adet)

    # ---- küçültme
    def _belge_ac(self, bid, sayfalar=None):
        """sayfalar verilmişse (bazı sayfalar silinmiş) yalnızca onlardan yeni belge kurar."""
        if sayfalar:
            return sayfalari_kur(self._belgeler, sayfalar)
        return fitz.open(self._belgeler[bid])

    def kucultme_tahmini(self, bid, sayfalar=None):
        # Sayfa silindikçe yeni tahmin istenir; daha yenisi geldiyse eskisini bırak, kilidi boşuna tutma
        self._tahmin_no += 1
        no = self._tahmin_no
        cikti = {}
        for seviye in SEVIYELER:
            try:
                with _FITZ:
                    if no != self._tahmin_no:
                        return None
                    d = self._belge_ac(bid, sayfalar)
                    try:
                        kucult_doc(d, seviye)
                        cikti[seviye] = len(d.tobytes(garbage=4, deflate=True, clean=True))
                    finally:
                        d.close()
            except Exception:
                cikti[seviye] = None
        return cikti

    def kucult(self, bid, seviye, sayfalar=None):
        once = os.path.getsize(self._belgeler[bid])
        hedef = benzersiz(os.path.splitext(self._asil_yol(bid))[0] + "_kucuk.pdf")
        with _FITZ:
            d = self._belge_ac(bid, sayfalar)
            try:
                kucult_doc(d, seviye)
                gecici = guvenli_kaydet(d, hedef)
            finally:
                d.close()
        sonra = os.path.getsize(gecici)
        if sonra >= once and not sayfalar:
            os.remove(gecici)
            return {"path": None, "before": once, "after": once}
        yerine_tasi(gecici, hedef)
        return {"path": hedef, "name": os.path.basename(hedef), "before": once, "after": sonra}

    # ---- PDF → JPG
    def pdf_resim(self, bid, idxler, dpi=150, bicim="jpg"):
        """Seçilen sayfaları (boşsa hepsini) seçilen klasöre resim olarak kaydeder."""
        from PIL import Image
        s = self._diyalog("klasor", directory=self._ilk_klasor(bid))
        if not s:
            return None
        klasor = s[0] if isinstance(s, (list, tuple)) else s
        kok = os.path.splitext(os.path.basename(self._asil_yol(bid)))[0]
        uz = "png" if bicim == "png" else "jpg"
        dpi = int(dpi) if dpi else 150
        yollar = []
        with _FITZ:
            with fitz.open(self._belgeler[bid]) as d:
                for i in (idxler or range(d.page_count)):
                    pix = d[i].get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
                    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    hedef = benzersiz(os.path.join(klasor, f"{kok}_s{i + 1}.{uz}"))
                    if uz == "png":
                        im.save(hedef, "PNG", dpi=(dpi, dpi))
                    else:
                        im.save(hedef, "JPEG", quality=90, optimize=True, dpi=(dpi, dpi))
                    yollar.append(hedef)
        return {"count": len(yollar), "path": yollar[0] if yollar else klasor, "folder": klasor}

    # ---- metin düzenleme
    def metin_sayfa(self, bid, idx, genislik=1600):
        from metin import satirlar, font_ozeti, renk_hex, aile_ve_stil, alt_kume_onekini_sil
        with _FITZ:
            with fitz.open(self._belgeler[bid]) as d:
                s = d[idx]
                donuk = bool(s.rotation)
                liste = []
                if not donuk:
                    onbellek = {}
                    for g in satirlar(s):
                        oz = font_ozeti(d, s, g, self._dizin, onbellek)
                        ad = alt_kume_onekini_sil(g["font"])
                        _, kalin, egik = aile_ve_stil(ad)
                        liste.append({
                            "key": g["key"], "text": g["text"], "bbox": [round(v, 2) for v in g["bbox"]],
                            "origin": [round(v, 2) for v in g["origin"]], "size": round(g["size"], 2),
                            "color": renk_hex(g["color"]), "bold": bool(g["flags"] & 16) or kalin,
                            "italic": bool(g["flags"] & 2) or egik,
                            "serif": bool(g["flags"] & 4), "mono": bool(g["flags"] & 8),
                            "font": ad, "fontTur": oz["tur"], "fontAd": oz["ad"], "fontYedek": oz.get("yedek"),
                        })
                return {"img": self._render(s, genislik), "w": s.rect.width, "h": s.rect.height,
                        "satirlar": liste, "donuk": donuk}

    def metin_onizle(self, bid, idx, duzen, genislik=1600):
        from metin import uygula
        with _FITZ:
            with fitz.open(self._belgeler[bid]) as d:
                s = d[idx]
                bilgi = uygula(d, s, duzen, self._dizin)
                s = d[idx]  # içerik değişti, sayfayı tazele
                return {"img": self._render(s, genislik), "fontlar": bilgi}

    def metin_kaydet(self, bid, duzenler, onerilen):
        from metin import uygula
        hedef = self._kayit_yeri(onerilen, self._ilk_klasor(bid))
        if not hedef:
            return None
        with _FITZ:
            with fitz.open(self._belgeler[bid]) as d:
                for idx, duzen in duzenler.items():
                    uygula(d, d[int(idx)], duzen, self._dizin)
                try:
                    d.subset_fonts()  # eklenen yazı tiplerinden yalnızca kullanılan harfleri tut
                except Exception:
                    pass
                gecici = guvenli_kaydet(d, hedef)
        yerine_tasi(gecici, hedef)
        return sonuc(hedef)

    # ---- dönüştürme
    def pdf_word(self, yol, sayfalar=None):
        try:
            from pdf2docx import Converter
        except ImportError:
            raise RuntimeError("pdf2docx kurulu değil. kurulum.bat dosyasını tekrar çalıştır.")
        hedef = benzersiz(os.path.splitext(yol)[0] + ".docx")
        with _FITZ:
            cv = Converter(yol)
            try:
                if sayfalar:
                    cv.convert(hedef, pages=sorted(int(i) for i in sayfalar))
                else:
                    cv.convert(hedef)
            finally:
                cv.close()
        return sonuc(hedef)

    def word_pdf(self, yol):
        hedef = benzersiz(os.path.splitext(yol)[0] + ".pdf")
        hata = None
        if sys.platform in ("win32", "darwin"):
            com = None
            try:
                if sys.platform == "win32":
                    import pythoncom as com
                    com.CoInitialize()
                from docx2pdf import convert
                convert(yol, hedef)
                if os.path.exists(hedef):
                    return sonuc(hedef)
            except Exception as e:
                hata = e
            finally:
                if com is not None:
                    com.CoUninitialize()
        soffice = _soffice_bul()
        if soffice:
            gecici = tempfile.mkdtemp(prefix="pdfaraci_")
            # Ayrı profil: LibreOffice zaten açıksa da dönüştürme çalışır
            profil = pathlib.Path(gecici, "profil").as_uri()
            try:
                try:
                    subprocess.run([soffice, f"-env:UserInstallation={profil}", "--headless", "--convert-to", "pdf",
                                    "--outdir", gecici, yol],
                                   check=True, capture_output=True, timeout=300,
                                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                except subprocess.TimeoutExpired:
                    raise RuntimeError("LibreOffice dönüştürmeyi 5 dakikada bitiremedi.")
                except subprocess.CalledProcessError:
                    raise RuntimeError("LibreOffice dosyayı PDF'e çeviremedi. Dosya bozuk olabilir.")
                cikti = os.path.join(gecici, os.path.splitext(os.path.basename(yol))[0] + ".pdf")
                if not os.path.exists(cikti):
                    raise RuntimeError("LibreOffice PDF üretmedi. Dosya bozuk ya da desteklenmeyen bir biçimde olabilir.")
                shutil.move(cikti, hedef)
            finally:
                shutil.rmtree(gecici, ignore_errors=True)
            return sonuc(hedef)
        raise RuntimeError("Word → PDF için bilgisayarda Microsoft Word ya da ücretsiz LibreOffice kurulu olmalı."
                           + (f" (Ayrıntı: {hata})" if hata else ""))


def _soffice_bul():
    for ad in ("soffice", "libreoffice"):
        y = shutil.which(ad)
        if y:
            return y
    for y in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(y):
            return y
    return None


# ------------------------------------------------------------------ pencere ayrıntıları

def _surukle_birak_bagla(pencere):
    try:
        from webview.dom import DOMEventHandler
    except Exception:
        return  # eski pywebview: arayüz yedek yolu (dosya içeriğini gönderme) kullanır

    def birakildi(e):
        try:
            dosyalar = e.get("dataTransfer", {}).get("files", [])
            yollar = [f.get("pywebviewFullPath") for f in dosyalar if f.get("pywebviewFullPath")]
        except Exception:
            yollar = []
        if yollar:
            pencere.evaluate_js(f"window.__dosyaBirakildi({json.dumps(yollar)})")

    try:
        pencere.dom.document.events.drop += DOMEventHandler(birakildi, True, False)
        pencere.evaluate_js("window.__pyBirakma = true")
    except Exception:
        pass


AYAR_DOSYASI = os.path.join(ayar_klasoru(), "ayar.json")


def ayarlari_oku():
    try:
        with open(AYAR_DOSYASI, encoding="utf-8") as f:
            d = json.load(f)
            return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def windows_koyu_mu():
    """Windows'ta uygulamalar için koyu tema seçili mi?"""
    if not sys.platform.startswith("win"):
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except Exception:
        return False


def tema_koyu_mu():
    t = ayarlari_oku().get("tema", "sistem")
    return t == "koyu" or (t == "sistem" and windows_koyu_mu())


def _hwnd_bul(deneme=40):
    if not sys.platform.startswith("win"):
        return None
    import ctypes
    from ctypes import c_void_p, c_wchar_p
    u = ctypes.windll.user32
    u.FindWindowW.restype = c_void_p
    u.FindWindowW.argtypes = [c_wchar_p, c_wchar_p]
    for _ in range(deneme):
        h = u.FindWindowW(None, UYGULAMA)
        if h:
            return h
        time.sleep(0.1)
    return None


def _baslik_rengi(hwnd, koyu):
    try:
        import ctypes
        from ctypes import byref, c_int, c_void_p
        deger = c_int(1 if koyu else 0)
        dwm = ctypes.windll.dwmapi
        if dwm.DwmSetWindowAttribute(c_void_p(hwnd), 20, byref(deger), 4) != 0:  # Windows 11 / 10 20H1+
            dwm.DwmSetWindowAttribute(c_void_p(hwnd), 19, byref(deger), 4)        # eski Windows 10
        u = ctypes.windll.user32
        u.SetWindowPos.argtypes = [c_void_p, c_void_p, c_int, c_int, c_int, c_int, ctypes.c_uint]
        u.SetWindowPos(c_void_p(hwnd), None, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0004 | 0x0020)  # çerçeveyi yeniden çiz
    except Exception:
        pass


def _windows_kimligi():
    """Görev çubuğunda Python simgesiyle gruplanmasın."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("PdfAraci.Uygulama")
    except Exception:
        pass


def _pencere_simgesi():
    """Pencere ve görev çubuğu simgesini simge.ico yapar (exe olmadan çalışırken de)."""
    if not sys.platform.startswith("win"):
        return
    ico = kaynak_yolu("simge.ico")
    if not os.path.exists(ico):
        return
    try:
        import ctypes
        from ctypes import c_int, c_uint, c_void_p, c_wchar_p
        u = ctypes.windll.user32
        u.FindWindowW.restype = c_void_p
        u.FindWindowW.argtypes = [c_wchar_p, c_wchar_p]
        u.LoadImageW.restype = c_void_p
        u.LoadImageW.argtypes = [c_void_p, c_wchar_p, c_uint, c_int, c_int, c_uint]
        u.SendMessageW.restype = c_void_p
        u.SendMessageW.argtypes = [c_void_p, c_uint, c_void_p, c_void_p]
        hwnd = _hwnd_bul()
        if not hwnd:
            return
        _baslik_rengi(hwnd, tema_koyu_mu())
        buyuk = u.LoadImageW(None, ico, 1, u.GetSystemMetrics(11), u.GetSystemMetrics(12), 0x10)
        kucuk = u.LoadImageW(None, ico, 1, u.GetSystemMetrics(49), u.GetSystemMetrics(50), 0x10)
        if buyuk:
            u.SendMessageW(hwnd, 0x80, c_void_p(1), buyuk)
        if kucuk:
            u.SendMessageW(hwnd, 0x80, c_void_p(0), kucuk)
    except Exception:
        pass


def _kapatma_onayi(api):
    """Kaydedilmemiş metin düzenlemesi varsa kapatmadan önce sorar. False dönerse pencere kapanmaz."""
    if not api._kaydedilmemis or not sys.platform.startswith("win"):
        return True
    import ctypes
    u = ctypes.windll.user32
    u.MessageBoxW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint]
    cevap = u.MessageBoxW(_hwnd_bul(1), "Metin düzenlemede kaydedilmemiş değişiklikler var. Yine de kapatılsın mı?",
                          UYGULAMA, 0x04 | 0x30)  # Evet/Hayır + uyarı simgesi
    return cevap == 6  # Evet


def main():
    if fitz is None:
        raise SystemExit("PyMuPDF kurulu değil. Önce kurulum.bat dosyasını çalıştır.")
    _windows_kimligi()
    api = Api()
    # Arayüz dosyadan, uygulamanın kendi yerel sunucusuyla açılır (stil ve yazı tipi dosyaları için)
    pencere = webview.create_window(UYGULAMA, url=kaynak_yolu("arayuz.html"), js_api=api, width=1320, height=860,
                                    min_size=(1040, 680),
                                    background_color="#23254A" if tema_koyu_mu() else "#F5F5FC")
    api._pencere = pencere
    pencere.events.loaded += lambda: _surukle_birak_bagla(pencere)
    pencere.events.closing += lambda: _kapatma_onayi(api)
    pencere.events.shown += lambda: threading.Thread(target=_pencere_simgesi, daemon=True).start()
    try:
        webview.start(http_server=True)
    finally:
        if _GECICI:
            shutil.rmtree(_GECICI, ignore_errors=True)


if __name__ == "__main__":
    main()

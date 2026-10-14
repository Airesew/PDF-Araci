"""
Metin düzenleme: PDF'teki yazıları bulur, değiştirir ve aynı yazı tipiyle yeniden yazar.

Yazı tipi seçimi sırası:
  1. Bilgisayarda kurulu aynı yazı tipi (Arial, Calibri, Times New Roman...)  → birebir aynı görünüm
  2. PDF'in içine gömülü yazı tipi (yeni yazılan tüm harfler içinde varsa)
  3. Benzer yedek yazı tipi (tırnaklı ise Times New Roman, değilse Arial)
Hiçbir yazı tipi indirilmez; Windows'un kendi yazı tipleri kullanılır.
"""
import hashlib
import io
import json
import os
import re
import sys
import threading

try:
    import pymupdf as fitz
except ImportError:
    import fitz

# get_text bayrakları: resimleri dahil etme, bitişik harfleri (fi, fl) aç
METIN_BAYRAK = fitz.TEXT_PRESERVE_WHITESPACE | getattr(fitz, "TEXT_MEDIABOX_CLIP", 64)

TAKMA_ADLAR = {
    "helvetica": "arial", "helveticaneue": "arial", "arialmt": "arial",
    "times": "timesnewroman", "timesroman": "timesnewroman",
    "courier": "couriernew",
}
YEDEK = {
    "sans": ["arial", "liberationsans", "dejavusans", "helvetica"],
    "serif": ["timesnewroman", "liberationserif", "dejavuserif", "times"],
    "mono": ["couriernew", "liberationmono", "dejavusansmono", "courier"],
}


def alt_kume_onekini_sil(ad):
    return re.sub(r"^[A-Z]{6}\+", "", ad or "")


def normal(ad):
    """'ABCDEF+TimesNewRomanPS-BoldMT' → 'timesnewromanbold' (eşleştirme anahtarı)."""
    n = re.sub(r"[^a-z0-9]", "", alt_kume_onekini_sil(ad).lower())
    n = re.sub(r"mt$", "", n)
    n = re.sub(r"ps(?=(bold|italic|oblique|bolditalic|boldoblique)?$)", "", n)
    n = re.sub(r"regular$", "", n)
    return n


def aile_ve_stil(ad):
    """PDF yazı tipi adından aile adını ve kalın/eğik bilgisini çıkarır."""
    temiz = alt_kume_onekini_sil(ad)
    parca = re.split(r"[-,]", temiz, maxsplit=1)
    aile = normal(parca[0])
    stil = (parca[1] if len(parca) > 1 else temiz).lower()
    kalin = any(k in stil for k in ("bold", "black", "heavy", "semibold", "demi"))
    egik = "italic" in stil or "oblique" in stil
    if len(parca) == 1:  # 'ArialBold' gibi tireli olmayan adlar
        for son in ("bolditalic", "boldoblique", "bold", "italic", "oblique"):
            if aile.endswith(son) and len(aile) > len(son):
                aile = aile[: -len(son)]
                break
    return TAKMA_ADLAR.get(aile, aile), kalin, egik


def renk_rgb(c):
    return ((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255


def renk_hex(c):
    return "#%06x" % (c & 0xFFFFFF)


# ------------------------------------------------------------------ bilgisayardaki yazı tipleri

class FontDizini:
    def __init__(self, onbellek_yolu):
        self._onbellek = onbellek_yolu
        self._hazir = threading.Event()
        self._kayitlar = []
        self._ps = {}
        self._aile = {}
        self._tampon = {}
        self._fontlar = {}

    def baslat(self):
        threading.Thread(target=self._kur, daemon=True).start()

    def bekle(self, sure=60):
        self._hazir.wait(sure)

    @staticmethod
    def _klasorler():
        if sys.platform.startswith("win"):
            windir = os.environ.get("WINDIR", r"C:\Windows")
            return [os.path.join(windir, "Fonts"),
                    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")]
        if sys.platform == "darwin":
            return ["/System/Library/Fonts", "/Library/Fonts", os.path.expanduser("~/Library/Fonts")]
        return ["/usr/share/fonts", "/usr/local/share/fonts",
                os.path.expanduser("~/.fonts"), os.path.expanduser("~/.local/share/fonts")]

    def _dosyalar(self):
        cikti = []
        for k in self._klasorler():
            if not os.path.isdir(k):
                continue
            for kok, _, adlar in os.walk(k):
                for ad in adlar:
                    if ad.lower().endswith((".ttf", ".otf", ".ttc")):
                        cikti.append(os.path.join(kok, ad))
        return sorted(cikti)

    def _kur(self):
        try:
            dosyalar = self._dosyalar()
            imza = hashlib.md5("|".join(
                f"{d}:{int(os.path.getmtime(d))}" for d in dosyalar).encode("utf-8", "replace")).hexdigest()
            kayitlar = None
            try:
                with open(self._onbellek, encoding="utf-8") as f:
                    veri = json.load(f)
                if veri.get("imza") == imza and veri.get("surum") == 2:
                    kayitlar = veri["kayitlar"]
            except Exception:
                pass
            if kayitlar is None:
                kayitlar = []
                for d in dosyalar:
                    kayitlar.extend(self._oku(d))
                try:
                    with open(self._onbellek, "w", encoding="utf-8") as f:
                        json.dump({"imza": imza, "surum": 2, "kayitlar": kayitlar}, f)
                except Exception:
                    pass
            self._kayitlar = kayitlar
            for k in kayitlar:
                self._ps.setdefault(k["ps"], k)
                self._aile.setdefault(k["aile"], []).append(k)
        except Exception as e:
            print("Yazı tipi dizini kurulamadı:", e, file=sys.stderr)
        finally:
            self._hazir.set()

    @staticmethod
    def _oku(yol):
        try:
            from fontTools.ttLib import TTCollection, TTFont
        except ImportError:
            return FontDizini._oku_yedek(yol)
        cikti = []
        try:
            if yol.lower().endswith(".ttc"):
                kol = TTCollection(yol, lazy=True)
                fontlar = list(enumerate(kol.fonts))
            else:
                kol = None
                fontlar = [(0, TTFont(yol, lazy=True))]
            for i, f in fontlar:
                try:
                    ad = f["name"]
                    ps = ad.getDebugName(6) or ""
                    aile = ad.getDebugName(1) or ""
                    alt = (ad.getDebugName(2) or "").lower()
                    tam = ad.getDebugName(4) or aile
                    if not (ps or aile):
                        continue
                    cikti.append({
                        "ps": normal(ps), "aile": normal(aile), "ad": tam, "yol": yol, "n": i,
                        "b": any(k in alt for k in ("bold", "black", "heavy")),
                        "i": "italic" in alt or "oblique" in alt,
                    })
                except Exception:
                    continue
            try:
                (kol or fontlar[0][1]).close()
            except Exception:
                pass
        except Exception:
            pass
        return cikti

    @staticmethod
    def _oku_yedek(yol):
        """fontTools yoksa dosya adından tahmin (yalnızca yaygın Windows yazı tipleri)."""
        tablo = {
            "arial": ("arial", "Arial", 0, 0), "arialbd": ("arial", "Arial Bold", 1, 0),
            "ariali": ("arial", "Arial Italic", 0, 1), "arialbi": ("arial", "Arial Bold Italic", 1, 1),
            "times": ("timesnewroman", "Times New Roman", 0, 0), "timesbd": ("timesnewroman", "Times New Roman Bold", 1, 0),
            "timesi": ("timesnewroman", "Times New Roman Italic", 0, 1), "timesbi": ("timesnewroman", "Times New Roman Bold Italic", 1, 1),
            "calibri": ("calibri", "Calibri", 0, 0), "calibrib": ("calibri", "Calibri Bold", 1, 0),
            "calibrii": ("calibri", "Calibri Italic", 0, 1), "calibriz": ("calibri", "Calibri Bold Italic", 1, 1),
            "verdana": ("verdana", "Verdana", 0, 0), "verdanab": ("verdana", "Verdana Bold", 1, 0),
            "tahoma": ("tahoma", "Tahoma", 0, 0), "tahomabd": ("tahoma", "Tahoma Bold", 1, 0),
            "cour": ("couriernew", "Courier New", 0, 0), "courbd": ("couriernew", "Courier New Bold", 1, 0),
            "segoeui": ("segoeui", "Segoe UI", 0, 0), "segoeuib": ("segoeui", "Segoe UI Bold", 1, 0),
            "georgia": ("georgia", "Georgia", 0, 0), "georgiab": ("georgia", "Georgia Bold", 1, 0),
        }
        kok = os.path.splitext(os.path.basename(yol))[0].lower()
        if kok not in tablo:
            return []
        aile, ad, b, i = tablo[kok]
        return [{"ps": normal(ad.replace(" ", "")), "aile": aile, "ad": ad, "yol": yol, "n": 0, "b": bool(b), "i": bool(i)}]

    def bul(self, pdf_font_adi, kalin, egik):
        self.bekle()
        anahtar = normal(pdf_font_adi)
        if anahtar in self._ps:
            return self._ps[anahtar]
        aile, k2, e2 = aile_ve_stil(pdf_font_adi)
        return self.aile_bul(aile, kalin or k2, egik or e2)

    def aile_bul(self, aile, kalin, egik):
        self.bekle()
        adaylar = self._aile.get(aile) or self._aile.get(TAKMA_ADLAR.get(aile, ""), [])
        if not adaylar:
            return None
        return min(adaylar, key=lambda k: (k["b"] != bool(kalin)) * 2 + (k["i"] != bool(egik)))

    def yedek(self, tur, kalin, egik):
        for aile in YEDEK[tur]:
            k = self.aile_bul(aile, kalin, egik)
            if k:
                return k
        return None

    def harfler_var(self, kayit, metin):
        """Yazı tipinde metnin bütün harfleri var mı (ör. ğ, ş, İ)? Okunamazsa var sayılır."""
        anahtar = (kayit["yol"], kayit["n"])
        if anahtar not in self._fontlar:
            try:
                t = self.tampon(kayit)
                self._fontlar[anahtar] = fitz.Font(fontbuffer=t) if t else fitz.Font(fontfile=kayit["yol"])
            except Exception:
                self._fontlar[anahtar] = None
        f = self._fontlar[anahtar]
        return f is None or all(f.has_glyph(ord(c)) for c in metin if not c.isspace())

    def tampon(self, kayit):
        """.ttc koleksiyonlarındaki yazı tipini tek başına kullanılabilir hale getirir."""
        if not kayit["yol"].lower().endswith(".ttc"):
            return None
        anahtar = (kayit["yol"], kayit["n"])
        if anahtar not in self._tampon:
            from fontTools.ttLib import TTCollection
            kol = TTCollection(kayit["yol"])
            b = io.BytesIO()
            kol.fonts[kayit["n"]].save(b)
            self._tampon[anahtar] = b.getvalue()
        return self._tampon[anahtar]


# ------------------------------------------------------------------ sayfadaki yazılar

def satirlar(sayfa):
    """Düzenlenebilir metin parçaları. Aynı stildeki bitişik parçalar birleşir;
    stil değişince (ör. kalın 'Ad:' + normal 'Ahmet') ayrı parça olur."""
    d = sayfa.get_text("dict", flags=METIN_BAYRAK)
    cikti = []
    for bi, blok in enumerate(d.get("blocks", [])):
        if blok.get("type", 0) != 0:
            continue
        for li, satir in enumerate(blok.get("lines", [])):
            dx, dy = satir.get("dir", (1, 0))
            if abs(dy) > 1e-3 or dx <= 0 or satir.get("wmode", 0):
                continue
            grup = None
            for si, s in enumerate(satir.get("spans", [])):
                t = s.get("text", "")
                if not t.strip():
                    if grup is not None:
                        grup["text"] += t
                    continue
                if (grup is not None and s["font"] == grup["font"] and abs(s["size"] - grup["size"]) < 0.1
                        and s["color"] == grup["color"] and s["flags"] == grup["flags"]
                        and s["bbox"][0] - grup["bbox"][2] < s["size"] * 0.6):
                    grup["text"] += t
                    grup["bbox"] = [min(grup["bbox"][0], s["bbox"][0]), min(grup["bbox"][1], s["bbox"][1]),
                                    max(grup["bbox"][2], s["bbox"][2]), max(grup["bbox"][3], s["bbox"][3])]
                    continue
                if grup is not None:
                    cikti.append(grup)
                grup = {"key": f"{bi}.{li}.{si}", "text": t, "font": s["font"], "size": s["size"],
                        "flags": s["flags"], "color": s["color"], "origin": list(s["origin"]), "bbox": list(s["bbox"])}
            if grup is not None:
                cikti.append(grup)
    for g in cikti:
        g["text"] = g["text"].strip()
    return [g for g in cikti if g["text"]]


def _silme_alani(g):
    """Satırın tamamını değil, yalnızca harflerin orta bandını işaretler; komşu satırlara dokunmaz."""
    y = g["origin"][1]
    return fitz.Rect(g["bbox"][0], y - 0.75 * g["size"], g["bbox"][2], y - 0.1 * g["size"])


def _redakte(sayfa):
    secenek = {"images": getattr(fitz, "PDF_REDACT_IMAGE_NONE", 0)}
    try:  # çizgi/tablo kenarlıklarına dokunma (PyMuPDF 1.24.2+)
        sayfa.apply_redactions(graphics=getattr(fitz, "PDF_REDACT_LINE_ART_NONE", 0), **secenek)
    except TypeError:
        sayfa.apply_redactions(**secenek)


def _gomulu_font(belge, sayfa, pdf_font_adi):
    hedef = normal(pdf_font_adi)
    try:
        fontlar = sayfa.get_fonts(full=True)
    except Exception:
        return None
    for f in fontlar:
        xref, uzanti, basefont = f[0], f[1], f[3]
        if normal(basefont) != hedef or uzanti not in ("ttf", "otf"):
            continue
        try:
            _, uz2, _, tampon = belge.extract_font(xref)
            if tampon and uz2 in ("ttf", "otf"):
                return tampon, fitz.Font(fontbuffer=tampon)
        except Exception:
            continue
    return None


def font_sec(belge, sayfa, g, metin, dizin):
    """Döner: {"tur": sistem|gomulu|yedek, "ad": görünen ad, "yol"/"tampon"}"""
    ad = alt_kume_onekini_sil(g["font"])
    _, k2, e2 = aile_ve_stil(ad)
    kalin = bool(g["flags"] & 16) or k2
    egik = bool(g["flags"] & 2) or e2
    k = dizin.bul(ad, kalin, egik)
    if k and (metin is None or dizin.harfler_var(k, metin)):
        return {"tur": "sistem", "ad": k["ad"], "kayit": k}
    if metin is not None:
        gom = _gomulu_font(belge, sayfa, ad)
        if gom:
            tampon, font = gom
            if all(font.has_glyph(ord(c)) for c in metin):
                return {"tur": "gomulu", "ad": ad, "tampon": tampon}
    tur = "mono" if g["flags"] & 8 else ("serif" if g["flags"] & 4 else "sans")
    k = dizin.yedek(tur, kalin, egik)
    if k:
        return {"tur": "yedek", "ad": k["ad"], "kayit": k}
    return {"tur": "yedek", "ad": "Helvetica", "base14": {"sans": "helv", "serif": "tiro", "mono": "cour"}[tur]}


def font_ozeti(belge, sayfa, g, dizin, onbellek):
    """Düzenlemeden önce gösterilecek bilgi (yeni metin bilinmeden)."""
    anahtar = (g["font"], g["flags"] & (16 | 2 | 4 | 8))
    if anahtar not in onbellek:
        s = font_sec(belge, sayfa, g, None, dizin)
        if s["tur"] == "yedek" and _gomulu_font(belge, sayfa, alt_kume_onekini_sil(g["font"])):
            s = {"tur": "gomulu", "ad": alt_kume_onekini_sil(g["font"]), "yedek": s["ad"]}
        onbellek[anahtar] = {"tur": s["tur"], "ad": s["ad"], "yedek": s.get("yedek")}
    return onbellek[anahtar]


class _Yazici:
    """Bir sayfaya yazı tiplerini bir kez ekleyip metin yazar."""

    def __init__(self, sayfa, dizin):
        self.sayfa = sayfa
        self.dizin = dizin
        self.eklenen = {}
        try:
            if not sayfa.is_wrapped:
                sayfa.wrap_contents()
        except Exception:
            pass

    def yaz(self, nokta, metin, boyut, renk, secim):
        if "base14" in secim:
            self.sayfa.insert_text(nokta, metin, fontname=secim["base14"], fontsize=boyut, color=renk)
            return
        kimlik = secim["kayit"]["yol"] + f"#{secim['kayit']['n']}" if "kayit" in secim \
            else hashlib.md5(secim["tampon"]).hexdigest()
        if kimlik not in self.eklenen:
            ad = f"PdfAraci{len(self.eklenen)}x{hashlib.md5(kimlik.encode('utf-8', 'replace')).hexdigest()[:6]}"
            if "tampon" in secim:
                self.sayfa.insert_font(fontname=ad, fontbuffer=secim["tampon"])
            else:
                tampon = self.dizin.tampon(secim["kayit"])
                if tampon:
                    self.sayfa.insert_font(fontname=ad, fontbuffer=tampon)
                else:
                    self.sayfa.insert_font(fontname=ad, fontfile=secim["kayit"]["yol"])
            self.eklenen[kimlik] = ad
        self.sayfa.insert_text(nokta, metin, fontname=self.eklenen[kimlik], fontsize=boyut, color=renk)


def uygula(belge, sayfa, duzen, dizin):
    """duzen = {"mod": {anahtar: yeni_metin}, "add": [{"x","y","text","size"}]}
    Döner: {anahtar: {"tur","ad"}} — hangi yazı tipiyle yazıldığı."""
    degisen = duzen.get("mod") or {}
    eklenen = duzen.get("add") or []
    if not degisen and not eklenen:
        return {}
    if sayfa.rotation:
        raise ValueError("Döndürülmüş sayfalarda metin düzenleme desteklenmiyor.")
    bilgi = {}
    liste = []
    if degisen:
        tablo = {g["key"]: g for g in satirlar(sayfa)}
        for anahtar, yeni in degisen.items():
            g = tablo.get(anahtar)
            if g is None or yeni == g["text"]:
                continue
            sayfa.add_redact_annot(_silme_alani(g), fill=False)
            liste.append((anahtar, g, yeni))
        if liste:
            _redakte(sayfa)
    yazici = _Yazici(sayfa, dizin)
    for anahtar, g, yeni in liste:
        if not yeni.strip():
            continue
        secim = font_sec(belge, sayfa, g, yeni, dizin)
        yazici.yaz(fitz.Point(g["origin"]), yeni, g["size"], renk_rgb(g["color"]), secim)
        bilgi[anahtar] = {"tur": secim["tur"], "ad": secim["ad"]}
    for i, e in enumerate(eklenen):
        metin = (e.get("text") or "").strip()
        if not metin:
            continue
        k = dizin.yedek("sans", bool(e.get("bold")), False)
        secim = {"tur": "yedek", "ad": k["ad"], "kayit": k} if k else {"tur": "yedek", "ad": "Helvetica", "base14": "helv"}
        yazici.yaz(fitz.Point(e["x"], e["y"]), metin, float(e.get("size") or 11), (0, 0, 0), secim)
        bilgi[f"ek{i}"] = {"tur": secim["tur"], "ad": secim["ad"]}
    return bilgi

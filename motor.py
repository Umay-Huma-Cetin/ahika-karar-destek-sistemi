# -*- coding: utf-8 -*-
"""
AHİKA Karar Destek Sistemi — Analiz Motoru, Kayıt Sorguları ve Raporlama
=======================================================================
Arayüzden bağımsızdır (tkinter içermez).
"""

import os
import re
from datetime import datetime

import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder

import guvenlik as g

GEREKLI_SUTUNLAR = ['sektör', 'Hibe_Tutari', 'Istihdam_Artisi', 'Ciro_Artisi', 'Lokasyon',
                    'Sektorel_Buyume_Endeksi', 'Hibe_Turu', 'Basari_Durumu']

KRITER_ADLARI = {
    'Sektor_Kod': 'Sektör', 'Hibe_Tutari': 'Hibe Tutarı', 'Istihdam_Artisi': 'İstihdam Artışı',
    'Ciro_Artisi': 'Ciro Artışı', 'Lokasyon_Kod': 'Lokasyon', 'Sektorel_Buyume_Endeksi': 'Sektörel Büyüme',
    'Hibe_Turu_Kod': 'Hibe Türü',
}

NEVSEHIR_ILCE_KOORDINAT = {
    "Nevşehir Merkez": (38.6939, 34.7144), "Avanos": (38.7183, 34.8519), "Ürgüp": (38.6324, 34.9155),
    "Derinkuyu": (38.3683, 34.7431), "Gülşehir": (38.7497, 34.6294), "Hacıbektaş": (38.9419, 34.5492),
    "Kozaklı": (39.2275, 34.8394), "Acıgöl": (38.5714, 34.6722),
}

# =====================================================================================
# BİÇİMLENDİRME
# =====================================================================================

def tl(x, ondalik=2):
    try:
        s = f"{float(x):,.{ondalik}f}"
    except (TypeError, ValueError):
        return "-"
    return s.replace(",", "§").replace(".", ",").replace("§", ".") + " TL"


def yuzde(x):
    try:
        return "%" + f"{float(x):.1f}".replace(".", ",")
    except (TypeError, ValueError):
        return "-"


def ondalik(x, n=2):
    try:
        return f"{float(x):,.{n}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    except (TypeError, ValueError):
        return "-"


def tarih_tr(t):
    """'2026-08-03 09:42:05' -> '03.08.2026 09:42'"""
    try:
        return datetime.strptime(str(t)[:19], "%Y-%m-%d %H:%M:%S").strftime("%d.%m.%Y %H:%M")
    except ValueError:
        return str(t or "")


def sayi_coz(metin):
    """Türkçe ve İngilizce sayı biçimlerini anlar: 1.250.000 / 1,250,000 / 1250000,50"""
    m = str(metin).strip().replace(" ", "").replace("TL", "").replace("₺", "")
    if not m:
        raise ValueError("boş")
    if "." in m and "," in m:
        if m.rfind(",") > m.rfind("."):
            m = m.replace(".", "").replace(",", ".")
        else:
            m = m.replace(",", "")
    elif "," in m:
        p = m.split(",")
        m = m.replace(",", "") if (len(p) > 2 or (len(p) == 2 and len(p[1]) == 3)) else m.replace(",", ".")
    elif "." in m:
        p = m.split(".")
        if len(p) > 2 or (len(p) == 2 and len(p[1]) == 3):
            m = m.replace(".", "")
    return float(m)


def tr_ascii(text):
    return str(text).translate(str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU"))


def tr_kucuk(s):
    return str(s or "").replace("I", "ı").replace("İ", "i").lower().strip()

# =====================================================================================
# KARAR MOTORU
# =====================================================================================

class KararMotoru:
    def __init__(self):
        self.model = None
        self.le_sektor = self.le_lokasyon = self.le_turu = None
        self.X_cols = []
        self.sektorler, self.lokasyonlar, self.hibe_turleri = [], [], []
        self.dogruluk = None
        self.kayit_sayisi = 0
        self.veri_yolu = None

    @property
    def hazir(self):
        return self.model is not None

    def egit(self, yol):
        """Excel veri setinden modeli eğitir. Hata durumunda ValueError fırlatır."""
        if not os.path.exists(yol):
            raise ValueError(f"Veri seti bulunamadı:\n{yol}")
        df = pd.read_excel(yol)
        df.columns = [str(c).strip() for c in df.columns]
        eksik = [s for s in GEREKLI_SUTUNLAR if s not in df.columns]
        if eksik:
            raise ValueError("Veri setinde şu zorunlu sütunlar bulunamadı:\n• " + "\n• ".join(eksik))
        df = df.dropna(subset=GEREKLI_SUTUNLAR)
        if len(df) < 20:
            raise ValueError("Model eğitimi için veri setinde en az 20 geçerli satır olmalıdır.")

        for s in ['sektör', 'Lokasyon', 'Hibe_Turu']:
            df[s] = df[s].astype(str).str.strip()

        le_s, le_l, le_t = LabelEncoder(), LabelEncoder(), LabelEncoder()
        df['Sektor_Kod'] = le_s.fit_transform(df['sektör'])
        df['Lokasyon_Kod'] = le_l.fit_transform(df['Lokasyon'])
        df['Hibe_Turu_Kod'] = le_t.fit_transform(df['Hibe_Turu'])

        X_cols = ['Sektor_Kod', 'Hibe_Tutari', 'Istihdam_Artisi', 'Ciro_Artisi', 'Lokasyon_Kod',
                  'Sektorel_Buyume_Endeksi', 'Hibe_Turu_Kod']
        X, y = df[X_cols], df['Basari_Durumu']
        X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.2, random_state=42)
        model = RandomForestClassifier(n_estimators=200, random_state=42, n_jobs=-1)
        model.fit(X_tr, y_tr)

        self.model, self.le_sektor, self.le_lokasyon, self.le_turu = model, le_s, le_l, le_t
        self.X_cols = X_cols
        self.sektorler = sorted(df['sektör'].unique())
        self.lokasyonlar = sorted(df['Lokasyon'].unique())
        self.hibe_turleri = sorted(df['Hibe_Turu'].unique())
        self.dogruluk = float(model.score(X_te, y_te)) * 100
        self.kayit_sayisi = len(df)
        self.veri_yolu = yol
        return True

    @staticmethod
    def _bul(deger, enc):
        hedef = tr_kucuk(deger)
        for sinif in enc.classes_:
            if tr_kucuk(sinif) == hedef or tr_ascii(tr_kucuk(sinif)) == tr_ascii(hedef):
                return sinif
        raise ValueError(f"'{deger}' değeri veri setinde bulunamadı.")

    def _proba(self, sektor, hibe, istihdam, ciro, lok_kod, buyume, tur_kod):
        s_kod = self.le_sektor.transform([sektor])[0]
        X = pd.DataFrame([[s_kod, hibe, istihdam, ciro, lok_kod, buyume, tur_kod]], columns=self.X_cols)
        return float(self.model.predict_proba(X)[0][1]) * 100, int(self.model.predict(X)[0])

    def analiz(self, girdi):
        """girdi: ad_soyad, telefon, eposta, sektor, hibe, istihdam, ciro, lokasyon, buyume, turu"""
        sektor = self._bul(girdi['sektor'], self.le_sektor)
        lokasyon = self._bul(girdi['lokasyon'], self.le_lokasyon)
        turu = self._bul(girdi['turu'], self.le_turu)
        lok_kod = self.le_lokasyon.transform([lokasyon])[0]
        tur_kod = self.le_turu.transform([turu])[0]

        sonuclar = {}
        for s in self.le_sektor.classes_:
            sonuclar[s] = self._proba(s, girdi['hibe'], girdi['istihdam'], girdi['ciro'],
                                      lok_kod, girdi['buyume'], tur_kod)
        oran, tahmin = sonuclar[sektor]
        en_iyi = max(sonuclar, key=lambda k: sonuclar[k][0])

        hibe = girdi['hibe']
        verimlilik = girdi['ciro'] / hibe if hibe > 0 else 0
        ist_kap = girdi['istihdam'] / hibe * 100000 if hibe > 0 else 0

        riskler = []
        if oran < 50:
            riskler.append("Başarı olasılığı kritik eşik olan %50'nin altında.")
        if verimlilik < 1.5:
            riskler.append(f"Ciro geri dönüş katsayısı düşük ({ondalik(verimlilik)}; önerilen ≥ 1,50).")
        if ist_kap < 1.0:
            riskler.append(f"İstihdam etkisi hibe bütçesine göre zayıf (100.000 TL başına {ondalik(ist_kap)} kişi).")
        if girdi['buyume'] < 4.0:
            riskler.append("Sektörel büyüme endeksi bölge ortalamasının altında (< 4,0).")

        if oran >= 70:
            seviye, karar_onerisi = "yuksek", "ÖNERİLİR"
        elif oran >= 50:
            seviye, karar_onerisi = "orta", "KOŞULLU ÖNERİLİR"
        else:
            seviye, karar_onerisi = "dusuk", "RİSKLİ"

        yorum = (f"{girdi['ad_soyad']} tarafından {lokasyon} ilçesinde {sektor} sektörü için talep edilen "
                 f"{tl(hibe)} tutarındaki {turu} desteği için modelin öngördüğü başarı olasılığı "
                 f"{yuzde(oran)}'dir. ")
        if en_iyi != sektor and sonuclar[en_iyi][0] - oran >= 5:
            yorum += (f"Aynı parametrelerle {en_iyi} sektörü {yuzde(sonuclar[en_iyi][0])} ile daha yüksek "
                      f"başarı olasılığı göstermektedir. ")
        yorum += (f"Hibenin ciro geri dönüş katsayısı {ondalik(verimlilik)}, her 100.000 TL hibe başına "
                  f"öngörülen ek istihdam {ondalik(ist_kap)} kişidir. ")
        if seviye == "yuksek":
            yorum += "Proje, bölgesel büyüme ve kalkınma hedefleriyle güçlü bir uyum göstermektedir."
        elif seviye == "orta":
            yorum += "Proje desteklenebilir; ancak risk uyarıları değerlendirilerek bütçe optimizasyonu yapılması önerilir."
        else:
            yorum += "Yatırım hacmine göre beklenen çıktılar yetersiz kalmaktadır; duyarlılık analizi ile bütçe revizyonu önerilir."

        return {
            'sektor': sektor, 'lokasyon': lokasyon, 'turu': turu,
            'sonuclar': sonuclar, 'oran': oran, 'tahmin': tahmin, 'en_iyi': en_iyi,
            'verimlilik': verimlilik, 'istihdam_kapasitesi': ist_kap,
            'riskler': riskler, 'seviye': seviye, 'karar_onerisi': karar_onerisi, 'yorum': yorum,
        }

    def duyarlilik(self, girdi, sonuc):
        """Bütçeyi %-100..+100 aralığında %5 adımlarla değiştirip her sektör için başarıyı hesaplar."""
        lok_kod = self.le_lokasyon.transform([sonuc['lokasyon']])[0]
        tur_kod = self.le_turu.transform([sonuc['turu']])[0]
        senaryolar = []
        en_iyi = None
        for i in range(-100, 105, 5):
            carpan = 1.0 + i / 100.0
            isim = "Talep edilen (baz)" if i == 0 else (f"Bütçe %{abs(i)} azaltılırsa" if i < 0 else f"Bütçe %{i} artırılırsa")
            tutar = max(0.0, girdi['hibe'] * carpan)
            faktor = carpan if carpan > 0 else 0.1
            oranlar = {}
            for s in self.le_sektor.classes_:
                p, _ = self._proba(s, tutar, girdi['istihdam'], girdi['ciro'], lok_kod, girdi['buyume'], tur_kod)
                oranlar[s] = min(100.0, max(5.0, p * (0.8 + 0.2 * faktor)))
            satir = {'degisim': i, 'isim': isim, 'tutar': tutar, 'oranlar': oranlar}
            senaryolar.append(satir)
            if en_iyi is None or oranlar[sonuc['sektor']] > en_iyi['oranlar'][sonuc['sektor']]:
                en_iyi = satir
        return senaryolar, en_iyi

    def onem_dereceleri(self):
        if not self.hazir:
            return []
        cift = [(KRITER_ADLARI.get(c, c), float(v)) for c, v in zip(self.X_cols, self.model.feature_importances_)]
        return sorted(cift, key=lambda x: x[1])

# =====================================================================================
# KAYIT SORGULARI
# =====================================================================================

def analiz_kaydet(girdi, sonuc, kullanici):
    with g.baglanti() as conn:
        cur = conn.execute("""INSERT INTO gecmis_sorgular
            (tarih, ad_soyad, telefon, eposta, talep_edilen_sektor, hibe_tutari, istihdam, ciro_artisi,
             lokasyon, buyume_endeksi, hibe_turu, en_iyi_sektor, basari_orani, analiz_eden)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (g.simdi(), g.sifrele(girdi['ad_soyad']), g.sifrele(girdi['telefon']), g.sifrele(girdi['eposta']),
             sonuc['sektor'], girdi['hibe'], girdi['istihdam'], girdi['ciro'], sonuc['lokasyon'], girdi['buyume'],
             sonuc['turu'], sonuc['en_iyi'], sonuc['oran'], kullanici))
        kid = cur.lastrowid
    # KVKK: denetim kaydına kişisel veri yazılmaz, yalnızca kayıt numarası
    g.log_ekle(kullanici, "YAPAY ZEKA ANALİZİ",
               f"Analiz #{kid} • {sonuc['sektor']} • {sonuc['lokasyon']} • Başarı {yuzde(sonuc['oran'])}")
    return kid


def karar_kaydet(girdi, sonuc, durum, kullanici):
    with g.baglanti() as conn:
        cur = conn.execute("""INSERT INTO onaylanan_hibeler
            (tarih, ad_soyad, telefon, sektor, hibe_turu, hibe_tutari, istihdam, ciro_artisi, lokasyon,
             basari_orani, durum, eposta, karar_veren) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (g.simdi(), g.sifrele(girdi['ad_soyad']), g.sifrele(girdi['telefon']), sonuc['sektor'], sonuc['turu'],
             girdi['hibe'], girdi['istihdam'], girdi['ciro'], sonuc['lokasyon'], sonuc['oran'], durum,
             g.sifrele(girdi['eposta']), kullanici))
        kid = cur.lastrowid
    g.log_ekle(kullanici, "KURUMSAL KARAR", f"[{durum}] Karar #{kid} • {sonuc['sektor']} • {tl(girdi['hibe'])}")


def _coz_kayit(r):
    d = dict(r)
    for a in ("ad_soyad", "telefon", "eposta"):
        if a in d:
            d[a] = g.coz(d[a])
    return d


def gecmis_kayitlari():
    with g.baglanti() as conn:
        return [_coz_kayit(r) for r in conn.execute("""SELECT id, tarih, ad_soyad, telefon, eposta,
            talep_edilen_sektor AS sektor, hibe_tutari, istihdam, ciro_artisi, lokasyon, buyume_endeksi,
            hibe_turu, basari_orani, analiz_eden FROM gecmis_sorgular ORDER BY id DESC""")]


def karar_kayitlari(durum):
    with g.baglanti() as conn:
        return [_coz_kayit(r) for r in conn.execute("""SELECT id, tarih, ad_soyad, telefon, eposta, sektor,
            hibe_turu, hibe_tutari, istihdam, ciro_artisi, lokasyon, basari_orani, karar_veren
            FROM onaylanan_hibeler WHERE durum=? ORDER BY id DESC""", (durum,))]


def kayitlari_sil(tur, idler, kullanici):
    tablo = "gecmis_sorgular" if tur == "gecmis" else "onaylanan_hibeler"
    idler = [int(i) for i in idler]
    with g.baglanti() as conn:
        conn.executemany(f"DELETE FROM {tablo} WHERE id=?", [(i,) for i in idler])
    g.log_ekle(kullanici, "SİLME", f"{tablo}: {len(idler)} kayıt silindi (#{', #'.join(map(str, idler))}).")


def liste_temizle(tur, kullanici):
    with g.baglanti() as conn:
        if tur == "gecmis":
            n = conn.execute("DELETE FROM gecmis_sorgular").rowcount
        else:
            n = conn.execute("DELETE FROM onaylanan_hibeler WHERE durum=?", (tur,)).rowcount
    g.log_ekle(kullanici, "TEMİZLEME", f"{tur} listesi temizlendi ({n} kayıt).")
    return n


def bolge_verilerini_sil(kullanici):
    with g.baglanti() as conn:
        a = conn.execute("DELETE FROM gecmis_sorgular").rowcount
        b = conn.execute("DELETE FROM onaylanan_hibeler").rowcount
    g.log_ekle(kullanici, "VERİ TEMİZLİK", f"Bölge verileri sıfırlandı ({a} analiz, {b} karar).")


def bolge_istatistikleri():
    """İlçe bazında proje sayısı ve ortalama başarı (Türkçe büyük/küçük harf duyarlı)."""
    with g.baglanti() as conn:
        satirlar = conn.execute("SELECT lokasyon, basari_orani FROM gecmis_sorgular").fetchall()
    grup = {}
    for r in satirlar:
        ad = (r["lokasyon"] or "Bilinmiyor").strip()
        k = tr_kucuk(ad)
        adet, top, _ = grup.get(k, (0, 0.0, ad))
        grup[k] = (adet + 1, top + (r["basari_orani"] or 0.0), ad)
    return sorted([(v[2], v[0], v[1] / v[0]) for v in grup.values()], key=lambda x: -x[1])


def ilce_detay(ilce):
    with g.baglanti() as conn:
        satirlar = conn.execute("SELECT talep_edilen_sektor, hibe_turu, basari_orani, lokasyon FROM gecmis_sorgular").fetchall()
    kayit = [r for r in satirlar if tr_kucuk(r["lokasyon"]) == tr_kucuk(ilce)]

    def dagilim(alan):
        grp = {}
        for r in kayit:
            k = r[alan] or "Bilinmiyor"
            a, t = grp.get(k, (0, 0.0))
            grp[k] = (a + 1, t + (r["basari_orani"] or 0.0))
        return sorted([(k, a, t / a) for k, (a, t) in grp.items()], key=lambda x: -x[1])

    ort = sum((r["basari_orani"] or 0) for r in kayit) / len(kayit) if kayit else 0.0
    return len(kayit), ort, dagilim("talep_edilen_sektor"), dagilim("hibe_turu")


def ozet_istatistikler():
    with g.baglanti() as conn:
        top_analiz, ort = conn.execute("SELECT COUNT(*), AVG(basari_orani) FROM gecmis_sorgular").fetchone()
        onay_adet, onay_tutar = conn.execute(
            "SELECT COUNT(*), COALESCE(SUM(hibe_tutari),0) FROM onaylanan_hibeler WHERE durum='Onaylandi'").fetchone()
        red_adet = conn.execute("SELECT COUNT(*) FROM onaylanan_hibeler WHERE durum='Reddedildi'").fetchone()[0]
        bugun = conn.execute("SELECT COUNT(*) FROM gecmis_sorgular WHERE substr(tarih,1,10)=?",
                             (datetime.now().strftime("%Y-%m-%d"),)).fetchone()[0]
        sektor = conn.execute("""SELECT talep_edilen_sektor, COUNT(*) n, AVG(basari_orani) FROM gecmis_sorgular
                                 GROUP BY talep_edilen_sektor ORDER BY n DESC""").fetchall()
    return {
        'toplam_analiz': top_analiz or 0, 'ortalama_basari': ort or 0.0, 'bugun': bugun or 0,
        'onay_adet': onay_adet or 0, 'onay_tutar': onay_tutar or 0.0, 'red_adet': red_adet or 0,
        'sektorler': [(r[0] or "Bilinmiyor", r[1], r[2] or 0) for r in sektor],
    }


def log_kayitlari():
    with g.baglanti() as conn:
        return [dict(r) for r in conn.execute(
            "SELECT id, tarih, kullanici, islem_tipi, detay FROM denetim_loglari ORDER BY id DESC")]


def excel_yaz(kayitlar, sutunlar, yol):
    """kayitlar: dict listesi, sutunlar: [(anahtar, başlık)]. Formül enjeksiyonu temizlenir."""
    df = pd.DataFrame([{b: g.excel_hucre_temizle(k.get(a)) for a, b in sutunlar} for k in kayitlar])
    with pd.ExcelWriter(yol, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Rapor")
        ws = w.sheets["Rapor"]
        for i, (_, baslik) in enumerate(sutunlar, start=1):
            uzun = max([len(str(baslik))] + [len(str(v)) for v in df[baslik].tolist()[:200]])
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(48, uzun + 3)
        ws.freeze_panes = "A2"

# =====================================================================================
# PDF RAPORLARI (Türkçe karakter destekli)
# =====================================================================================

_PDF_FONT = None


def _pdf_font():
    """Türkçe karakterleri doğru basmak için TTF font kaydeder (Windows Arial ya da
    matplotlib ile gelen DejaVu Sans). Bulunamazsa Helvetica + ASCII'ye düşer."""
    global _PDF_FONT
    if _PDF_FONT:
        return _PDF_FONT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    adaylar = [(r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf")]
    try:
        import matplotlib
        d = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
        adaylar.append((os.path.join(d, "DejaVuSans.ttf"), os.path.join(d, "DejaVuSans-Bold.ttf")))
    except Exception:
        pass
    for normal, kalin in adaylar:
        if os.path.exists(normal) and os.path.exists(kalin):
            try:
                pdfmetrics.registerFont(TTFont("TRFont", normal))
                pdfmetrics.registerFont(TTFont("TRFont-Bold", kalin))
                from reportlab.lib.fonts import addMapping
                addMapping("TRFont", 0, 0, "TRFont")
                addMapping("TRFont", 1, 0, "TRFont-Bold")
                _PDF_FONT = ("TRFont", "TRFont-Bold", True)
                return _PDF_FONT
            except Exception:
                continue
    _PDF_FONT = ("Helvetica", "Helvetica-Bold", False)
    return _PDF_FONT


def _pdf_metin(t):
    _, _, unicode_ok = _pdf_font()
    t = str(t)
    if not unicode_ok:
        t = tr_ascii(t)
    t = re.sub(r"[\U0001F000-\U0001FFFF\u2600-\u27BF\uFE0F]", "", t)
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return t.replace("\n", "<br/>")


def _pdf_sablon(yol, baslik, alt_baslik, govde_olustur, hazirlayan):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

    normal, kalin, _ = _pdf_font()
    LACI = colors.HexColor("#0f172a")
    GRI = colors.HexColor("#64748b")
    st = {
        'baslik': ParagraphStyle('b', fontName=kalin, fontSize=16, leading=20, textColor=LACI),
        'alt': ParagraphStyle('a', fontName=normal, fontSize=10, leading=13, textColor=GRI, spaceAfter=10),
        'h2': ParagraphStyle('h', fontName=kalin, fontSize=11.5, leading=15, textColor=LACI, spaceBefore=12, spaceAfter=6),
        'n': ParagraphStyle('n', fontName=normal, fontSize=9.5, leading=13.5, textColor=colors.HexColor("#1e293b")),
        'nb': ParagraphStyle('nb', fontName=kalin, fontSize=9.5, leading=13.5, textColor=colors.HexColor("#1e293b")),
        'th': ParagraphStyle('th', fontName=kalin, fontSize=9, leading=12, textColor=colors.white),
    }

    def sayfa(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(LACI)
        canvas.rect(0, A4[1] - 8 * mm, A4[0], 8 * mm, fill=1, stroke=0)
        canvas.setFont(normal, 7.5)
        canvas.setFillColor(GRI)
        canvas.drawString(18 * mm, 10 * mm, _pdf_metin_duz(f"AHİKA Karar Destek Sistemi • Hazırlayan: {hazirlayan} • {datetime.now():%d.%m.%Y %H:%M}"))
        canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"{doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(yol, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=18 * mm, bottomMargin=18 * mm, title=baslik, author="AHİKA KDS")
    hikaye = [Paragraph(_pdf_metin("AHİLER KALKINMA AJANSI"), st['alt']),
              Paragraph(_pdf_metin(baslik), st['baslik']),
              Paragraph(_pdf_metin(alt_baslik), st['alt']), Spacer(1, 4)]
    hikaye += govde_olustur(st)
    doc.build(hikaye, onFirstPage=sayfa, onLaterPages=sayfa)


def _pdf_metin_duz(t):
    _, _, ok = _pdf_font()
    return str(t) if ok else tr_ascii(t)


def _pdf_tablo(veri, genislikler, st, vurgu_satir=None):
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle, Paragraph
    satirlar = [[Paragraph(_pdf_metin(h), st['th']) for h in veri[0]]]
    for r in veri[1:]:
        satirlar.append([Paragraph(_pdf_metin(h), st['n']) for h in r])
    t = Table(satirlar, colWidths=genislikler, repeatRows=1)
    stil = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0f172a")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f9")]),
        ('LINEBELOW', (0, 0), (-1, -1), 0.4, colors.HexColor("#e2e8f0")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5), ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]
    if vurgu_satir is not None:
        stil.append(('BACKGROUND', (0, vurgu_satir), (-1, vurgu_satir), colors.HexColor("#dbeafe")))
    t.setStyle(TableStyle(stil))
    return t


def pdf_analiz_raporu(yol, girdi, sonuc, durum, hazirlayan, kisisel_goster=True):
    from reportlab.platypus import Paragraph, Spacer
    from reportlab.lib.units import mm

    tel = girdi['telefon'] if kisisel_goster else g.telefon_maskele(girdi['telefon'])
    ep = girdi['eposta'] if kisisel_goster else g.eposta_maskele(girdi['eposta'])

    def govde(st):
        h = [Paragraph("1. Başvuru ve Proje Bilgileri", st['h2'])]
        h.append(_pdf_tablo([
            ["Parametre", "Değer"],
            ["Başvuru sahibi", girdi['ad_soyad']], ["Telefon", tel], ["E-posta", ep],
            ["Talep edilen sektör", sonuc['sektor']], ["Hibe türü", sonuc['turu']],
            ["Lokasyon", sonuc['lokasyon']], ["Hibe tutarı", tl(girdi['hibe'])],
            ["Hedef istihdam", f"{girdi['istihdam']} kişi"], ["Beklenen ciro artışı", tl(girdi['ciro'])],
            ["Sektörel büyüme endeksi", f"{girdi['buyume']:.1f}".replace(".", ",")],
            ["Kurumsal karar", durum],
        ], [60 * mm, 114 * mm], st))
        h.append(Paragraph("2. Model Sonucu", st['h2']))
        h.append(Paragraph(_pdf_metin(f"Seçilen sektör için başarı olasılığı: {yuzde(sonuc['oran'])} — Öneri: {sonuc['karar_onerisi']}"), st['nb']))
        h.append(Spacer(1, 6))
        sirali = sorted(sonuc['sonuclar'].items(), key=lambda kv: -kv[1][0])
        veri = [["Sektör", "Başarı olasılığı", "Model tahmini"]]
        vurgu = None
        for i, (s, (p, t)) in enumerate(sirali, start=1):
            veri.append([s, yuzde(p), "Başarılı" if t == 1 else "Riskli"])
            if s == sonuc['sektor']:
                vurgu = i
        h.append(_pdf_tablo(veri, [74 * mm, 50 * mm, 50 * mm], st, vurgu))
        h.append(Paragraph("3. Finansal ve Sosyal Göstergeler", st['h2']))
        h.append(_pdf_tablo([
            ["Gösterge", "Değer"],
            ["Ciro geri dönüş katsayısı (ciro / hibe)", f"{sonuc['verimlilik']:.2f}".replace(".", ",")],
            ["100.000 TL hibe başına ek istihdam", f"{sonuc['istihdam_kapasitesi']:.2f}".replace(".", ",")],
        ], [110 * mm, 64 * mm], st))
        h.append(Paragraph("4. Risk Uyarıları", st['h2']))
        if sonuc['riskler']:
            for r in sonuc['riskler']:
                h.append(Paragraph(_pdf_metin("• " + r), st['n']))
        else:
            h.append(Paragraph("Belirgin bir risk tespit edilmedi.", st['n']))
        h.append(Paragraph("5. Değerlendirme", st['h2']))
        h.append(Paragraph(_pdf_metin(sonuc['yorum']), st['n']))
        h.append(Spacer(1, 10))
        h.append(Paragraph(_pdf_metin("Not: Bu rapor karar destek amaçlıdır; nihai karar yetkili kurul tarafından verilir."), st['alt']))
        return h

    _pdf_sablon(yol, "Hibe Destek Değerlendirme Raporu", f"Rapor tarihi: {datetime.now():%d.%m.%Y %H:%M}", govde, hazirlayan)


def pdf_karsilastirma(yol, basliklar, kriterler, sonuc_metni, hazirlayan):
    from reportlab.platypus import Paragraph
    from reportlab.lib.units import mm

    def govde(st):
        return [Paragraph("Kriter Karşılaştırması", st['h2']),
                _pdf_tablo([["Kriter"] + basliklar] + [list(k) for k in kriterler], [58 * mm, 58 * mm, 58 * mm], st),
                Paragraph("Sonuç", st['h2']), Paragraph(_pdf_metin(sonuc_metni), st['nb'])]

    _pdf_sablon(yol, "Proje Karşılaştırma Raporu", f"Rapor tarihi: {datetime.now():%d.%m.%Y %H:%M}", govde, hazirlayan)

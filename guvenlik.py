# -*- coding: utf-8 -*-
"""
AHİKA Karar Destek Sistemi — Güvenlik ve Veri Katmanı
=====================================================

Bu modül arayüzden bağımsızdır (tkinter kullanmaz); bu sayede tek başına test
edilebilir. Sağladıkları:

  * Şifrelerin PBKDF2-HMAC-SHA256 + kullanıcıya özel tuz ile saklanması
    (düz metin şifre veritabanında ASLA tutulmaz)
  * Eski sürümlerdeki düz metin şifrelerin ilk açılışta otomatik olarak
    hash'e çevrilmesi ve bu hesaplara "ilk girişte şifre değiştir" zorunluluğu
  * Şifre politikası (en az 8 karakter, harf + rakam, kullanıcı adından farklı)
  * Kaba kuvvet koruması: 5 hatalı denemede hesabın 15 dakika kilitlenmesi
  * Tek kullanımlık, süreli davet kodları (veritabanında yalnızca özetleri tutulur)
  * Merkezî rol–yetki matrisi
  * Değiştirilmeye karşı iz bırakan (hash zincirli) denetim günlüğü
  * Sürümlü veritabanı göçleri (PRAGMA user_version)
  * SQLite yedekleme API'si ile tutarlı yedek + otomatik yedek rotasyonu
  * KVKK için kişisel veri maskeleme ve Excel formül enjeksiyonu temizliği
  * Başvuru sahiplerinin ad / telefon / e-posta bilgilerinin veritabanında şifreli
    (Fernet: AES-128-CBC + HMAC-SHA256) saklanması. Şifreleme anahtarı Windows DPAPI
    ile bu bilgisayara bağlı olarak korunur; kaybolmaya karşı kurtarma anahtarı vardır.
"""

import os
import re
import sys
import hmac
import glob
import sqlite3
import hashlib
import secrets
import platform
from datetime import datetime, timedelta
from contextlib import contextmanager

# =====================================================================================
# YOLLAR
# =====================================================================================

def uygulama_klasoru():
    """Veritabanı / veri seti / yedekler her zaman programın bulunduğu klasörde tutulur.
    (Eski sürüm "o an çalışılan klasörü" kullanıyordu; kısayoldan açılınca veritabanı
    başka yerde oluşabiliyordu.)"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = uygulama_klasoru()
DB_YOLU = os.path.join(APP_DIR, "analiz_gecmisi.db")
YEDEK_KLASORU = os.path.join(APP_DIR, "yedekler")
ANAHTAR_YOLU = os.path.join(APP_DIR, "veri_anahtari.bin")

# =====================================================================================
# GÜVENLİK PARAMETRELERİ
# =====================================================================================

PBKDF2_TEKRAR = 390_000          # OWASP önerisi düzeyinde
HASH_ONEK = "pbkdf2_sha256"
MAKS_HATALI_DENEME = 5
KILIT_SURESI_DK = 15
OTURUM_ZAMAN_ASIMI_DK = 15
DAVET_GECERLILIK_GUN = 7
OTOMATIK_YEDEK_SAKLAMA = 30      # en son 30 otomatik yedek tutulur
MIN_SIFRE_UZUNLUK = 8
SEMA_SURUMU = 3
SIFRELI_ONEK = "enc1:"
KISISEL_ALANLAR = {"gecmis_sorgular": ("ad_soyad", "telefon", "eposta"),
                   "onaylanan_hibeler": ("ad_soyad", "telefon", "eposta")}

ROLLER = ["Ana Yönetici", "Yönetici", "Uzman", "Stajyer"]
ROL_SEVIYE = {"Ana Yönetici": 4, "Yönetici": 3, "Uzman": 2, "Stajyer": 1}

# İşlem -> izin verilen roller. Arayüzdeki her buton bu tabloya bakar.
YETKILER = {
    "analiz":            {"Ana Yönetici", "Yönetici", "Uzman", "Stajyer"},
    "karar":             {"Ana Yönetici", "Yönetici", "Uzman"},
    "kayit_sil":         {"Ana Yönetici", "Yönetici", "Uzman"},
    "liste_temizle":     {"Ana Yönetici", "Yönetici"},
    "veri_yukle":        {"Ana Yönetici", "Yönetici"},
    "disa_aktar":        {"Ana Yönetici", "Yönetici", "Uzman"},
    "kisisel_veri_gor":  {"Ana Yönetici", "Yönetici", "Uzman"},
    "log_gor":           {"Ana Yönetici", "Yönetici"},
    "log_sil":           {"Ana Yönetici"},
    "yedekle":           {"Ana Yönetici", "Yönetici"},
    "kullanici_yonet":   {"Ana Yönetici", "Yönetici"},
    "bolge_sil":         {"Ana Yönetici", "Yönetici"},
}

YAYGIN_SIFRELER = {
    "12345678", "123456789", "1234567890", "password", "password1", "qwerty123",
    "11111111", "00000000", "abcd1234", "sifre123", "parola123", "ahika123",
    "admin123", "123qwe123", "q1w2e3r4",
}


def yetki_var(rol, islem):
    return rol in YETKILER.get(islem, set())


def yonetebilir(yonetici_rol, hedef_rol):
    """Bir yönetici yalnızca kendinden düşük seviyedeki rolleri yönetebilir.
    Ana Yönetici her rolü yönetebilir."""
    if yonetici_rol == "Ana Yönetici":
        return True
    return ROL_SEVIYE.get(yonetici_rol, 0) > ROL_SEVIYE.get(hedef_rol, 0)


def atanabilir_roller(yonetici_rol):
    return [r for r in ROLLER if yonetebilir(yonetici_rol, r)]


class GuvenlikHatasi(Exception):
    """Kullanıcıya gösterilebilecek, anlaşılır güvenlik/doğrulama hatası."""


class AnahtarEksik(GuvenlikHatasi):
    """Şifreli kişisel veriler var ama bu bilgisayarda çözme anahtarı yok/okunamıyor."""


# =====================================================================================
# ŞİFRE İŞLEMLERİ
# =====================================================================================

def sifre_hashle(sifre, tekrar=PBKDF2_TEKRAR):
    tuz = secrets.token_bytes(16)
    ozet = hashlib.pbkdf2_hmac("sha256", sifre.encode("utf-8"), tuz, tekrar)
    return f"{HASH_ONEK}${tekrar}${tuz.hex()}${ozet.hex()}"


def hash_mi(deger):
    return isinstance(deger, str) and deger.startswith(HASH_ONEK + "$")


def sifre_dogrula(sifre, kayitli):
    if not hash_mi(kayitli):
        return False
    try:
        _, tekrar, tuz_hex, ozet_hex = kayitli.split("$")
        hesap = hashlib.pbkdf2_hmac("sha256", sifre.encode("utf-8"),
                                    bytes.fromhex(tuz_hex), int(tekrar))
        return hmac.compare_digest(hesap.hex(), ozet_hex)
    except Exception:
        return False


_SAHTE_HASH = sifre_hashle(secrets.token_hex(8))


def sifre_politikasi_kontrol(sifre, kullanici_adi=""):
    """Politikaya uymayan şifre için GuvenlikHatasi fırlatır."""
    if len(sifre) < MIN_SIFRE_UZUNLUK:
        raise GuvenlikHatasi(f"Şifre en az {MIN_SIFRE_UZUNLUK} karakter olmalıdır.")
    if not re.search(r"[A-Za-zÇĞİÖŞÜçğıöşü]", sifre) or not re.search(r"\d", sifre):
        raise GuvenlikHatasi("Şifre en az bir harf ve bir rakam içermelidir.")
    if kullanici_adi and kullanici_adi.casefold() in sifre.casefold():
        raise GuvenlikHatasi("Şifre kullanıcı adını içeremez.")
    if sifre.casefold() in YAYGIN_SIFRELER:
        raise GuvenlikHatasi("Bu şifre çok yaygın; lütfen tahmin edilmesi zor bir şifre seçin.")


def sifre_gucu(sifre):
    """0-4 arası puan ve açıklama döndürür (arayüzdeki güç çubuğu için)."""
    puan = 0
    if len(sifre) >= MIN_SIFRE_UZUNLUK:
        puan += 1
    if len(sifre) >= 12:
        puan += 1
    if re.search(r"[a-zçğıöşü]", sifre) and re.search(r"[A-ZÇĞİÖŞÜ]", sifre):
        puan += 1
    if re.search(r"\d", sifre) and re.search(r"[^\wÇĞİÖŞÜçğıöşü]", sifre):
        puan += 1
    if not sifre:
        return 0, ""
    etiket = ["Çok zayıf", "Zayıf", "Orta", "İyi", "Güçlü"][puan]
    return puan, etiket


def gecici_sifre_uret():
    harfler = "ABCDEFGHJKLMNPRSTUVYZabcdefghjkmnprstuvyz"
    rakam = "23456789"
    govde = "".join(secrets.choice(harfler) for _ in range(7))
    return govde + "".join(secrets.choice(rakam) for _ in range(3))


KULLANICI_ADI_DESENI = re.compile(r"^[A-Za-z0-9ÇĞİÖŞÜçğıöşü._-]{3,32}$")


def kullanici_adi_kontrol(kadi):
    if not KULLANICI_ADI_DESENI.match(kadi or ""):
        raise GuvenlikHatasi("Kullanıcı adı 3-32 karakter olmalı; yalnızca harf, rakam, nokta, "
                             "tire ve alt çizgi içerebilir.")


# =====================================================================================
# VERİTABANI BAĞLANTISI
# =====================================================================================

@contextmanager
def baglanti():
    """Her işlem için güvenli bağlantı: hata olursa geri alır, her durumda kapatır."""
    conn = sqlite3.connect(DB_YOLU, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 5000")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def simdi():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _sutun_ekle(conn, tablo, kolon, tip):
    mevcut = {r[1] for r in conn.execute(f"PRAGMA table_info({tablo})")}
    if kolon not in mevcut:
        conn.execute(f"ALTER TABLE {tablo} ADD COLUMN {kolon} {tip}")


def veritabani_hazirla():
    """Tabloları oluşturur ve eski sürüm veritabanlarını güvenle günceller."""
    otomatik_yedek_al()
    with baglanti() as conn:
        conn.execute("PRAGMA journal_mode = WAL")
        surum = conn.execute("PRAGMA user_version").fetchone()[0]

        conn.execute("""CREATE TABLE IF NOT EXISTS kullanicilar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kullanici_adi TEXT UNIQUE,
            sifre TEXT,
            rol TEXT)""")
        conn.execute("""CREATE TABLE IF NOT EXISTS gecmis_sorgular (
            id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT)""")
        conn.execute("""CREATE TABLE IF NOT EXISTS onaylanan_hibeler (
            id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT)""")
        conn.execute("""CREATE TABLE IF NOT EXISTS denetim_loglari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tarih TEXT, kullanici TEXT, islem_tipi TEXT, detay TEXT)""")

        for kolon, tip in [("ad_soyad", "TEXT"), ("telefon", "TEXT"), ("eposta", "TEXT"),
                           ("talep_edilen_sektor", "TEXT"), ("hibe_tutari", "REAL"),
                           ("istihdam", "INTEGER"), ("ciro_artisi", "REAL"), ("lokasyon", "TEXT"),
                           ("buyume_endeksi", "REAL"), ("hibe_turu", "TEXT"),
                           ("en_iyi_sektor", "TEXT"), ("basari_orani", "REAL"),
                           ("analiz_eden", "TEXT")]:
            _sutun_ekle(conn, "gecmis_sorgular", kolon, tip)

        for kolon, tip in [("ad_soyad", "TEXT"), ("telefon", "TEXT"), ("eposta", "TEXT"),
                           ("sektor", "TEXT"), ("hibe_turu", "TEXT"), ("hibe_tutari", "REAL"),
                           ("istihdam", "INTEGER"), ("ciro_artisi", "REAL"), ("lokasyon", "TEXT"),
                           ("basari_orani", "REAL"), ("durum", "TEXT"), ("karar_veren", "TEXT")]:
            _sutun_ekle(conn, "onaylanan_hibeler", kolon, tip)

        # --- v2: güvenlik alanları ---
        for kolon, tip in [("sifre_degismeli", "INTEGER DEFAULT 0"),
                           ("basarisiz_deneme", "INTEGER DEFAULT 0"),
                           ("kilit_bitis", "TEXT"),
                           ("aktif", "INTEGER DEFAULT 1"),
                           ("olusturma", "TEXT"),
                           ("son_giris", "TEXT"),
                           ("sifre_degisim", "TEXT")]:
            _sutun_ekle(conn, "kullanicilar", kolon, tip)

        for kolon, tip in [("onceki_ozet", "TEXT"), ("ozet", "TEXT")]:
            _sutun_ekle(conn, "denetim_loglari", kolon, tip)

        conn.execute("""CREATE TABLE IF NOT EXISTS davet_kodlari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kod_ozet TEXT UNIQUE NOT NULL,
            kod_ipucu TEXT,
            rol TEXT NOT NULL,
            olusturan TEXT,
            olusturma TEXT,
            son_kullanma TEXT,
            kullanan TEXT,
            kullanma_tarihi TEXT,
            iptal INTEGER DEFAULT 0)""")
        conn.execute("""CREATE TABLE IF NOT EXISTS ayarlar (
            anahtar TEXT PRIMARY KEY, deger TEXT)""")

        conn.execute("CREATE INDEX IF NOT EXISTS ix_gecmis_tarih ON gecmis_sorgular(tarih)")
        conn.execute("CREATE INDEX IF NOT EXISTS ix_onay_durum ON onaylanan_hibeler(durum)")

        # Düz metin şifreleri hash'e çevir (geri dönüşsüz). Bu hesaplar ilk girişte
        # şifrelerini değiştirmek zorundadır, çünkü eski şifreler kaynak kodda /
        # veritabanında açıkta duruyordu.
        donusen = 0
        for r in conn.execute("SELECT id, sifre FROM kullanicilar").fetchall():
            if r["sifre"] is not None and not hash_mi(r["sifre"]):
                conn.execute("UPDATE kullanicilar SET sifre=?, sifre_degismeli=1 WHERE id=?",
                             (sifre_hashle(r["sifre"]), r["id"]))
                donusen += 1
        conn.execute("UPDATE kullanicilar SET aktif=1 WHERE aktif IS NULL")
        conn.execute("UPDATE kullanicilar SET basarisiz_deneme=0 WHERE basarisiz_deneme IS NULL")

        if surum < SEMA_SURUMU:
            conn.execute(f"PRAGMA user_version = {SEMA_SURUMU}")

    _log_zincirini_tamamla()
    if donusen:
        log_ekle("SİSTEM", "GÜVENLİK GÜNCELLEMESİ",
                 f"{donusen} hesabın düz metin şifresi güvenli özete dönüştürüldü.")
    anahtar_hazirla()          # anahtar yoksa ve şifreli veri varsa AnahtarEksik fırlatır
    s, mk = _kisisel_veri_gocu()
    if donusen or s or mk:
        try:
            c = sqlite3.connect(DB_YOLU)
            c.execute("VACUUM")  # eski düz metin sayfalarını dosyadan temizle
            c.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            c.close()
        except Exception:
            pass
    return donusen


# =====================================================================================
# KİŞİSEL VERİ ŞİFRELEME (KVKK)
# =====================================================================================

_FERNET = None


def _dpapi(veri, koru):
    """Windows DPAPI: veriyi bu bilgisayara bağlı olarak şifreler/çözer."""
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    tampon = ctypes.create_string_buffer(veri, len(veri))
    giris = BLOB(len(veri), ctypes.cast(tampon, ctypes.POINTER(ctypes.c_char)))
    cikis = BLOB()
    crypt32, kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    CRYPTPROTECT_UI_FORBIDDEN, CRYPTPROTECT_LOCAL_MACHINE = 0x1, 0x4
    if koru:
        ok = crypt32.CryptProtectData(ctypes.byref(giris), "AHIKA KDS", None, None, None,
                                      CRYPTPROTECT_UI_FORBIDDEN | CRYPTPROTECT_LOCAL_MACHINE, ctypes.byref(cikis))
    else:
        ok = crypt32.CryptUnprotectData(ctypes.byref(giris), None, None, None, None,
                                        CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(cikis))
    if not ok:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(cikis.pbData, cikis.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(cikis.pbData, ctypes.c_void_p))


def _anahtar_yaz(anahtar):
    if platform.system() == "Windows":
        try:
            icerik = b"DPAPI1:" + _dpapi(anahtar, True)
        except Exception:
            icerik = b"RAW1:" + anahtar
    else:
        icerik = b"RAW1:" + anahtar
    gecici = ANAHTAR_YOLU + ".tmp"
    with open(gecici, "wb") as f:
        f.write(icerik)
    os.replace(gecici, ANAHTAR_YOLU)
    try:
        os.chmod(ANAHTAR_YOLU, 0o600)
    except OSError:
        pass


def _anahtar_oku():
    with open(ANAHTAR_YOLU, "rb") as f:
        icerik = f.read()
    if icerik.startswith(b"DPAPI1:"):
        try:
            return _dpapi(icerik[7:], False)
        except Exception:
            raise AnahtarEksik("Veri şifreleme anahtarı bu bilgisayarda açılamadı (dosya başka bir "
                               "bilgisayardan kopyalanmış olabilir). Kurtarma anahtarını girin.")
    if icerik.startswith(b"RAW1:"):
        return icerik[5:]
    raise AnahtarEksik("Veri şifreleme anahtarı dosyası bozuk. Kurtarma anahtarını girin.")


def _sifreli_veri_var():
    with baglanti() as conn:
        for tablo in KISISEL_ALANLAR:
            if conn.execute(f"SELECT 1 FROM {tablo} WHERE ad_soyad LIKE ? LIMIT 1", (SIFRELI_ONEK + "%",)).fetchone():
                return True
    return False


def anahtar_hazirla():
    global _FERNET
    from cryptography.fernet import Fernet
    if os.path.exists(ANAHTAR_YOLU):
        anahtar = _anahtar_oku()
    else:
        if _sifreli_veri_var():
            raise AnahtarEksik("Veritabanında şifreli kişisel veriler var ancak bu klasörde veri şifreleme "
                               "anahtarı (veri_anahtari.bin) bulunamadı. Kurtarma anahtarını girin.")
        anahtar = Fernet.generate_key()
        _anahtar_yaz(anahtar)
        ayar_yaz("kurtarma_anahtari_onaylandi", "0")
    _FERNET = Fernet(anahtar)


def kurtarma_anahtari():
    """Ana Yönetici'nin güvenli bir yerde saklaması gereken anahtar (metin)."""
    if _FERNET is None:
        anahtar_hazirla()
    return _anahtar_oku().decode("ascii")


def kurtarma_anahtari_geri_yukle(metin):
    global _FERNET
    from cryptography.fernet import Fernet, InvalidToken
    anahtar = (metin or "").strip().encode("ascii", "ignore")
    try:
        f = Fernet(anahtar)
    except Exception:
        raise GuvenlikHatasi("Kurtarma anahtarı geçersiz biçimde. 44 karakterlik anahtarın tamamını yapıştırın.")
    with baglanti() as conn:
        ornek = None
        for tablo in KISISEL_ALANLAR:
            r = conn.execute(f"SELECT ad_soyad FROM {tablo} WHERE ad_soyad LIKE ? LIMIT 1", (SIFRELI_ONEK + "%",)).fetchone()
            if r:
                ornek = r[0]
                break
    if ornek:
        try:
            f.decrypt(ornek[len(SIFRELI_ONEK):].encode("ascii"))
        except InvalidToken:
            raise GuvenlikHatasi("Bu kurtarma anahtarı bu veritabanına ait değil.")
    _anahtar_yaz(anahtar)
    _FERNET = f
    ayar_yaz("kurtarma_anahtari_onaylandi", "1")
    log_ekle("SİSTEM", "ANAHTAR GERİ YÜKLEME", "Veri şifreleme anahtarı kurtarma anahtarından geri yüklendi.")


def sifrele(deger):
    if deger is None or deger == "" or (isinstance(deger, str) and deger.startswith(SIFRELI_ONEK)):
        return deger
    if _FERNET is None:
        anahtar_hazirla()
    return SIFRELI_ONEK + _FERNET.encrypt(str(deger).encode("utf-8")).decode("ascii")


def coz(deger):
    if not isinstance(deger, str) or not deger.startswith(SIFRELI_ONEK):
        return deger
    if _FERNET is None:
        anahtar_hazirla()
    try:
        return _FERNET.decrypt(deger[len(SIFRELI_ONEK):].encode("ascii")).decode("utf-8")
    except Exception:
        return "⚠ çözülemedi"


_LOG_MASKELERI = [
    (re.compile(r"(Başvuru Sahibi|Sahip):\s*[^,\n]+"), r"\1: ***"),
    (re.compile(r"Tel:\s*\d+"), "Tel: ***"),
    (re.compile(r"->\s*.+\s+vs\s+.+$"), "-> *** vs ***"),
]


def _kisisel_veri_gocu():
    """Eski düz metin kişisel verileri şifreler, eski denetim kayıtlarındaki kişisel verileri maskeler."""
    sifrelenen = 0
    with baglanti() as conn:
        for tablo, alanlar in KISISEL_ALANLAR.items():
            for r in conn.execute(f"SELECT id, {', '.join(alanlar)} FROM {tablo}").fetchall():
                yeni = [sifrele(r[a]) for a in alanlar]
                if yeni != [r[a] for a in alanlar]:
                    conn.execute(f"UPDATE {tablo} SET {', '.join(a + '=?' for a in alanlar)} WHERE id=?", (*yeni, r["id"]))
                    sifrelenen += 1

        maskelenen = 0
        for r in conn.execute("SELECT id, detay FROM denetim_loglari").fetchall():
            d = r["detay"] or ""
            for desen, yerine in _LOG_MASKELERI:
                d = desen.sub(yerine, d)
            if d != (r["detay"] or ""):
                conn.execute("UPDATE denetim_loglari SET detay=? WHERE id=?", (d, r["id"]))
                maskelenen += 1
        if maskelenen:  # içerik değişti -> zinciri baştan kur
            onceki = ""
            for r in conn.execute("SELECT * FROM denetim_loglari ORDER BY id").fetchall():
                oz = _log_ozet(onceki, r["tarih"], r["kullanici"], r["islem_tipi"], r["detay"])
                conn.execute("UPDATE denetim_loglari SET onceki_ozet=?, ozet=? WHERE id=?", (onceki, oz, r["id"]))
                onceki = oz
    if sifrelenen or maskelenen:
        log_ekle("SİSTEM", "KVKK GÜNCELLEMESİ", f"{sifrelenen} kaydın kişisel verileri şifrelendi, "
                                                f"{maskelenen} denetim kaydındaki kişisel veri maskelendi.")
    if ayar_al("yedekler_korundu") != "1":
        n = _eski_yedekleri_koru()
        ayar_yaz("yedekler_korundu", "1")
        if n:
            log_ekle("SİSTEM", "KVKK GÜNCELLEMESİ", f"{n} eski yedek dosyasındaki şifreler ve kişisel veriler korumaya alındı.")
    return sifrelenen, maskelenen


def _eski_yedekleri_koru():
    """Eski sürümün aldığı yedeklerde düz metin şifre ve kişisel veri bulunur. Yedekler silinmez;
    içlerindeki şifreler hash'lenir, kişisel veriler şifrelenir, günlükler maskelenir."""
    islenen = 0
    for yol in glob.glob(os.path.join(YEDEK_KLASORU, "*.db")):
        try:
            conn = sqlite3.connect(yol, timeout=5)
            conn.row_factory = sqlite3.Row
            tablolar = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            degisti = False
            if "kullanicilar" in tablolar:
                for r in conn.execute("SELECT id, sifre FROM kullanicilar").fetchall():
                    if r["sifre"] is not None and not hash_mi(r["sifre"]):
                        conn.execute("UPDATE kullanicilar SET sifre=? WHERE id=?", (sifre_hashle(r["sifre"], 100_000), r["id"]))
                        degisti = True
            for tablo, alanlar in KISISEL_ALANLAR.items():
                if tablo not in tablolar:
                    continue
                mevcut = [a for a in alanlar if a in {c[1] for c in conn.execute(f"PRAGMA table_info({tablo})")}]
                if not mevcut:
                    continue
                for r in conn.execute(f"SELECT id, {', '.join(mevcut)} FROM {tablo}").fetchall():
                    yeni = [sifrele(r[a]) for a in mevcut]
                    if yeni != [r[a] for a in mevcut]:
                        conn.execute(f"UPDATE {tablo} SET {', '.join(a + '=?' for a in mevcut)} WHERE id=?", (*yeni, r["id"]))
                        degisti = True
            if "denetim_loglari" in tablolar:
                for r in conn.execute("SELECT id, detay FROM denetim_loglari").fetchall():
                    d = r["detay"] or ""
                    for desen, yerine in _LOG_MASKELERI:
                        d = desen.sub(yerine, d)
                    if d != (r["detay"] or ""):
                        conn.execute("UPDATE denetim_loglari SET detay=? WHERE id=?", (d, r["id"]))
                        degisti = True
            conn.commit()
            if degisti:
                conn.execute("VACUUM")  # silinen düz metin parçaları dosyada kalmasın
                islenen += 1
            conn.close()
        except Exception as e:
            print("Yedek koruma hatası:", yol, e)
    return islenen


# =====================================================================================
# DENETİM GÜNLÜĞÜ (HASH ZİNCİRİ)
# =====================================================================================

def _log_ozet(onceki, tarih, kullanici, islem, detay):
    veri = "|".join(str(x or "") for x in (onceki, tarih, kullanici, islem, detay))
    return hashlib.sha256(veri.encode("utf-8")).hexdigest()


def _log_zincirini_tamamla():
    """Zincir alanı boş olan (eski sürümden gelen) kayıtları zincire bağlar."""
    with baglanti() as conn:
        satirlar = conn.execute("SELECT * FROM denetim_loglari ORDER BY id").fetchall()
        onceki = ""
        for r in satirlar:
            if r["ozet"]:
                onceki = r["ozet"]
                continue
            oz = _log_ozet(onceki, r["tarih"], r["kullanici"], r["islem_tipi"], r["detay"])
            conn.execute("UPDATE denetim_loglari SET onceki_ozet=?, ozet=? WHERE id=?",
                         (onceki, oz, r["id"]))
            onceki = oz


def log_ekle(kullanici, islem_tipi, detay):
    try:
        with baglanti() as conn:
            conn.execute("BEGIN IMMEDIATE")
            son = conn.execute("SELECT ozet FROM denetim_loglari ORDER BY id DESC LIMIT 1").fetchone()
            onceki = son["ozet"] if son and son["ozet"] else ""
            t = simdi()
            oz = _log_ozet(onceki, t, kullanici, islem_tipi, detay)
            conn.execute("""INSERT INTO denetim_loglari
                (tarih, kullanici, islem_tipi, detay, onceki_ozet, ozet) VALUES (?,?,?,?,?,?)""",
                         (t, kullanici, islem_tipi, detay, onceki, oz))
    except Exception as e:  # günlük yazılamasa bile uygulama durmamalı
        print("Loglama hatası:", e)


def log_zinciri_dogrula():
    """(True, None) ya da (False, bozulan_kayit_id) döndürür."""
    with baglanti() as conn:
        onceki = ""
        for r in conn.execute("SELECT * FROM denetim_loglari ORDER BY id"):
            beklenen = _log_ozet(onceki, r["tarih"], r["kullanici"], r["islem_tipi"], r["detay"])
            if (r["onceki_ozet"] or "") != onceki or r["ozet"] != beklenen:
                return False, r["id"]
            onceki = r["ozet"]
    return True, None


def loglari_temizle(yapan):
    with baglanti() as conn:
        adet = conn.execute("SELECT COUNT(*) FROM denetim_loglari").fetchone()[0]
        conn.execute("DELETE FROM denetim_loglari")
    # Silme işleminin kendisi yeni zincirin ilk kaydı olur; iz kaybolmaz.
    log_ekle(yapan, "LOG TEMİZLEME", f"{adet} adet denetim kaydı silindi.")
    return adet


# =====================================================================================
# KULLANICI / OTURUM
# =====================================================================================

def ilk_kurulum_gerekli():
    with baglanti() as conn:
        return conn.execute("SELECT COUNT(*) FROM kullanicilar").fetchone()[0] == 0


def ilk_yonetici_olustur(kadi, sifre):
    kullanici_adi_kontrol(kadi)
    sifre_politikasi_kontrol(sifre, kadi)
    with baglanti() as conn:
        if conn.execute("SELECT COUNT(*) FROM kullanicilar").fetchone()[0]:
            raise GuvenlikHatasi("Kurulum zaten tamamlanmış.")
        conn.execute("""INSERT INTO kullanicilar (kullanici_adi, sifre, rol, aktif, olusturma, sifre_degisim)
                        VALUES (?,?,?,?,?,?)""", (kadi, sifre_hashle(sifre), "Ana Yönetici", 1, simdi(), simdi()))
    log_ekle(kadi, "İLK KURULUM", "Ana Yönetici hesabı oluşturuldu.")


def giris_yap(kadi, sifre):
    """Başarılıysa kullanıcı sözlüğü döndürür, değilse GuvenlikHatasi fırlatır.
    Kullanıcı adının var olup olmadığı dışarıya sızdırılmaz."""
    kadi = (kadi or "").strip()
    genel_hata = "Kullanıcı adı veya şifre hatalı."
    with baglanti() as conn:
        r = conn.execute("SELECT * FROM kullanicilar WHERE kullanici_adi=?", (kadi,)).fetchone()
        if r is None:
            sifre_dogrula(sifre, _SAHTE_HASH)  # zamanlama farkını gizle
            raise GuvenlikHatasi(genel_hata)

        if r["kilit_bitis"]:
            bitis = datetime.strptime(r["kilit_bitis"], "%Y-%m-%d %H:%M:%S")
            if bitis > datetime.now():
                kalan = max(1, int((bitis - datetime.now()).total_seconds() // 60) + 1)
                raise GuvenlikHatasi(f"Çok sayıda hatalı deneme nedeniyle hesap kilitlendi. "
                                     f"Yaklaşık {kalan} dakika sonra tekrar deneyin veya yöneticinize başvurun.")

        if not sifre_dogrula(sifre, r["sifre"]):
            deneme = (r["basarisiz_deneme"] or 0) + 1
            if deneme >= MAKS_HATALI_DENEME:
                kilit = (datetime.now() + timedelta(minutes=KILIT_SURESI_DK)).strftime("%Y-%m-%d %H:%M:%S")
                conn.execute("UPDATE kullanicilar SET basarisiz_deneme=0, kilit_bitis=? WHERE id=?", (kilit, r["id"]))
                conn.commit()
                log_ekle(kadi, "HESAP KİLİTLENDİ", f"{MAKS_HATALI_DENEME} hatalı giriş denemesi.")
                raise GuvenlikHatasi(f"Çok sayıda hatalı deneme. Hesap {KILIT_SURESI_DK} dakika kilitlendi.")
            conn.execute("UPDATE kullanicilar SET basarisiz_deneme=? WHERE id=?", (deneme, r["id"]))
            conn.commit()
            log_ekle(kadi, "BAŞARISIZ GİRİŞ", f"Hatalı şifre ({deneme}/{MAKS_HATALI_DENEME}).")
            kalan = MAKS_HATALI_DENEME - deneme
            raise GuvenlikHatasi(genel_hata + (f" ({kalan} deneme hakkınız kaldı)" if kalan <= 2 else ""))

        if not r["aktif"]:
            raise GuvenlikHatasi("Bu hesap devre dışı bırakılmış. Yöneticinize başvurun.")

        yeni_hash = r["sifre"]
        if int(r["sifre"].split("$")[1]) < PBKDF2_TEKRAR:  # eski parametreyle hash'lenmişse güncelle
            yeni_hash = sifre_hashle(sifre)
        conn.execute("""UPDATE kullanicilar SET basarisiz_deneme=0, kilit_bitis=NULL,
                        son_giris=?, sifre=? WHERE id=?""", (simdi(), yeni_hash, r["id"]))
        kullanici = {"id": r["id"], "kullanici_adi": r["kullanici_adi"], "rol": r["rol"],
                     "sifre_degismeli": bool(r["sifre_degismeli"]), "son_giris": r["son_giris"]}
    log_ekle(kadi, "GİRİŞ", f"Kullanıcı sisteme giriş yaptı [{kullanici['rol']}].")
    return kullanici


def sifre_degistir(kadi, eski, yeni):
    with baglanti() as conn:
        r = conn.execute("SELECT * FROM kullanicilar WHERE kullanici_adi=?", (kadi,)).fetchone()
        if r is None or not sifre_dogrula(eski, r["sifre"]):
            raise GuvenlikHatasi("Mevcut şifreniz hatalı.")
        if eski == yeni:
            raise GuvenlikHatasi("Yeni şifre eskisiyle aynı olamaz.")
        sifre_politikasi_kontrol(yeni, kadi)
        conn.execute("UPDATE kullanicilar SET sifre=?, sifre_degismeli=0, sifre_degisim=? WHERE id=?",
                     (sifre_hashle(yeni), simdi(), r["id"]))
    log_ekle(kadi, "ŞİFRE DEĞİŞTİRME", "Kullanıcı şifresini güncelledi.")


def kullanicilari_listele():
    with baglanti() as conn:
        return [dict(r) for r in conn.execute(
            """SELECT id, kullanici_adi, rol, aktif, son_giris, kilit_bitis, sifre_degismeli, olusturma
               FROM kullanicilar ORDER BY id""")]


def _hedef_kullanici(conn, yapan, hedef_id):
    yapan_r = conn.execute("SELECT * FROM kullanicilar WHERE kullanici_adi=?", (yapan,)).fetchone()
    hedef = conn.execute("SELECT * FROM kullanicilar WHERE id=?", (hedef_id,)).fetchone()
    if yapan_r is None or hedef is None:
        raise GuvenlikHatasi("Kullanıcı bulunamadı.")
    if not yetki_var(yapan_r["rol"], "kullanici_yonet"):
        raise GuvenlikHatasi("Bu işlem için yetkiniz yok.")
    if hedef["id"] == yapan_r["id"]:
        raise GuvenlikHatasi("Kendi hesabınız üzerinde bu işlemi yapamazsınız.")
    if not yonetebilir(yapan_r["rol"], hedef["rol"]):
        raise GuvenlikHatasi("Kendi seviyenizdeki veya üstündeki bir hesabı yönetemezsiniz.")
    return yapan_r, hedef


def _aktif_ana_yonetici_sayisi(conn):
    return conn.execute("SELECT COUNT(*) FROM kullanicilar WHERE rol='Ana Yönetici' AND aktif=1").fetchone()[0]


def rol_degistir(yapan, hedef_id, yeni_rol):
    if yeni_rol not in ROLLER:
        raise GuvenlikHatasi("Geçersiz rol.")
    with baglanti() as conn:
        yapan_r, hedef = _hedef_kullanici(conn, yapan, hedef_id)
        if not yonetebilir(yapan_r["rol"], yeni_rol):
            raise GuvenlikHatasi("Bu rolü atama yetkiniz yok.")
        if hedef["rol"] == "Ana Yönetici" and yeni_rol != "Ana Yönetici" and _aktif_ana_yonetici_sayisi(conn) <= 1:
            raise GuvenlikHatasi("Sistemde en az bir aktif Ana Yönetici kalmalıdır.")
        conn.execute("UPDATE kullanicilar SET rol=? WHERE id=?", (yeni_rol, hedef_id))
    log_ekle(yapan, "ROL DEĞİŞİKLİĞİ", f"{hedef['kullanici_adi']}: {hedef['rol']} → {yeni_rol}")


def aktiflik_degistir(yapan, hedef_id, aktif):
    with baglanti() as conn:
        _, hedef = _hedef_kullanici(conn, yapan, hedef_id)
        if not aktif and hedef["rol"] == "Ana Yönetici" and _aktif_ana_yonetici_sayisi(conn) <= 1:
            raise GuvenlikHatasi("Sistemde en az bir aktif Ana Yönetici kalmalıdır.")
        conn.execute("UPDATE kullanicilar SET aktif=? WHERE id=?", (1 if aktif else 0, hedef_id))
    log_ekle(yapan, "HESAP DURUMU", f"{hedef['kullanici_adi']} {'etkinleştirildi' if aktif else 'devre dışı bırakıldı'}.")


def kilidi_ac(yapan, hedef_id):
    with baglanti() as conn:
        _, hedef = _hedef_kullanici(conn, yapan, hedef_id)
        conn.execute("UPDATE kullanicilar SET kilit_bitis=NULL, basarisiz_deneme=0 WHERE id=?", (hedef_id,))
    log_ekle(yapan, "KİLİT AÇMA", f"{hedef['kullanici_adi']} hesabının kilidi açıldı.")


def sifre_sifirla(yapan, hedef_id):
    """Geçici şifre üretir; kullanıcı ilk girişte değiştirmek zorundadır."""
    gecici = gecici_sifre_uret()
    with baglanti() as conn:
        _, hedef = _hedef_kullanici(conn, yapan, hedef_id)
        conn.execute("""UPDATE kullanicilar SET sifre=?, sifre_degismeli=1, kilit_bitis=NULL,
                        basarisiz_deneme=0 WHERE id=?""", (sifre_hashle(gecici), hedef_id))
    log_ekle(yapan, "ŞİFRE SIFIRLAMA", f"{hedef['kullanici_adi']} için geçici şifre oluşturuldu.")
    return gecici


def kullanici_sil(yapan, hedef_id):
    with baglanti() as conn:
        _, hedef = _hedef_kullanici(conn, yapan, hedef_id)
        if hedef["rol"] == "Ana Yönetici" and _aktif_ana_yonetici_sayisi(conn) <= 1:
            raise GuvenlikHatasi("Sistemde en az bir aktif Ana Yönetici kalmalıdır.")
        conn.execute("DELETE FROM kullanicilar WHERE id=?", (hedef_id,))
    log_ekle(yapan, "KULLANICI SİLME", f"{hedef['kullanici_adi']} [{hedef['rol']}] silindi.")


# =====================================================================================
# DAVET KODLARI (tek kullanımlık, süreli)
# =====================================================================================

_KOD_ALFABE = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"


def _kod_ozet(kod):
    return hashlib.sha256(kod.strip().upper().replace(" ", "").encode("utf-8")).hexdigest()


def davet_kodu_olustur(yapan, yapan_rol, rol, gun=DAVET_GECERLILIK_GUN):
    if rol not in ROLLER:
        raise GuvenlikHatasi("Geçersiz rol.")
    if not yetki_var(yapan_rol, "kullanici_yonet") or not yonetebilir(yapan_rol, rol):
        raise GuvenlikHatasi("Bu rol için davet kodu oluşturma yetkiniz yok.")
    parcalar = ["".join(secrets.choice(_KOD_ALFABE) for _ in range(4)) for _ in range(3)]
    kod = "AHK-" + "-".join(parcalar)
    son = (datetime.now() + timedelta(days=gun)).strftime("%Y-%m-%d %H:%M:%S")
    with baglanti() as conn:
        conn.execute("""INSERT INTO davet_kodlari (kod_ozet, kod_ipucu, rol, olusturan, olusturma, son_kullanma)
                        VALUES (?,?,?,?,?,?)""", (_kod_ozet(kod), kod[-4:], rol, yapan, simdi(), son))
    log_ekle(yapan, "DAVET KODU", f"{rol} rolü için {gun} gün geçerli davet kodu oluşturuldu (…{kod[-4:]}).")
    return kod, son


def davet_kodlarini_listele():
    with baglanti() as conn:
        satirlar = [dict(r) for r in conn.execute("SELECT * FROM davet_kodlari ORDER BY id DESC")]
    simd = simdi()
    for s in satirlar:
        if s["iptal"]:
            s["durum"] = "İptal edildi"
        elif s["kullanan"]:
            s["durum"] = f"Kullanıldı ({s['kullanan']})"
        elif s["son_kullanma"] < simd:
            s["durum"] = "Süresi doldu"
        else:
            s["durum"] = "Geçerli"
    return satirlar


def davet_kodu_iptal(yapan, kod_id):
    with baglanti() as conn:
        conn.execute("UPDATE davet_kodlari SET iptal=1 WHERE id=? AND kullanan IS NULL", (kod_id,))
    log_ekle(yapan, "DAVET KODU", f"Davet kodu #{kod_id} iptal edildi.")


def kayit_ol(kadi, sifre, kod):
    kadi = (kadi or "").strip()
    kullanici_adi_kontrol(kadi)
    sifre_politikasi_kontrol(sifre, kadi)
    with baglanti() as conn:
        conn.execute("BEGIN IMMEDIATE")
        d = conn.execute("""SELECT * FROM davet_kodlari WHERE kod_ozet=? AND iptal=0
                            AND kullanan IS NULL AND son_kullanma > ?""", (_kod_ozet(kod), simdi())).fetchone()
        if d is None:
            raise GuvenlikHatasi("Davet kodu geçersiz, kullanılmış ya da süresi dolmuş.")
        if conn.execute("SELECT 1 FROM kullanicilar WHERE kullanici_adi=?", (kadi,)).fetchone():
            raise GuvenlikHatasi("Bu kullanıcı adı zaten kullanılıyor.")
        conn.execute("""INSERT INTO kullanicilar (kullanici_adi, sifre, rol, aktif, olusturma, sifre_degisim)
                        VALUES (?,?,?,?,?,?)""", (kadi, sifre_hashle(sifre), d["rol"], 1, simdi(), simdi()))
        conn.execute("UPDATE davet_kodlari SET kullanan=?, kullanma_tarihi=? WHERE id=?", (kadi, simdi(), d["id"]))
        rol = d["rol"]
    log_ekle(kadi, "YENİ KULLANICI", f"Davet koduyla kayıt olundu [{rol}].")
    return rol


# =====================================================================================
# AYARLAR
# =====================================================================================

def ayar_al(anahtar, varsayilan=None):
    try:
        with baglanti() as conn:
            r = conn.execute("SELECT deger FROM ayarlar WHERE anahtar=?", (anahtar,)).fetchone()
            return r["deger"] if r else varsayilan
    except Exception:
        return varsayilan


def ayar_yaz(anahtar, deger):
    with baglanti() as conn:
        conn.execute("INSERT INTO ayarlar(anahtar, deger) VALUES(?,?) "
                     "ON CONFLICT(anahtar) DO UPDATE SET deger=excluded.deger", (anahtar, str(deger)))


# =====================================================================================
# YEDEKLEME
# =====================================================================================

def _yedek_al(hedef):
    """SQLite yedekleme API'si: veritabanı açıkken bile tutarlı kopya üretir."""
    kaynak = sqlite3.connect(DB_YOLU)
    hedef_db = sqlite3.connect(hedef)
    try:
        kaynak.backup(hedef_db)
    finally:
        hedef_db.close()
        kaynak.close()


def otomatik_yedek_al():
    try:
        if not os.path.exists(DB_YOLU):
            return
        os.makedirs(YEDEK_KLASORU, exist_ok=True)
        hedef = os.path.join(YEDEK_KLASORU, f"db_yedek_{datetime.now():%Y-%m-%d_%H-%M-%S}.db")
        _yedek_al(hedef)
        yedekler = sorted(glob.glob(os.path.join(YEDEK_KLASORU, "db_yedek_*.db")))
        for eski in yedekler[:-OTOMATIK_YEDEK_SAKLAMA]:
            try:
                os.remove(eski)
            except OSError:
                pass
    except Exception as e:
        print("Otomatik yedekleme hatası:", e)


def manuel_yedek_al(hedef, yapan):
    _yedek_al(hedef)
    log_ekle(yapan, "YEDEKLEME", f"Veritabanı yedeklendi: {os.path.basename(hedef)}")


# =====================================================================================
# KVKK / DIŞA AKTARIM YARDIMCILARI
# =====================================================================================

def telefon_maskele(tel):
    t = str(tel or "")
    return t[:3] + "*" * max(0, len(t) - 5) + t[-2:] if len(t) > 5 else "***"


def eposta_maskele(ep):
    ep = str(ep or "")
    if "@" not in ep:
        return "***"
    ad, alan = ep.split("@", 1)
    return (ad[:1] + "***@" + alan) if ad else "***@" + alan


def ad_maskele(ad):
    parcalar = str(ad or "").split()
    return " ".join(p[:1] + "*" * max(1, len(p) - 1) for p in parcalar) or "***"


def excel_hucre_temizle(deger):
    """Excel/CSV formül enjeksiyonunu önler (=, +, -, @ ile başlayan metinler)."""
    if isinstance(deger, str) and deger[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + deger
    return deger

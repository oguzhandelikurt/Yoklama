#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Yoklama Uygulaması
-------------------
Kurum içi dil kursları için yoklama takip uygulaması.

- 5 grup (İngilizce, Batı, Balkan, Doğu Dilleri, Çağdaş Türk Lehçeleri), her birinin
  kendi derslikleri ve kendi giriş şifresi vardır.
- Sağ üstte "Kurs Ekle", "Kursiyer Ekle" ve "Yoklama" (yönetici) butonları
  bulunur. Ders Ekle ve Kursiyer Ekle butonları da yönetici şifresi ister.
- Öğretmen, dersin başlangıç saatinden itibaren ilk 10 dakika içinde
  yoklamayı girmek zorundadır. 10 dakika içinde girilmezse sistem otomatik
  olarak "öğretmen girmedi" / devamsız olarak işaretler.
- Veriler, ayarlar dosyasında belirtilen ortak ağ klasöründeki tek bir
  SQLite dosyasında (yoklama.db) tutulur, böylece tüm bilgisayarlar aynı
  veriyi görür.
- Yoklama kayıtları, gerçek bir Excel (.xlsx) dosyası olarak, özet ve
  kursiyer bazlı detay sayfalarıyla dışa aktarılabilir.

Bu dosya Python'un standart kütüphanesini (tkinter, sqlite3, hashlib,
configparser...) ve Excel dosyası oluşturmak için "openpyxl" paketini
kullanır. openpyxl, uygulama EXE'ye dönüştürülürken (build_exe.bat ile)
otomatik olarak kurulur; son kullanıcı bilgisayarlarında ayrıca bir şey
kurulması gerekmez.
"""

import os
import sys
import ctypes
import hashlib
import sqlite3
import configparser
from datetime import datetime, timedelta, date

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL_MEVCUT = True
except ImportError:
    OPENPYXL_MEVCUT = False

try:
    from tkcalendar import DateEntry
    TKCALENDAR_MEVCUT = True
except ImportError:
    TKCALENDAR_MEVCUT = False


# --------------------------------------------------------------------------
# Sabitler
# --------------------------------------------------------------------------

APP_ADI = "Yoklama Uygulaması"

VARSAYILAN_APP_BASLIGI = "MSÜ YADYO KURSLAR MÜDÜRLÜĞÜ\nKURSİYER YOKLAMA UYGULAMASI"

VARSAYILAN_ANA_ACIKLAMA = (
    "1- Kursiyer yoklamaları kurs başlangıç saatlerini müteakip 5 dakika "
    "içerisinde girilecektir.\n"
    "2- Kursta olmayan kursiyer açıklaması seçilecektir."
)
VARSAYILAN_SAG_PANEL_BASLIGI = "YÖNETİM"

GUN_ADLARI = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

# Ders eklerken sadece hafta içi günler seçilebilir (hafta sonu ders yok).
HAFTA_ICI_GUNLERI = GUN_ADLARI[:5]  # Pazartesi .. Cuma

# Ders Ekle ekranında başlangıç/bitiş saati, saat ve dakika ayrı ayrı
# seçilebilecek şekilde iki açılır listeden oluşur (her dakika seçilebilir).
SAAT_SECENEKLERI = [f"{saat:02d}" for saat in range(24)]
DAKIKA_SECENEKLERI = [f"{dakika:02d}" for dakika in range(60)]

# Varsayılan grup adları (dersliklerden bağımsız - derslikler artık kuruma
# ait ORTAK bir havuzdur, gruba özel değildir; bkz. DERSLIK_SAYISI).
GRUP_TANIMLARI = [
    "İngilizce",
    "Batı",
    "Balkan",
    "Doğu Dilleri",
    "Çağdaş Türk Lehçeleri",
]

# Kurumdaki toplam (tüm gruplar arasında ORTAK/paylaşılan) derslik sayısı;
# ilk kurulumda "Derslik 1".."Derslik 40" olarak otomatik oluşturulur.
DERSLIK_SAYISI = 40

# "Derslik 1".."Derslik 40" havuzuna ek olarak, aynı ORTAK havuzda yer alan
# ama numaralı derslik olmayan ek mekanlar (ör. seminer salonları). İlk
# kurulumda ve daha önce kurulmuş veritabanlarında (eksikse) otomatik
# eklenir; bkz. VeriTabani.kurulumu_baslat.
EK_DERSLIKLER = ["Seminer Salonu 1", "Seminer Salonu 2"]

# Uygulama daha önce eski isimle kurulmuş bir veritabanı üzerinde
# çalışıyorsa, grup/derslik adlarını otomatik olarak yeni isme taşımak
# için kullanılır (eski_ad -> yeni_ad).
GRUP_ADI_GOCLERI = {
    "Çağatay Dilleri": "Çağdaş Türk Lehçeleri",
}

DEVAMSIZLIK_SURESI_DK = 5  # varsayılan: öğretmenin ders başından itibaren giriş yapması gereken süre
# (Her ders için bu süre artık ayrı ayrı, Ders Ekle / Düzenle ekranından
# belirlenebilir; burada sadece eski kayıtlar için varsayılan değerdir.)

# Uygulamanın çalıştığı klasör (exe olarak paketlendiğinde de doğru çalışır)
if getattr(sys, "frozen", False):
    UYGULAMA_KLASORU = os.path.dirname(sys.executable)
else:
    UYGULAMA_KLASORU = os.path.dirname(os.path.abspath(__file__))

AYAR_DOSYASI = os.path.join(UYGULAMA_KLASORU, "ayarlar.ini")


# --------------------------------------------------------------------------
# Yardımcı fonksiyonlar
# --------------------------------------------------------------------------

def sifre_hashle(sifre: str) -> str:
    return hashlib.sha256(("yoklama_tuzu::" + sifre).encode("utf-8")).hexdigest()


def sifre_dogrula(girilen: str, hash_deger: str) -> bool:
    if not hash_deger:
        return False
    return sifre_hashle(girilen) == hash_deger


def bugun_gun_index() -> int:
    """0=Pazartesi ... 6=Pazar"""
    return date.today().weekday()


def saat_str_to_time(saat_str: str):
    return datetime.strptime(saat_str, "%H:%M").time()


def saat_dakika_birlestir(saat_str, dakika_str):
    """Saat/dakika alanlarına yazılan veya seçilen değerleri "SS:DD"
    biçiminde birleştirir. Kullanıcı serbestçe hem seçip hem yazabildiği
    için "9", "09" gibi farklı yazımları da kabul eder.

    Dönüş:
      - "SS:DD"  -> geçerli bir saat girildi
      - ""       -> her iki alan da boş bırakıldı (opsiyonel alanlar için)
      - None     -> girilen değer geçersiz (aralık dışı veya sayı değil)
    """
    saat_str = (saat_str or "").strip()
    dakika_str = (dakika_str or "").strip()
    if not saat_str and not dakika_str:
        return ""
    try:
        saat = int(saat_str)
        dakika = int(dakika_str)
        if not (0 <= saat <= 23) or not (0 <= dakika <= 59):
            raise ValueError
    except (ValueError, TypeError):
        return None
    return f"{saat:02d}:{dakika:02d}"


def tarih_metnini_isoya_cevir(metin):
    """Kullanıcının GG.AA.YYYY (veya GG/AA/YYYY, ya da doğrudan YYYY-AA-GG)
    biçiminde yazdığı bir tarihi ISO (YYYY-AA-GG) biçimine çevirir.

    Dönüş:
      - "YYYY-AA-GG" -> geçerli bir tarih girildi
      - ""           -> alan boş bırakıldı (sınırsız/opsiyonel anlamına gelir)
      - None         -> girilen değer geçersiz bir tarih
    """
    metin = (metin or "").strip()
    if not metin:
        return ""
    for bicim in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(metin, bicim).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def iso_tarihi_goruntu_metnine_cevir(iso_metin):
    """ISO (YYYY-AA-GG) tarihi ekranda göstermek için GG.AA.YYYY yapar."""
    iso_metin = (iso_metin or "").strip()
    if not iso_metin:
        return ""
    try:
        return datetime.strptime(iso_metin, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return iso_metin


def _devamsiz_kursiyer_satiri(kayit_row):
    """Excel'de devamsız kursiyer listesinde bir satır: 'KOD  -  Ad Soyad  (Mazeret)'.
    Ad soyad boşsa atlanır, mazeret girilmemişse parantez eklenmez."""
    parca = kayit_row["kod"]
    if kayit_row["ad_soyad"]:
        parca += f"  -  {kayit_row['ad_soyad']}"
    if kayit_row["mazeret_metin"]:
        parca += f"  ({kayit_row['mazeret_metin']})"
    return parca


def ders_tarih_araliginda_mi(ders_row, tarih_str):
    """Bir dersin, verilen (YYYY-AA-GG) tarihinde aktif olup olmadığını
    kontrol eder. Başlangıç/bitiş tarihi boşsa o yönde sınır yok demektir."""
    bas = ders_row["baslangic_tarihi"] if "baslangic_tarihi" in ders_row.keys() else None
    bit = ders_row["bitis_tarihi"] if "bitis_tarihi" in ders_row.keys() else None
    if bas and tarih_str < bas:
        return False
    if bit and tarih_str > bit:
        return False
    return True


def ders_gununde_ders_var_mi(ders_row, tarih_str, db):
    """ders_tarih_araliginda_mi ile AYNI kontrolü yapar, EK olarak verilen
    tarihin ya genel bir RESMİ TATİL ya da bu kursun kendi ARA TATİL
    aralığına denk gelip gelmediğine bakar - denk geliyorsa o gün ders
    YOKTUR (kurs kendi tarih aralığında olsa bile). "ders_row" hem
    "dersler" tablosundan (id kolonu) hem de ders_saatleri ile birleşmiş
    satırlardan (ders_id kolonu) gelebildiği için ikisi de desteklenir."""
    if not ders_tarih_araliginda_mi(ders_row, tarih_str):
        return False
    anahtarlar = ders_row.keys()
    ders_id = ders_row["id"] if "id" in anahtarlar else ders_row["ders_id"]
    if db.tarih_tatil_mi(tarih_str, ders_id):
        return False
    return True


# --------------------------------------------------------------------------
# Ayar dosyası (ortak ağ klasörünün yolu buradan okunur)
# --------------------------------------------------------------------------

def veri_dosyasi_yolunu_al() -> str:
    cfg = configparser.ConfigParser()
    if os.path.exists(AYAR_DOSYASI):
        cfg.read(AYAR_DOSYASI, encoding="utf-8")
        if cfg.has_option("ayarlar", "veri_klasoru"):
            klasor = cfg.get("ayarlar", "veri_klasoru")
            if klasor:
                return os.path.join(klasor, "yoklama.db")
    return ""


def veri_dosyasi_yolunu_kaydet(klasor: str):
    cfg = configparser.ConfigParser()
    cfg["ayarlar"] = {"veri_klasoru": klasor}
    with open(AYAR_DOSYASI, "w", encoding="utf-8") as f:
        cfg.write(f)


def ortak_klasoru_sec_ve_kaydet(parent) -> str:
    """İlk çalıştırmada (veya klasör bulunamazsa) kullanıcıdan ortak ağ
    klasörünü seçmesini ister ve ayarlar.ini içine kaydeder."""
    messagebox.showinfo(
        APP_ADI,
        "Yoklama verilerinin saklanacağı ORTAK AĞ KLASÖRÜNÜ seçmeniz gerekiyor.\n\n"
        "Bu klasör, kurumdaki tüm bilgisayarların erişebildiği paylaşımlı bir "
        "ağ klasörü olmalıdır (örneğin \\\\sunucu\\yoklama gibi). Böylece tüm "
        "bilgisayarlar aynı verileri görür.",
        parent=parent,
    )
    while True:
        klasor = filedialog.askdirectory(
            parent=parent, title="Ortak veri klasörünü seçin"
        )
        if not klasor:
            cevap = messagebox.askretrycancel(
                APP_ADI,
                "Bir klasör seçmediniz. Uygulamanın çalışması için bir veri "
                "klasörü gereklidir. Tekrar denemek ister misiniz?",
                parent=parent,
            )
            if cevap:
                continue
            else:
                sys.exit(0)
        try:
            with open(os.path.join(klasor, ".yazma_testi"), "w") as t:
                t.write("test")
            os.remove(os.path.join(klasor, ".yazma_testi"))
        except Exception as e:
            messagebox.showerror(
                APP_ADI,
                f"Seçilen klasöre yazılamıyor:\n{klasor}\n\nHata: {e}\n\n"
                "Lütfen yazma izniniz olan başka bir klasör seçin.",
                parent=parent,
            )
            continue
        veri_dosyasi_yolunu_kaydet(klasor)
        return os.path.join(klasor, "yoklama.db")


# --------------------------------------------------------------------------
# Veritabanı katmanı
# --------------------------------------------------------------------------

class VeriTabani:
    def __init__(self, db_yolu: str):
        self.db_yolu = db_yolu

    def baglan(self):
        # timeout: ağ klasöründe geçici kilitlenmelere karşı bekleme süresi
        conn = sqlite3.connect(self.db_yolu, timeout=15)
        conn.execute("PRAGMA foreign_keys = ON")
        conn.row_factory = sqlite3.Row
        return conn

    def calistir(self, sorgu, params=(), commit=False, retries=5):
        """Ağ paylaşımlı dosyalarda oluşabilecek 'database is locked'
        hatalarına karşı birkaç kez tekrar dener."""
        son_hata = None
        for deneme in range(retries):
            try:
                conn = self.baglan()
                try:
                    cur = conn.execute(sorgu, params)
                    if commit:
                        conn.commit()
                    sonuc = cur.fetchall()
                    return sonuc
                finally:
                    conn.close()
            except sqlite3.OperationalError as e:
                son_hata = e
                if "locked" in str(e).lower() or "busy" in str(e).lower():
                    import time
                    time.sleep(0.5 * (deneme + 1))
                    continue
                raise
        raise son_hata

    def kurulumu_baslat(self):
        conn = self.baglan()
        try:
            c = conn.cursor()

            # Eskiden (haftalık program modeli eklenmeden önce) bir kurs tek
            # bir gün/saate sahipti ("dersler" tablosunda "gun" kolonu vardı).
            # Artık bir kurs, bir derslikte her gün (Pzt-Per aynı, Cuma ayrı)
            # 6 periyotluk bir programa sahip; bu yüzden eski şemadaki kurs/
            # kursiyer/yoklama verileri (grup ve derslik verileri KORUNARAK)
            # burada temizlenip yeni şemayla sıfırdan kuruluyor.
            eski_ders_kolonlari = [
                r[1] for r in c.execute("PRAGMA table_info(dersler)").fetchall()
            ]
            # Eskiden her derslik belirli bir gruba aitti ("derslikler"
            # tablosunda "grup_id" kolonu vardı). Artık derslikler kurumun
            # TAMAMINA ait, gruplar arasında ORTAK paylaşılan bir havuz
            # (ör. Derslik 1..40); bu yüzden eski gruba-özel derslik verisi
            # de (ona bağlı kurs verileriyle birlikte) burada temizlenip
            # yeni ortak havuz olarak sıfırdan kuruluyor.
            eski_derslik_kolonlari = [
                r[1] for r in c.execute("PRAGMA table_info(derslikler)").fetchall()
            ]
            gecis_gerekli = ("gun" in eski_ders_kolonlari) or ("grup_id" in eski_derslik_kolonlari)
            if gecis_gerekli:
                c.executescript(
                    """
                    DROP TABLE IF EXISTS yoklama_kayitlari;
                    DROP TABLE IF EXISTS oturumlar;
                    DROP TABLE IF EXISTS kursiyerler;
                    DROP TABLE IF EXISTS ders_saatleri;
                    DROP TABLE IF EXISTS dersler;
                    """
                )
                if "grup_id" in eski_derslik_kolonlari:
                    c.execute("DROP TABLE IF EXISTS derslikler")

            c.executescript(
                """
                CREATE TABLE IF NOT EXISTS ayarlar (
                    anahtar TEXT PRIMARY KEY,
                    deger TEXT
                );
                CREATE TABLE IF NOT EXISTS gruplar (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ad TEXT UNIQUE NOT NULL,
                    sifre_hash TEXT
                );
                CREATE TABLE IF NOT EXISTS derslikler (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ad TEXT UNIQUE NOT NULL
                );
                CREATE TABLE IF NOT EXISTS dersler (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    grup_id INTEGER NOT NULL REFERENCES gruplar(id),
                    derslik_id INTEGER NOT NULL REFERENCES derslikler(id),
                    ad TEXT NOT NULL,
                    baslangic_tarihi TEXT,
                    bitis_tarihi TEXT,
                    devamsizlik_suresi_dk INTEGER
                );
                CREATE TABLE IF NOT EXISTS ders_saatleri (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ders_id INTEGER NOT NULL REFERENCES dersler(id),
                    gun INTEGER NOT NULL,
                    sira INTEGER NOT NULL,
                    baslangic TEXT NOT NULL,
                    bitis TEXT NOT NULL,
                    UNIQUE(ders_id, gun, sira)
                );
                CREATE TABLE IF NOT EXISTS kursiyerler (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ders_id INTEGER NOT NULL REFERENCES dersler(id),
                    kod TEXT NOT NULL,
                    ad_soyad TEXT NOT NULL,
                    UNIQUE(ders_id, kod)
                );
                CREATE TABLE IF NOT EXISTS oturumlar (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ders_id INTEGER NOT NULL REFERENCES dersler(id),
                    ders_saat_id INTEGER NOT NULL REFERENCES ders_saatleri(id),
                    tarih TEXT NOT NULL,
                    giris_zamani TEXT,
                    durum TEXT NOT NULL,
                    UNIQUE(ders_saat_id, tarih)
                );
                CREATE TABLE IF NOT EXISTS yoklama_kayitlari (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    oturum_id INTEGER NOT NULL REFERENCES oturumlar(id),
                    kursiyer_id INTEGER NOT NULL REFERENCES kursiyerler(id),
                    durum TEXT NOT NULL,
                    mazeret_id INTEGER REFERENCES mazeretler(id),
                    UNIQUE(oturum_id, kursiyer_id)
                );
                CREATE TABLE IF NOT EXISTS mazeretler (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    metin TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS resmi_tatiller (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tarih TEXT UNIQUE NOT NULL,
                    ad TEXT
                );
                CREATE TABLE IF NOT EXISTS ders_ara_tatiller (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ders_id INTEGER NOT NULL REFERENCES dersler(id),
                    baslangic_tarihi TEXT NOT NULL,
                    bitis_tarihi TEXT NOT NULL
                );
                """
            )
            # Daha önceden (bu kolonlar eklenmeden önce) oluşturulmuş bir
            # veritabanı olabilir; eksik kolonları burada sonradan ekliyoruz.
            self._kolon_ekle_eger_yoksa(c, "dersler", "baslangic_tarihi", "TEXT")
            self._kolon_ekle_eger_yoksa(c, "dersler", "bitis_tarihi", "TEXT")
            self._kolon_ekle_eger_yoksa(c, "dersler", "devamsizlik_suresi_dk", "INTEGER")
            self._kolon_ekle_eger_yoksa(c, "yoklama_kayitlari", "mazeret_id", "INTEGER")
            c.execute(
                "UPDATE dersler SET devamsizlik_suresi_dk=? WHERE devamsizlik_suresi_dk IS NULL",
                (DEVAMSIZLIK_SURESI_DK,),
            )

            # grupları sadece daha önce oluşturulmadıysa ekle
            mevcut_grup = c.execute("SELECT COUNT(*) FROM gruplar").fetchone()[0]
            if mevcut_grup == 0:
                for grup_ad in GRUP_TANIMLARI:
                    c.execute("INSERT INTO gruplar (ad, sifre_hash) VALUES (?, NULL)", (grup_ad,))
            else:
                # Daha önce eski isimle oluşturulmuş bir veritabanı varsa,
                # grup adlarını yeni isme otomatik taşı.
                for eski_ad, yeni_ad in GRUP_ADI_GOCLERI.items():
                    c.execute("UPDATE gruplar SET ad=? WHERE ad=?", (yeni_ad, eski_ad))

            # Derslikler artık gruptan bağımsız, kurumun TAMAMINA ait ORTAK
            # bir havuzdur; sadece daha önce hiç oluşturulmadıysa "Derslik 1"
            # .. "Derslik {DERSLIK_SAYISI}" olarak otomatik doldurulur.
            mevcut_derslik = c.execute("SELECT COUNT(*) FROM derslikler").fetchone()[0]
            if mevcut_derslik == 0:
                for i in range(1, DERSLIK_SAYISI + 1):
                    c.execute("INSERT INTO derslikler (ad) VALUES (?)", (f"Derslik {i}",))
                for ek_ad in EK_DERSLIKLER:
                    c.execute("INSERT INTO derslikler (ad) VALUES (?)", (ek_ad,))
            else:
                # Daha önce kurulmuş bir veritabanında seminer salonları
                # (EK_DERSLIKLER) henüz eklenmemiş olabilir; eksik olanları
                # var olan derslikleri etkilemeden ekler.
                for ek_ad in EK_DERSLIKLER:
                    var_mi = c.execute(
                        "SELECT COUNT(*) FROM derslikler WHERE ad=?", (ek_ad,)
                    ).fetchone()[0]
                    if var_mi == 0:
                        c.execute("INSERT INTO derslikler (ad) VALUES (?)", (ek_ad,))
            conn.commit()
        finally:
            conn.close()

    @staticmethod
    def _kolon_ekle_eger_yoksa(cursor, tablo, kolon, tip_tanimi):
        """SQLite'ta ALTER TABLE ... ADD COLUMN IF NOT EXISTS olmadığı için,
        önce PRAGMA table_info ile kolonun var olup olmadığına bakılır."""
        mevcut_kolonlar = [r[1] for r in cursor.execute(f"PRAGMA table_info({tablo})").fetchall()]
        if kolon not in mevcut_kolonlar:
            cursor.execute(f"ALTER TABLE {tablo} ADD COLUMN {kolon} {tip_tanimi}")

    # ---- ayarlar (admin şifresi vb.) ----
    def ayar_getir(self, anahtar, varsayilan=None):
        sonuc = self.calistir("SELECT deger FROM ayarlar WHERE anahtar=?", (anahtar,))
        if sonuc:
            return sonuc[0]["deger"]
        return varsayilan

    def ayar_kaydet(self, anahtar, deger):
        self.calistir(
            "INSERT INTO ayarlar (anahtar, deger) VALUES (?, ?) "
            "ON CONFLICT(anahtar) DO UPDATE SET deger=excluded.deger",
            (anahtar, deger),
            commit=True,
        )

    def kurulum_tamam_mi(self) -> bool:
        return self.ayar_getir("kurulum_tamam") == "1"

    # ---- gruplar / derslikler ----
    def gruplari_getir(self):
        return self.calistir("SELECT * FROM gruplar ORDER BY id")

    def grup_sifre_ayarla(self, grup_id, sifre):
        self.calistir(
            "UPDATE gruplar SET sifre_hash=? WHERE id=?",
            (sifre_hashle(sifre), grup_id),
            commit=True,
        )

    def grup_ekle(self, ad):
        """Yeni bir grup ekler. Grup adı aynı olan başka bir grup varsa
        sqlite3.IntegrityError fırlatır (ad UNIQUE); çağıran taraf bunu
        yakalayıp kullanıcıya uygun bir uyarı göstermelidir."""
        self.calistir(
            "INSERT INTO gruplar (ad, sifre_hash) VALUES (?, NULL)", (ad,), commit=True
        )

    def grup_guncelle(self, grup_id, yeni_ad):
        """Var olan bir grubun adını değiştirir. Yeni ad başka bir grupta
        zaten kullanılıyorsa sqlite3.IntegrityError fırlatır."""
        self.calistir(
            "UPDATE gruplar SET ad=? WHERE id=?", (yeni_ad, grup_id), commit=True
        )

    def grup_ders_sayisi(self, grup_id):
        """Bu gruba tanımlı toplam kurs sayısı; grup silinmeden önce sıfır
        olması gerekir (yoklama geçmişini korumak için)."""
        sonuc = self.calistir(
            "SELECT COUNT(*) AS n FROM dersler WHERE grup_id=?", (grup_id,)
        )
        return sonuc[0]["n"]

    def grup_sil(self, grup_id):
        """Grubu siler. Çağıran taraf önce grup_ders_sayisi() ile bu grupta
        kurs kalmadığından emin olmalıdır; grupta hâlâ kurs varsa o kursları
        silmeden grup silinmemelidir (kursiyer/yoklama geçmişi kaybolmasın
        diye). Derslikler artık gruptan bağımsız ORTAK bir havuz olduğu için
        grup silinirken derslikler etkilenmez."""
        self.calistir("DELETE FROM gruplar WHERE id=?", (grup_id,), commit=True)

    # ---- derslikler (kurumun tamamına ait, gruplar arası ORTAK havuz) ----
    def derslikleri_getir(self):
        return self.calistir("SELECT * FROM derslikler ORDER BY id")

    def musait_derslikleri_getir(self, haric_ders_id=None):
        """Kurs Ekle / Kurs Düzenle ekranlarındaki derslik seçim listesi
        için: başka bir kursa "dolu" olarak ayrılmamış derslikleri
        döndürür. Bir derslik, üzerinde tanımlı bir kurs varsa VE o
        kursun bitiş tarihi ya boşsa (sınırsız) ya da bugün ya da
        bugünden SONRAysa "dolu" sayılır; kurs henüz başlamamış olsa
        bile derslik baştan o kursa ayrılmış kabul edilir (yani sadece
        "bugün aktif mi" değil, "bitiş tarihine kadar hâlâ ona ayrılı
        mı" bakılır). Bir kurs bitiş tarihini geçtiyse derslik otomatik
        olarak tekrar müsait hale gelir.

        haric_ders_id verilirse (var olan bir kurs düzenlenirken), o
        kursun kendisi doluluk hesabına dahil edilmez - böylece kurs
        kendi mevcut dersliğini de seçenekler arasında görmeye devam
        eder."""
        bugun_str = date.today().strftime("%Y-%m-%d")
        sorgu = (
            "SELECT * FROM derslikler WHERE id NOT IN ("
            "  SELECT derslik_id FROM dersler"
            "  WHERE (bitis_tarihi IS NULL OR bitis_tarihi = '' OR bitis_tarihi >= ?)"
        )
        params = [bugun_str]
        if haric_ders_id is not None:
            sorgu += " AND id != ?"
            params.append(haric_ders_id)
        # NOT: "ORDER BY ad" kullanılırsa alfabetik (metinsel) sıralama
        # yüzünden "Derslik 1", "Derslik 10", "Derslik 2", ... gibi kötü bir
        # sıra ortaya çıkıyordu (10, 2'den önce geliyor çünkü '1' < '2').
        # derslikleri_getir() ile aynı doğal sırayı (ekleniş/id sırası:
        # Derslik 1..40, sonra Seminer Salonu 1-2) korumak için id'ye göre
        # sıralanır.
        sorgu += ") ORDER BY id"
        return self.calistir(sorgu, tuple(params))

    def derslik_sayisi(self):
        sonuc = self.calistir("SELECT COUNT(*) AS n FROM derslikler")
        return sonuc[0]["n"]

    def derslik_ekle(self, ad):
        self.calistir(
            "INSERT INTO derslikler (ad) VALUES (?)", (ad,), commit=True
        )

    def derslik_guncelle(self, derslik_id, ad):
        self.calistir(
            "UPDATE derslikler SET ad=? WHERE id=?", (ad, derslik_id), commit=True
        )

    def derslik_kullanim_sayisi(self, derslik_id):
        sonuc = self.calistir(
            "SELECT COUNT(*) AS n FROM dersler WHERE derslik_id=?", (derslik_id,)
        )
        return sonuc[0]["n"]

    def derslik_sil(self, derslik_id):
        self.calistir("DELETE FROM derslikler WHERE id=?", (derslik_id,), commit=True)

    def grubun_aktif_derslikleri(self, grup_id, tarih_str):
        """Bu grubun, verilen tarihte aktif kursu bulunan derslikleri
        döndürür (öğretmen ekranında, o gruba ait öğretmenin sadece kendi
        grubunun ders verdiği dersliklerin listelenmesi için kullanılır)."""
        dersler = self.dersleri_getir(grup_id=grup_id)
        aktif_derslik_idler = {
            d["derslik_id"] for d in dersler if ders_tarih_araliginda_mi(d, tarih_str)
        }
        if not aktif_derslik_idler:
            return []
        yer_tutucular = ",".join("?" * len(aktif_derslik_idler))
        return self.calistir(
            f"SELECT * FROM derslikler WHERE id IN ({yer_tutucular}) ORDER BY ad",
            tuple(aktif_derslik_idler),
        )

    # ---- mazeretler ----
    def mazeretleri_getir(self):
        return self.calistir("SELECT * FROM mazeretler ORDER BY id")

    def mazeret_ekle(self, metin):
        self.calistir(
            "INSERT INTO mazeretler (metin) VALUES (?)", (metin,), commit=True
        )

    def mazeret_guncelle(self, mazeret_id, metin):
        self.calistir(
            "UPDATE mazeretler SET metin=? WHERE id=?", (metin, mazeret_id), commit=True
        )

    def mazeret_kullanim_sayisi(self, mazeret_id):
        sonuc = self.calistir(
            "SELECT COUNT(*) AS n FROM yoklama_kayitlari WHERE mazeret_id=?", (mazeret_id,)
        )
        return sonuc[0]["n"]

    def mazeret_sil(self, mazeret_id):
        """Mazereti siler; bu mazerete atıfta bulunan geçmiş yoklama
        kayıtlarını (geçmişi bozmamak için silmeden) mazeretsiz bırakır."""
        conn = self.baglan()
        try:
            conn.execute(
                "UPDATE yoklama_kayitlari SET mazeret_id=NULL WHERE mazeret_id=?", (mazeret_id,)
            )
            conn.execute("DELETE FROM mazeretler WHERE id=?", (mazeret_id,))
            conn.commit()
        finally:
            conn.close()

    # ---- resmi tatiller ----
    # Yönetici tarafından eklenen, KURUMUN TAMAMINI etkileyen (tüm
    # kurslarda, tüm dersliklerde) tatil günleri. Bu günlerde hiçbir
    # kursun hiçbir periyodu için ders yapılmaz (bkz. ders_gununde_ders_var_mi).
    def resmi_tatilleri_getir(self):
        return self.calistir("SELECT * FROM resmi_tatiller ORDER BY tarih")

    def resmi_tatil_ekle(self, tarih, ad=""):
        self.calistir(
            "INSERT OR IGNORE INTO resmi_tatiller (tarih, ad) VALUES (?, ?)",
            (tarih, ad.strip() or None), commit=True,
        )

    def resmi_tatil_sil(self, tatil_id):
        self.calistir("DELETE FROM resmi_tatiller WHERE id=?", (tatil_id,), commit=True)

    def tarih_resmi_tatil_mi(self, tarih_str):
        sonuc = self.calistir("SELECT 1 FROM resmi_tatiller WHERE tarih=?", (tarih_str,))
        return bool(sonuc)

    # ---- kurs ara tatilleri ----
    # Bir kursun KENDİ tarih aralığı(başlangıç/bitiş) İÇİNDE, o kursa özel
    # olarak ders yapılmayacak (ör. o grubun/dönemin ara tatili) tarih
    # aralıkları. Girilmesi zorunlu değildir; girilirse o aralıktaki
    # günlerde SADECE bu kurs için ders yapılmaz (diğer kurslar etkilenmez).
    def ders_ara_tatillerini_getir(self, ders_id):
        return self.calistir(
            "SELECT * FROM ders_ara_tatiller WHERE ders_id=? ORDER BY baslangic_tarihi",
            (ders_id,),
        )

    def ders_ara_tatil_ekle(self, ders_id, baslangic_tarihi, bitis_tarihi):
        self.calistir(
            "INSERT INTO ders_ara_tatiller (ders_id, baslangic_tarihi, bitis_tarihi) "
            "VALUES (?, ?, ?)",
            (ders_id, baslangic_tarihi, bitis_tarihi), commit=True,
        )

    def ders_ara_tatil_sil(self, ara_tatil_id):
        self.calistir(
            "DELETE FROM ders_ara_tatiller WHERE id=?", (ara_tatil_id,), commit=True
        )

    def ders_ara_tatil_getir(self, ara_tatil_id):
        sonuc = self.calistir("SELECT * FROM ders_ara_tatiller WHERE id=?", (ara_tatil_id,))
        return sonuc[0] if sonuc else None

    def ders_ara_tatil_guncelle(self, ara_tatil_id, baslangic_tarihi, bitis_tarihi):
        self.calistir(
            "UPDATE ders_ara_tatiller SET baslangic_tarihi=?, bitis_tarihi=? WHERE id=?",
            (baslangic_tarihi, bitis_tarihi, ara_tatil_id), commit=True,
        )

    def tum_ara_tatilleri_getir(self):
        """Tüm kurslara ait TÜM ara tatilleri (kurs/grup/derslik bilgisiyle
        birlikte) döndürür - "Ara Tatil Yönetimi" ekranındaki toplu listede
        kullanılır (bkz. AraTatilYonetimPaneli)."""
        return self.calistir(
            "SELECT t.*, d.ad AS ders_adi, g.ad AS grup_ad, dl.ad AS derslik_ad "
            "FROM ders_ara_tatiller t "
            "JOIN dersler d ON d.id=t.ders_id "
            "JOIN gruplar g ON g.id=d.grup_id "
            "JOIN derslikler dl ON dl.id=d.derslik_id "
            "ORDER BY d.ad, t.baslangic_tarihi"
        )

    def tarih_ders_ara_tatilinde_mi(self, ders_id, tarih_str):
        sonuc = self.calistir(
            "SELECT 1 FROM ders_ara_tatiller WHERE ders_id=? "
            "AND baslangic_tarihi<=? AND bitis_tarihi>=?",
            (ders_id, tarih_str, tarih_str),
        )
        return bool(sonuc)

    def tarih_tatil_mi(self, tarih_str, ders_id=None):
        """Verilen tarih ya genel bir RESMİ TATİL ise ya da (ders_id
        verildiyse) o kursun kendi ARA TATİL aralığına denk geliyorsa
        True döner - her iki durumda da o gün ders YAPILMAZ."""
        if self.tarih_resmi_tatil_mi(tarih_str):
            return True
        if ders_id is not None and self.tarih_ders_ara_tatilinde_mi(ders_id, tarih_str):
            return True
        return False

    # ---- dersler (kurslar) ----
    # Bir kurs artık tek bir gün/saate değil, haftalık bir PROGRAMA
    # sahiptir: Pazartesi-Perşembe her gün aynı 6 periyot (saat), Cuma
    # ayrı bir 6 periyotluk saat listesiyle. Programın kendisi
    # "ders_saatleri" tablosunda (gun, sira, baslangic, bitis) olarak
    # tutulur; "dersler" tablosu artık sadece kursun kendisini (grup,
    # derslik, ad, tarih aralığı, yoklama süresi) tutar.
    def ders_ekle(self, grup_id, derslik_id, ad, saatler,
                  baslangic_tarihi="", bitis_tarihi="", devamsizlik_suresi_dk=None):
        """saatler: [(gun, sira, baslangic, bitis), ...] - genelde 30 satır
        (Pzt/Sal/Çar/Per için 6'şar aynı saatler + Cuma için 6 saat)."""
        conn = self.baglan()
        try:
            cur = conn.execute(
                "INSERT INTO dersler (grup_id, derslik_id, ad, baslangic_tarihi, "
                "bitis_tarihi, devamsizlik_suresi_dk) VALUES (?, ?, ?, ?, ?, ?)",
                (
                    grup_id, derslik_id, ad,
                    baslangic_tarihi or None, bitis_tarihi or None,
                    devamsizlik_suresi_dk or DEVAMSIZLIK_SURESI_DK,
                ),
            )
            ders_id = cur.lastrowid
            for gun, sira, baslangic, bitis in saatler:
                conn.execute(
                    "INSERT INTO ders_saatleri (ders_id, gun, sira, baslangic, bitis) "
                    "VALUES (?, ?, ?, ?, ?)",
                    (ders_id, gun, sira, baslangic, bitis),
                )
            conn.commit()
            return ders_id
        finally:
            conn.close()

    def dersleri_getir(self, grup_id=None, derslik_id=None):
        sorgu = (
            "SELECT d.*, g.ad AS grup_ad, dl.ad AS derslik_ad FROM dersler d "
            "JOIN gruplar g ON g.id=d.grup_id "
            "JOIN derslikler dl ON dl.id=d.derslik_id WHERE 1=1"
        )
        params = []
        if grup_id is not None:
            sorgu += " AND d.grup_id=?"
            params.append(grup_id)
        if derslik_id is not None:
            sorgu += " AND d.derslik_id=?"
            params.append(derslik_id)
        sorgu += " ORDER BY d.ad"
        return self.calistir(sorgu, tuple(params))

    def ders_getir(self, ders_id):
        sonuc = self.calistir(
            "SELECT d.*, g.ad AS grup_ad, dl.ad AS derslik_ad FROM dersler d "
            "JOIN gruplar g ON g.id=d.grup_id "
            "JOIN derslikler dl ON dl.id=d.derslik_id WHERE d.id=?",
            (ders_id,),
        )
        return sonuc[0] if sonuc else None

    def derslik_aktif_dersleri(self, derslik_id, tarih_str):
        """Bu derslikte, verilen tarihte aktif olan kurs(lar)ı döndürür
        (normalde 0 ya da 1 tane olur; bir kurs bitip yenisi başladıysa
        tarih aralıkları çakışmadığı sürece ikisi aynı anda listeye
        girmez). Resmi tatil ya da kursun kendi ara tatili olan günlerde
        kurs burada AKTİF SAYILMAZ (bkz. ders_gununde_ders_var_mi) -
        böylece öğretmenin ekranında o günün sekmesinde hiç periyot
        görünmez (o gün ders yokmuş gibi davranılır)."""
        tum_dersler = self.dersleri_getir(derslik_id=derslik_id)
        return [d for d in tum_dersler if ders_gununde_ders_var_mi(d, tarih_str, self)]

    # ---- ders saatleri (haftalık program) ----
    def ders_saatleri_getir(self, ders_id):
        return self.calistir(
            "SELECT * FROM ders_saatleri WHERE ders_id=? ORDER BY gun, sira", (ders_id,)
        )

    def ders_saatleri_guncelle(self, ders_id, saatler):
        """Bir kursun haftalık programını günceller (Kurs Düzenle ekranında
        kullanılır). Var olan (gün, sıra) periyotlarının saatleri YERİNDE
        güncellenir (id'leri değişmez), böylece o periyoda daha önce
        alınmış yoklama oturumları (oturumlar.ders_saat_id) bozulmaz.
        Yeni programda artık bulunmayan periyotlar, varsa önce onlara ait
        oturum/yoklama kayıtlarıyla birlikte silinir."""
        conn = self.baglan()
        try:
            mevcutlar = {
                (r["gun"], r["sira"]): r["id"]
                for r in conn.execute(
                    "SELECT id, gun, sira FROM ders_saatleri WHERE ders_id=?", (ders_id,)
                ).fetchall()
            }
            yeni_anahtarlar = set()
            for gun, sira, baslangic, bitis in saatler:
                yeni_anahtarlar.add((gun, sira))
                if (gun, sira) in mevcutlar:
                    conn.execute(
                        "UPDATE ders_saatleri SET baslangic=?, bitis=? WHERE id=?",
                        (baslangic, bitis, mevcutlar[(gun, sira)]),
                    )
                else:
                    conn.execute(
                        "INSERT INTO ders_saatleri (ders_id, gun, sira, baslangic, bitis) "
                        "VALUES (?, ?, ?, ?, ?)",
                        (ders_id, gun, sira, baslangic, bitis),
                    )
            for anahtar, ders_saat_id in mevcutlar.items():
                if anahtar in yeni_anahtarlar:
                    continue
                oturum_idler = [
                    r["id"] for r in conn.execute(
                        "SELECT id FROM oturumlar WHERE ders_saat_id=?", (ders_saat_id,)
                    ).fetchall()
                ]
                for oturum_id in oturum_idler:
                    conn.execute("DELETE FROM yoklama_kayitlari WHERE oturum_id=?", (oturum_id,))
                conn.execute("DELETE FROM oturumlar WHERE ders_saat_id=?", (ders_saat_id,))
                conn.execute("DELETE FROM ders_saatleri WHERE id=?", (ders_saat_id,))
            conn.commit()
        finally:
            conn.close()

    # ---- kursiyerler ----
    def kursiyer_ekle(self, ders_id, kod, ad_soyad):
        self.calistir(
            "INSERT INTO kursiyerler (ders_id, kod, ad_soyad) VALUES (?, ?, ?)",
            (ders_id, kod, ad_soyad),
            commit=True,
        )

    def kursiyerleri_getir(self, ders_id):
        return self.calistir(
            "SELECT * FROM kursiyerler WHERE ders_id=? ORDER BY kod", (ders_id,)
        )

    def kod_kullanilmis_mi(self, ders_id, kod, haric_kursiyer_id=None):
        if haric_kursiyer_id is None:
            sonuc = self.calistir(
                "SELECT COUNT(*) AS n FROM kursiyerler WHERE ders_id=? AND kod=?",
                (ders_id, kod),
            )
        else:
            sonuc = self.calistir(
                "SELECT COUNT(*) AS n FROM kursiyerler WHERE ders_id=? AND kod=? AND id<>?",
                (ders_id, kod, haric_kursiyer_id),
            )
        return sonuc[0]["n"] > 0

    def kursiyer_getir(self, kursiyer_id):
        sonuc = self.calistir("SELECT * FROM kursiyerler WHERE id=?", (kursiyer_id,))
        return sonuc[0] if sonuc else None

    def kursiyer_guncelle(self, kursiyer_id, kod, ad_soyad):
        self.calistir(
            "UPDATE kursiyerler SET kod=?, ad_soyad=? WHERE id=?",
            (kod, ad_soyad, kursiyer_id),
            commit=True,
        )

    def kursiyer_sil(self, kursiyer_id):
        """Kursiyeri ve ona ait tüm yoklama kayıtlarını siler."""
        conn = self.baglan()
        try:
            conn.execute("DELETE FROM yoklama_kayitlari WHERE kursiyer_id=?", (kursiyer_id,))
            conn.execute("DELETE FROM kursiyerler WHERE id=?", (kursiyer_id,))
            conn.commit()
        finally:
            conn.close()

    def ders_guncelle(self, ders_id, ad, saatler, derslik_id=None,
                       baslangic_tarihi="", bitis_tarihi="", devamsizlik_suresi_dk=None):
        if derslik_id is not None:
            self.calistir(
                "UPDATE dersler SET ad=?, derslik_id=?, baslangic_tarihi=?, bitis_tarihi=?, "
                "devamsizlik_suresi_dk=? WHERE id=?",
                (
                    ad, derslik_id, baslangic_tarihi or None, bitis_tarihi or None,
                    devamsizlik_suresi_dk or DEVAMSIZLIK_SURESI_DK,
                    ders_id,
                ),
                commit=True,
            )
        else:
            self.calistir(
                "UPDATE dersler SET ad=?, baslangic_tarihi=?, bitis_tarihi=?, "
                "devamsizlik_suresi_dk=? WHERE id=?",
                (
                    ad, baslangic_tarihi or None, bitis_tarihi or None,
                    devamsizlik_suresi_dk or DEVAMSIZLIK_SURESI_DK,
                    ders_id,
                ),
                commit=True,
            )
        self.ders_saatleri_guncelle(ders_id, saatler)
        self._bugunku_acik_oturumlari_guncelle(ders_id)

    def _bugunku_acik_oturumlari_guncelle(self, ders_id):
        """Bir kursun yoklama girme süresi (devamsizlik_suresi_dk) SONRADAN
        ARTIRILDIĞINDA: sistem daha önce (eski, kısa süreyle) otomatik
        olarak 'öğretmen girmedi' (devamsız) diye işaretlemiş olabileceği
        BUGÜNKÜ periyotları, eğer şu an YENİ süreye göre hâlâ yoklama
        girme penceresi içindeyse, tekrar 'bekliyor' durumuna döndürür -
        böylece öğretmen o periyot için yoklamayı sayfasında görüp
        girebilir. Öğretmenin GERÇEKTEN elle tamamladığı (durum=
        'tamamlandi') oturumlara asla dokunulmaz."""
        bugun_str = date.today().strftime("%Y-%m-%d")
        ders = self.ders_getir(ders_id)
        if not ders or not ders_gununde_ders_var_mi(ders, bugun_str, self):
            return  # kurs bugün zaten aktif değilse (resmi/ara tatil dahil) dokunmaya gerek yok
        sure_dk = ders["devamsizlik_suresi_dk"] or DEVAMSIZLIK_SURESI_DK
        simdi = datetime.now()
        conn = self.baglan()
        try:
            saatler_bugun = conn.execute(
                "SELECT id, baslangic FROM ders_saatleri WHERE ders_id=? AND gun=?",
                (ders_id, bugun_gun_index()),
            ).fetchall()
            for s in saatler_bugun:
                try:
                    bas_saat = saat_str_to_time(s["baslangic"])
                except ValueError:
                    continue
                bas_dt = datetime.combine(date.today(), bas_saat)
                son_giris_dt = bas_dt + timedelta(minutes=sure_dk)
                if simdi > son_giris_dt:
                    continue  # yeni süreyle de zaten dolmuş, dokunma
                oturum = conn.execute(
                    "SELECT * FROM oturumlar WHERE ders_saat_id=? AND tarih=?",
                    (s["id"], bugun_str),
                ).fetchone()
                if oturum and oturum["durum"] == "ogretmen_girmedi":
                    conn.execute(
                        "DELETE FROM yoklama_kayitlari WHERE oturum_id=?", (oturum["id"],)
                    )
                    conn.execute("DELETE FROM oturumlar WHERE id=?", (oturum["id"],))
            conn.commit()
        finally:
            conn.close()

    def ders_kursiyer_sayisi(self, ders_id):
        sonuc = self.calistir(
            "SELECT COUNT(*) AS n FROM kursiyerler WHERE ders_id=?", (ders_id,)
        )
        return sonuc[0]["n"]

    def ders_sil(self, ders_id):
        """Dersi; ona ait tüm ders saatlerini (haftalık programı),
        kursiyerleri, oturumları, yoklama kayıtlarını ve ara tatillerini
        kalıcı olarak siler."""
        conn = self.baglan()
        try:
            oturum_idler = [
                r["id"] for r in conn.execute(
                    "SELECT id FROM oturumlar WHERE ders_id=?", (ders_id,)
                ).fetchall()
            ]
            for oturum_id in oturum_idler:
                conn.execute("DELETE FROM yoklama_kayitlari WHERE oturum_id=?", (oturum_id,))
            conn.execute("DELETE FROM oturumlar WHERE ders_id=?", (ders_id,))
            conn.execute("DELETE FROM kursiyerler WHERE ders_id=?", (ders_id,))
            conn.execute("DELETE FROM ders_saatleri WHERE ders_id=?", (ders_id,))
            conn.execute("DELETE FROM ders_ara_tatiller WHERE ders_id=?", (ders_id,))
            conn.execute("DELETE FROM dersler WHERE id=?", (ders_id,))
            conn.commit()
        finally:
            conn.close()

    # ---- oturumlar / yoklama ----
    # Bir kursun her periyodu (ders_saat) için ayrı bir "oturum" (ve dolayısıyla
    # ayrı bir yoklama) oluşur; öğretmen aynı kursta bir günde 6 kez yoklama
    # alabilir (her periyot için bir kez).
    def oturum_getir(self, ders_saat_id, tarih):
        sonuc = self.calistir(
            "SELECT * FROM oturumlar WHERE ders_saat_id=? AND tarih=?", (ders_saat_id, tarih)
        )
        return sonuc[0] if sonuc else None

    def oturum_olustur_veya_getir(self, ders_id, ders_saat_id, tarih):
        mevcut = self.oturum_getir(ders_saat_id, tarih)
        if mevcut:
            return mevcut
        self.calistir(
            "INSERT OR IGNORE INTO oturumlar (ders_id, ders_saat_id, tarih, giris_zamani, durum) "
            "VALUES (?, ?, ?, NULL, 'bekliyor')",
            (ders_id, ders_saat_id, tarih),
            commit=True,
        )
        return self.oturum_getir(ders_saat_id, tarih)

    def yoklama_kaydet(self, oturum_id, kayitlar):
        """kayitlar: [(kursiyer_id, 'var'/'yok'), ...] veya
        [(kursiyer_id, 'var'/'yok', mazeret_id_veya_None), ...]. Öğretmen
        aynı oturum için istediği kadar tekrar kaydedebilir (var olan
        kayıtlar güncellenir), her kayıtta 'giriş zamanı' o anki zamana
        güncellenir (en son giriş zamanı böylece Excel'de görünür)."""
        conn = self.baglan()
        try:
            for kayit in kayitlar:
                kursiyer_id, durum = kayit[0], kayit[1]
                mazeret_id = kayit[2] if len(kayit) > 2 else None
                conn.execute(
                    "INSERT INTO yoklama_kayitlari (oturum_id, kursiyer_id, durum, mazeret_id) "
                    "VALUES (?, ?, ?, ?) "
                    "ON CONFLICT(oturum_id, kursiyer_id) "
                    "DO UPDATE SET durum=excluded.durum, mazeret_id=excluded.mazeret_id",
                    (oturum_id, kursiyer_id, durum, mazeret_id),
                )
            conn.execute(
                "UPDATE oturumlar SET durum='tamamlandi', giris_zamani=? WHERE id=?",
                (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), oturum_id),
            )
            conn.commit()
        finally:
            conn.close()

    def yoklama_kayitlarini_getir(self, oturum_id):
        return self.calistir(
            "SELECT y.*, k.kod, k.ad_soyad, m.metin AS mazeret_metin FROM yoklama_kayitlari y "
            "JOIN kursiyerler k ON k.id=y.kursiyer_id "
            "LEFT JOIN mazeretler m ON m.id=y.mazeret_id "
            "WHERE y.oturum_id=? ORDER BY k.kod",
            (oturum_id,),
        )

    def _bugunun_ders_saatleri(self, tarih_str):
        """Verilen tarihte (gün indeksine göre) aktif olan tüm kursların o
        günkü periyotlarını (ders_saatleri satırları + kurs/derslik/grup
        bilgileriyle birlikte) döndürür. Resmi tatil ya da kursun kendi
        ara tatiline denk gelen günlerde o kursun periyotları hiç
        döndürülmez - böylece yönetici panelinde/Excel'de o gün için
        hayalet bir satır görünmez ve otomatik devamsızlık kontrolü de
        o periyotlara dokunmaz."""
        gun_idx = datetime.strptime(tarih_str, "%Y-%m-%d").weekday()
        satirlar = self.calistir(
            "SELECT ds.id AS ders_saat_id, ds.gun, ds.sira, ds.baslangic, ds.bitis, "
            "d.id AS ders_id, d.ad AS ders_adi, d.baslangic_tarihi, d.bitis_tarihi, "
            "d.devamsizlik_suresi_dk, g.ad AS grup_ad, dl.ad AS derslik_ad FROM ders_saatleri ds "
            "JOIN dersler d ON d.id=ds.ders_id "
            "JOIN gruplar g ON g.id=d.grup_id "
            "JOIN derslikler dl ON dl.id=d.derslik_id "
            "WHERE ds.gun=? ORDER BY ds.sira, dl.ad",
            (gun_idx,),
        )
        return [s for s in satirlar if ders_gununde_ders_var_mi(s, tarih_str, self)]

    def otomatik_devamsizlik_kontrolu(self):
        """Başlangıcından itibaren kursun kendi süresi (devamsizlik_suresi_dk)
        kadar zaman geçmiş ve henüz oturum açılmamış / tamamlanmamış bugünkü
        periyotları (ders_saatleri) bulup otomatik olarak 'ogretmen_girmedi'
        + tüm kursiyerleri 'yok' yapar."""
        simdi = datetime.now()
        bugun_str = date.today().strftime("%Y-%m-%d")
        for saat in self._bugunun_ders_saatleri(bugun_str):
            try:
                bas_saat = saat_str_to_time(saat["baslangic"])
            except ValueError:
                continue
            sure_dk = saat["devamsizlik_suresi_dk"] or DEVAMSIZLIK_SURESI_DK
            bas_dt = datetime.combine(date.today(), bas_saat)
            son_giris_dt = bas_dt + timedelta(minutes=sure_dk)
            if simdi <= son_giris_dt:
                continue  # süre daha dolmadı
            mevcut_oturum = self.oturum_getir(saat["ders_saat_id"], bugun_str)
            if mevcut_oturum and mevcut_oturum["durum"] in ("tamamlandi", "ogretmen_girmedi"):
                continue  # zaten işlenmiş
            oturum = self.oturum_olustur_veya_getir(saat["ders_id"], saat["ders_saat_id"], bugun_str)
            conn = self.baglan()
            try:
                conn.execute(
                    "UPDATE oturumlar SET durum='ogretmen_girmedi' WHERE id=?",
                    (oturum["id"],),
                )
                kursiyerler = conn.execute(
                    "SELECT id FROM kursiyerler WHERE ders_id=?", (saat["ders_id"],)
                ).fetchall()
                for k in kursiyerler:
                    conn.execute(
                        "INSERT INTO yoklama_kayitlari (oturum_id, kursiyer_id, durum) "
                        "VALUES (?, ?, 'yok') "
                        "ON CONFLICT(oturum_id, kursiyer_id) DO UPDATE SET durum='yok'",
                        (oturum["id"], k["id"]),
                    )
                conn.commit()
            finally:
                conn.close()

    def gunluk_ozet(self, tarih_str):
        """Yönetici paneli için: o güne ait tüm kurs periyotları + durumları
        + toplam/yok sayısı (her periyot ayrı bir satır - bir kurs günde
        6 kez listelenebilir)."""
        satirlar = []
        for saat in self._bugunun_ders_saatleri(tarih_str):
            toplam = self.calistir(
                "SELECT COUNT(*) AS n FROM kursiyerler WHERE ders_id=?", (saat["ders_id"],)
            )[0]["n"]
            oturum = self.oturum_getir(saat["ders_saat_id"], tarih_str)
            if oturum:
                yok_sayisi = self.calistir(
                    "SELECT COUNT(*) AS n FROM yoklama_kayitlari "
                    "WHERE oturum_id=? AND durum='yok'",
                    (oturum["id"],),
                )[0]["n"]
                durum = oturum["durum"]
                giris_zamani = oturum["giris_zamani"] or ""
                var_sayisi = toplam - yok_sayisi
            else:
                yok_sayisi = 0
                durum = "bekliyor"
                giris_zamani = ""
                var_sayisi = 0  # yoklama henüz alınmadı
            satirlar.append(
                {
                    "ders_id": saat["ders_id"],
                    "ders_saat_id": saat["ders_saat_id"],
                    "oturum_id": oturum["id"] if oturum else None,
                    "tarih": tarih_str,
                    "grup": saat["grup_ad"],
                    "derslik": saat["derslik_ad"],
                    "ders_adi": saat["ders_adi"],
                    "sira": saat["sira"],
                    "saat": f"{saat['sira']}. Saat  {saat['baslangic']} - {saat['bitis']}",
                    "durum": durum,
                    "giris_zamani": giris_zamani,
                    "kurs_baslangic": iso_tarihi_goruntu_metnine_cevir(saat["baslangic_tarihi"]),
                    "kurs_bitis": iso_tarihi_goruntu_metnine_cevir(saat["bitis_tarihi"]),
                    "toplam": toplam,
                    "var": var_sayisi,
                    "yok": yok_sayisi,
                }
            )
        return satirlar


# --------------------------------------------------------------------------
# Basit renk / stil sabitleri
# --------------------------------------------------------------------------

RENK_ARKAPLAN = "#f4f5f7"
RENK_SOL_PANEL = "#1f2937"
RENK_SOL_BUTON = "#374151"
RENK_VURGU = "#2563eb"
RENK_YESIL = "#16a34a"
RENK_KIRMIZI = "#dc2626"
RENK_METIN_ACIK = "#f9fafb"

DURUM_METIN = {
    "bekliyor": "Bekliyor",
    "tamamlandi": "Yoklama alındı",
    "ogretmen_girmedi": "Öğretmen girmedi (devamsız)",
}


def _baslik_cubugu_rengini_ayarla(pencere, renk_hex=RENK_SOL_PANEL):
    """Pencerenin Windows tarafından çizilen başlık çubuğunu (simge/kapat
    düğmesinin olduğu en üst şerit) uygulamanın kendi rengiyle boyar.
    Bu özellik yalnızca Windows 11'de (derleme 22000+) desteklenir; başka
    bir işletim sisteminde veya eski bir Windows sürümünde sessizce hiçbir
    şey yapmaz, uygulama normal (renklendirilmemiş başlık çubuğuyla)
    çalışmaya devam eder."""
    if sys.platform != "win32":
        return
    try:
        pencere.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(pencere.winfo_id())
        renk = renk_hex.lstrip("#")
        r = int(renk[0:2], 16)
        g = int(renk[2:4], 16)
        b = int(renk[4:6], 16)
        colorref = r | (g << 8) | (b << 16)  # Windows COLORREF: 0x00BBGGRR
        DWMWA_CAPTION_COLOR = 35
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_CAPTION_COLOR,
            ctypes.byref(ctypes.c_int(colorref)), ctypes.sizeof(ctypes.c_int),
        )
    except Exception:
        pass  # desteklenmiyorsa sessizce geç, uygulama yine çalışsın


def _grab_guvenli_baglat(pencere):
    """grab_set() aktifken pencere Windows'ta 'Masaüstünü Göster' ile (ya
    da başka bir şekilde) simge durumuna küçültülürse, Tk'nin iç grab
    durumu bozulabiliyor: pencere bir daha HİÇBİR ŞEKİLDE öne
    getirilemiyor ve uygulamayı kapatmak için görev yöneticisi
    gerekiyordu. Bunun nedeni, grab'i tutan pencere görünmez (iconic)
    haldeyken bile Tk'nin tüm olayları hâlâ o pencereye yönlendirmeye
    çalışması. Çözüm: pencere küçüldüğünde (Unmap) kilidi GEÇİCİ olarak
    bırakıp, pencere tekrar görünür olduğunda (Map) kilidi yeniden
    kurmak - böylece küçültme sırasında kilit asla "görünmeyen" bir
    pencerede takılı kalmıyor.

    self.grab_set() çağrısından HEMEN SONRA çağrılmalıdır; diyalogların
    açılma/kapanma veya iç içe (nested) modal akışını etkilemez, çünkü
    bir pencerenin ÜSTÜNE başka bir pencere açılması (ör. bir alt
    diyalog ya da messagebox) onu "unmap" etmez - sadece gerçekten
    simge durumuna küçültülmesi ya da gizlenmesi tetikler."""
    def _kuculdu(event, p=pencere):
        if event.widget is not p:
            return
        try:
            if p.grab_current() == p:
                p.grab_release()
        except tk.TclError:
            pass

    def _tekrar_gorundu(event, p=pencere):
        if event.widget is not p:
            return
        try:
            if p.winfo_exists() and p.state() != "iconic":
                p.grab_set()
        except tk.TclError:
            pass

    pencere.bind("<Unmap>", _kuculdu)
    pencere.bind("<Map>", _tekrar_gorundu)


def _canvasa_fare_tekeri_ekle(canvas):
    """Kaydırılabilir bir Canvas'a fare tekerleğiyle yukarı/aşağı kaydırma
    özelliği ekler (Windows/Mac: <MouseWheel>, Linux: <Button-4>/<Button-5>).
    İmleç canvas'ın üzerine geldiğinde etkinleşir, ayrıldığında devre dışı
    kalır; böylece aynı anda birden fazla kaydırılabilir pencere açık olsa
    bile sadece imlecin üzerinde olduğu alan tekerlekten etkilenir."""
    def _tekerlek(event):
        if getattr(event, "num", None) == 4:
            canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            canvas.yview_scroll(1, "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _baglan(event=None):
        canvas.bind_all("<MouseWheel>", _tekerlek)
        canvas.bind_all("<Button-4>", _tekerlek)
        canvas.bind_all("<Button-5>", _tekerlek)

    def _birak(event=None):
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Button-4>")
        canvas.unbind_all("<Button-5>")

    canvas.bind("<Enter>", _baglan)
    canvas.bind("<Leave>", _birak)


def _agaca_fare_tekeri_ekle(widget, canvas):
    """_canvasa_fare_tekeri_ekle'ye (Enter/Leave tabanlı) EK bir güvence
    olarak, verilen widget'a ve TÜM alt bileşenlerine DOĞRUDAN fare
    tekerleği bağlar; her biri verilen canvas'ı kaydırır. Bu, Enter/Leave
    olaylarının bazı ortamlarda beklenen şekilde tetiklenmediği durumlara
    karşı ek bir garantidir. Dinamik olarak yeniden oluşturulan içerikler
    (satır ekleme/silme, listeyi yenileme) için ilgili yerlerde tekrar
    çağrılmalıdır."""
    def _tekerlek(event):
        if getattr(event, "num", None) == 4:
            canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            canvas.yview_scroll(1, "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    try:
        widget.bind("<MouseWheel>", _tekerlek, add="+")
        widget.bind("<Button-4>", _tekerlek, add="+")
        widget.bind("<Button-5>", _tekerlek, add="+")
    except tk.TclError:
        pass
    for cocuk in widget.winfo_children():
        _agaca_fare_tekeri_ekle(cocuk, canvas)


# --------------------------------------------------------------------------
# Kurulum sihirbazı (ilk çalıştırma - şifreleri belirleme)
# --------------------------------------------------------------------------

class KurulumSihirbazi(tk.Toplevel):
    def __init__(self, master, db: VeriTabani, tamamlanma_geri_cagirma):
        super().__init__(master)
        self.db = db
        self.tamamlanma_geri_cagirma = tamamlanma_geri_cagirma
        self.title("İlk Kurulum - Şifreleri Belirleyin")
        self.geometry("480x560")
        self.resizable(False, False)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)
        self.protocol("WM_DELETE_WINDOW", self._kapatilamaz)

        tk.Label(
            self,
            text="İlk Kurulum",
            font=("Segoe UI", 16, "bold"),
        ).pack(pady=(16, 4))
        tk.Label(
            self,
            text="Her grup için bir giriş şifresi ve bir de yönetici\n"
                 "(Yoklama paneli, Kurs Ekle, Kursiyer Ekle) şifresi belirleyin.",
            justify="center",
        ).pack(pady=(0, 12))

        form = tk.Frame(self)
        form.pack(padx=20, fill="x")

        self.grup_girisler = {}
        gruplar = self.db.gruplari_getir()
        for grup in gruplar:
            satir = tk.Frame(form)
            satir.pack(fill="x", pady=4)
            tk.Label(satir, text=grup["ad"], width=16, anchor="w").pack(side="left")
            girdi = tk.Entry(satir, show="*")
            girdi.pack(side="left", fill="x", expand=True)
            self.grup_girisler[grup["id"]] = girdi

        ttk.Separator(self).pack(fill="x", pady=12, padx=20)

        admin_frame = tk.Frame(self)
        admin_frame.pack(padx=20, fill="x")
        tk.Label(admin_frame, text="Yönetici şifresi", width=16, anchor="w").pack(side="left")
        self.admin_girdi = tk.Entry(admin_frame, show="*")
        self.admin_girdi.pack(side="left", fill="x", expand=True)

        izleyici_frame = tk.Frame(self)
        izleyici_frame.pack(padx=20, fill="x", pady=(6, 0))
        tk.Label(izleyici_frame, text="İzleyici şifresi", width=16, anchor="w").pack(side="left")
        self.izleyici_girdi = tk.Entry(izleyici_frame, show="*")
        self.izleyici_girdi.pack(side="left", fill="x", expand=True)

        tk.Label(
            self,
            text="(Grup ve yönetici şifreleri boş bırakılamaz. İzleyici şifresi\n"
                 "isteğe bağlıdır, boş bırakılırsa sonra Yönetici panelinden\n"
                 "-> Şifreleri Yönet ile belirlenebilir.)",
            fg="#555555",
            justify="center",
        ).pack(pady=(12, 8))

        tk.Button(
            self, text="Kaydet ve Başla", bg=RENK_VURGU, fg="white",
            font=("Segoe UI", 11, "bold"), command=self._kaydet
        ).pack(pady=10, ipadx=10, ipady=6)

    def _kapatilamaz(self):
        messagebox.showwarning(
            APP_ADI, "Uygulamayı kullanmak için önce kurulumu tamamlamalısınız.",
            parent=self,
        )

    def _kaydet(self):
        degerler = {}
        for grup_id, girdi in self.grup_girisler.items():
            deger = girdi.get().strip()
            if not deger:
                messagebox.showerror(APP_ADI, "Lütfen tüm grup şifrelerini girin.", parent=self)
                return
            degerler[grup_id] = deger
        admin_sifre = self.admin_girdi.get().strip()
        if not admin_sifre:
            messagebox.showerror(APP_ADI, "Lütfen yönetici şifresini girin.", parent=self)
            return

        for grup_id, deger in degerler.items():
            self.db.grup_sifre_ayarla(grup_id, deger)
        self.db.ayar_kaydet("admin_sifre_hash", sifre_hashle(admin_sifre))
        izleyici_sifre = self.izleyici_girdi.get().strip()
        if izleyici_sifre:
            self.db.ayar_kaydet("izleyici_sifre_hash", sifre_hashle(izleyici_sifre))
        self.db.ayar_kaydet("kurulum_tamam", "1")

        messagebox.showinfo(APP_ADI, "Kurulum tamamlandı.", parent=self)
        self.destroy()
        self.tamamlanma_geri_cagirma()


# --------------------------------------------------------------------------
# Ders Ekle penceresi
# --------------------------------------------------------------------------

VARSAYILAN_PERIYOT_SAATLERI = [
    ("08:00", "08:50"),
    ("09:00", "09:50"),
    ("10:00", "10:50"),
    ("11:00", "11:50"),
    ("13:00", "13:50"),
    ("14:00", "14:50"),
]


class HaftalikProgramCercevesi(tk.Frame):
    """Bir kursun haftalık programını girmek için tekrar kullanılabilir
    çerçeve: Pazartesi-Perşembe için ortak 6 periyot + Cuma için ayrı 6
    periyot (her periyot için başlangıç/bitiş saati, HH:MM)."""

    def __init__(self, master):
        super().__init__(master)
        self.pzt_per_girisleri = []
        self.cuma_girisleri = []

        pzt_per_cercevesi = tk.LabelFrame(self, text="Pazartesi - Salı - Çarşamba - Perşembe", padx=10, pady=8)
        pzt_per_cercevesi.pack(fill="x", pady=(0, 8))
        self._periyot_satirlari_olustur(pzt_per_cercevesi, self.pzt_per_girisleri)

        cuma_cercevesi = tk.LabelFrame(self, text="Cuma", padx=10, pady=8)
        cuma_cercevesi.pack(fill="x")
        self._periyot_satirlari_olustur(cuma_cercevesi, self.cuma_girisleri)

    def _periyot_satirlari_olustur(self, parent, giris_listesi):
        # NOT: Saat/dakika alanları önceden serbest metin (Entry) olduğu
        # için buraya "1,2,3" gibi geçersiz herhangi bir şey yazılabiliyordu.
        # Artık sadece LİSTEDEN SEÇİLEBİLEN (state="readonly", elle yazma
        # kapalı), saat ve dakikanın AYRI kutucuklarda olduğu Combobox'lar
        # kullanılıyor.
        # NOT: "(Saat)"/"(Dakika)" ekleri kaldırıldı - başlık, altındaki
        # saat+dakika kutucuğu ÇİFTİNİN üzerine ortalanarak (columnspan=2)
        # tek bir "Başlangıç"/"Bitiş" yazısı olarak gösteriliyor.
        # NOT: "Başlangıç" ve "Bitiş" grupları birbirine çok yakın durup
        # karışıklık yaratmasın diye, Bitiş grubunun (3. sütun) solunda
        # normalden daha büyük bir boşluk (padx=(20, 4)) bırakılıyor.
        tk.Label(parent, text="Periyot", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=4, pady=2)
        tk.Label(parent, text="Başlangıç", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, columnspan=2, padx=4, pady=2)
        tk.Label(parent, text="Bitiş", font=("Segoe UI", 9, "bold")).grid(row=0, column=3, columnspan=2, padx=(20, 4), pady=2)
        for i in range(6):
            varsayilan_bas, varsayilan_bit = VARSAYILAN_PERIYOT_SAATLERI[i]
            bas_saat_v, bas_dakika_v = varsayilan_bas.split(":")
            bit_saat_v, bit_dakika_v = varsayilan_bit.split(":")
            tk.Label(parent, text=f"{i + 1}. Saat").grid(row=i + 1, column=0, padx=4, pady=2, sticky="w")

            bas_saat_combo = ttk.Combobox(
                parent, values=SAAT_SECENEKLERI, width=4, justify="center", state="readonly"
            )
            bas_saat_combo.set(bas_saat_v)
            bas_saat_combo.grid(row=i + 1, column=1, padx=4, pady=2)

            bas_dakika_combo = ttk.Combobox(
                parent, values=DAKIKA_SECENEKLERI, width=4, justify="center", state="readonly"
            )
            bas_dakika_combo.set(bas_dakika_v)
            bas_dakika_combo.grid(row=i + 1, column=2, padx=4, pady=2)

            bit_saat_combo = ttk.Combobox(
                parent, values=SAAT_SECENEKLERI, width=4, justify="center", state="readonly"
            )
            bit_saat_combo.set(bit_saat_v)
            bit_saat_combo.grid(row=i + 1, column=3, padx=(20, 4), pady=2)

            bit_dakika_combo = ttk.Combobox(
                parent, values=DAKIKA_SECENEKLERI, width=4, justify="center", state="readonly"
            )
            bit_dakika_combo.set(bit_dakika_v)
            bit_dakika_combo.grid(row=i + 1, column=4, padx=4, pady=2)

            giris_listesi.append((bas_saat_combo, bas_dakika_combo, bit_saat_combo, bit_dakika_combo))

    def saatleri_yukle(self, ders_saatleri_satirlari):
        """ders_saatleri_getir(ders_id) çıktısındaki satırlarla alanları
        doldurur (mevcut bir kursu düzenlerken kullanılır)."""
        pzt_per, cuma = {}, {}
        for s in ders_saatleri_satirlari:
            if s["gun"] == 4:
                cuma[s["sira"]] = (s["baslangic"], s["bitis"])
            elif s["gun"] in (0, 1, 2, 3):
                pzt_per.setdefault(s["sira"], (s["baslangic"], s["bitis"]))
        for sira in range(1, 7):
            if sira in pzt_per:
                bas, bit = pzt_per[sira]
                bas_saat_combo, bas_dakika_combo, bit_saat_combo, bit_dakika_combo = (
                    self.pzt_per_girisleri[sira - 1]
                )
                bas_saat_v, bas_dakika_v = bas.split(":")
                bit_saat_v, bit_dakika_v = bit.split(":")
                bas_saat_combo.set(bas_saat_v)
                bas_dakika_combo.set(bas_dakika_v)
                bit_saat_combo.set(bit_saat_v)
                bit_dakika_combo.set(bit_dakika_v)
            if sira in cuma:
                bas, bit = cuma[sira]
                bas_saat_combo, bas_dakika_combo, bit_saat_combo, bit_dakika_combo = (
                    self.cuma_girisleri[sira - 1]
                )
                bas_saat_v, bas_dakika_v = bas.split(":")
                bit_saat_v, bit_dakika_v = bit.split(":")
                bas_saat_combo.set(bas_saat_v)
                bas_dakika_combo.set(bas_dakika_v)
                bit_saat_combo.set(bit_saat_v)
                bit_dakika_combo.set(bit_dakika_v)

    def varsayilanlara_sifirla(self):
        """Tüm periyot alanlarını, pencere ilk açıldığındaki gibi varsayılan
        saatlere döndürür (Kurs Ekle'de "Kaydet"ten sonra formu sıfırlamak
        için kullanılır)."""
        for girisler in (self.pzt_per_girisleri, self.cuma_girisleri):
            for i, (bas_saat_combo, bas_dakika_combo, bit_saat_combo, bit_dakika_combo) in enumerate(girisler):
                varsayilan_bas, varsayilan_bit = VARSAYILAN_PERIYOT_SAATLERI[i]
                bas_saat_v, bas_dakika_v = varsayilan_bas.split(":")
                bit_saat_v, bit_dakika_v = varsayilan_bit.split(":")
                bas_saat_combo.set(bas_saat_v)
                bas_dakika_combo.set(bas_dakika_v)
                bit_saat_combo.set(bit_saat_v)
                bit_dakika_combo.set(bit_dakika_v)

    def saatleri_al(self):
        """Girilen tüm alanları doğrular ve dersler.ders_ekle/ders_guncelle
        için (gun, sira, baslangic, bitis) satırlarının listesini döndürür.
        Geçersiz bir değer varsa ValueError fırlatır (mesajı kullanıcıya
        gösterilebilir)."""
        satirlar = []
        for gunler, girisler in ((range(0, 4), self.pzt_per_girisleri), ((4,), self.cuma_girisleri)):
            periyot_saatleri = []
            for sira, (bas_saat_combo, bas_dakika_combo, bit_saat_combo, bit_dakika_combo) in enumerate(
                girisler, start=1
            ):
                bas_metin = saat_dakika_birlestir(bas_saat_combo.get(), bas_dakika_combo.get())
                bit_metin = saat_dakika_birlestir(bit_saat_combo.get(), bit_dakika_combo.get())
                if not bas_metin or not bit_metin:
                    raise ValueError(
                        f"{sira}. saat için hem saat hem dakika seçin."
                    )
                try:
                    bas_saat = saat_str_to_time(bas_metin)
                    bit_saat = saat_str_to_time(bit_metin)
                except ValueError:
                    raise ValueError(
                        f"{sira}. saat için geçerli bir saat girin (ÖR: 08:00)."
                    )
                if bit_saat <= bas_saat:
                    raise ValueError(
                        f"{sira}. saatte bitiş saati, başlangıç saatinden sonra olmalıdır."
                    )
                periyot_saatleri.append((bas_metin, bit_metin))
            for gun in gunler:
                for sira, (bas_metin, bit_metin) in enumerate(periyot_saatleri, start=1):
                    satirlar.append((gun, sira, bas_metin, bit_metin))
        return satirlar


class DersEkleDialog(tk.Toplevel):
    def __init__(self, master, db: VeriTabani):
        super().__init__(master)
        self.db = db
        self.title("Kurs Ekle")
        # İçerik (haftalık program dahil) uzun olduğu için, pencere ekran
        # yüksekliğine göre olabildiğince uzun açılır (ama ekrandan taşmaz);
        # içerik yine de kaydırılabilir (fare tekerleği dahil) bir alanda.
        # ÖNEMLİ: alt sınır (max(...)) ASLA ekran yüksekliğinden büyük
        # olmamalı - aksi halde pencere ekranın altına taşar ve "Kaydet"
        # butonu dahil alt kısım hiç görünmez olur (önceki hatanın sebebi
        # tam olarak buydu). Bu yüzden önce ekrana göre üst sınır konur,
        # sonra makul (küçük) bir alt sınırla taşma önlenir.
        pencere_genisligi = 540
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        pencere_yuksekligi = min(900, ekran_yuksekligi - 80)
        pencere_yuksekligi = max(480, pencere_yuksekligi)
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.minsize(480, 420)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)

        gruplar = self.db.gruplari_getir()
        self.grup_map = {g["ad"]: g["id"] for g in gruplar}
        # Derslikler artık gruptan bağımsız, kurumun TAMAMINA ait ORTAK bir
        # havuzdur (ör. Derslik 1..40); bu yüzden Grup seçiminden bağımsız
        # olarak baştan TÜMÜ listelenir. Ancak bir derslik, bitiş tarihine
        # kadar başka bir kursa ayrılmışsa (dolu ise) burada gözükmemeli -
        # bu yüzden TÜMÜ değil, MÜSAİT olanlar listelenir.
        derslikler = self.db.musait_derslikleri_getir()
        self.derslik_map = {d["ad"]: d["id"] for d in derslikler}

        # İçerik uzun olduğu için kaydırılabilir bir alana yerleştiriliyor.
        disaridaki = tk.Frame(self)
        disaridaki.pack(fill="both", expand=True)
        canvas = tk.Canvas(disaridaki, highlightthickness=0)
        kaydirma_cubugu = ttk.Scrollbar(disaridaki, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=kaydirma_cubugu.set)
        canvas.pack(side="left", fill="both", expand=True)
        kaydirma_cubugu.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(canvas)

        frm = tk.Frame(canvas, padx=16, pady=16)
        canvas_penceresi = canvas.create_window((0, 0), window=frm, anchor="nw")

        def _kaydirma_bolgesini_guncelle(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _icerik_genisligini_esitle(event):
            canvas.itemconfig(canvas_penceresi, width=event.width)

        frm.bind("<Configure>", _kaydirma_bolgesini_guncelle)
        canvas.bind("<Configure>", _icerik_genisligini_esitle)

        tk.Label(frm, text="Grup").grid(row=0, column=0, sticky="w", pady=4)
        self.grup_combo = ttk.Combobox(frm, values=list(self.grup_map.keys()), state="readonly")
        self.grup_combo.grid(row=0, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="Derslik").grid(row=1, column=0, sticky="w", pady=4)
        self.derslik_combo = ttk.Combobox(frm, values=list(self.derslik_map.keys()), state="readonly")
        self.derslik_combo.grid(row=1, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="Kurs adı").grid(row=2, column=0, sticky="w", pady=4)
        self.ad_entry = tk.Entry(frm)
        self.ad_entry.grid(row=2, column=1, sticky="ew", pady=4)

        ttk.Separator(frm, orient="horizontal").grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 8))

        tk.Label(
            frm,
            text="Haftalık ders programı (her gün 6 periyot; Pazartesi-Perşembe\n"
                 "aynı saatlerde, Cuma ayrı saatlerde girilir)",
            justify="left",
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self.program = HaftalikProgramCercevesi(frm)
        self.program.grid(row=5, column=0, columnspan=2, sticky="ew")

        ttk.Separator(frm, orient="horizontal").grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 8))

        tk.Label(frm, text="Başlangıç tarihi (GG.AA.YYYY)").grid(
            row=7, column=0, sticky="w", pady=4
        )
        self.baslangic_tarihi_entry = tk.Entry(frm)
        self.baslangic_tarihi_entry.grid(row=7, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="Bitiş tarihi (GG.AA.YYYY)").grid(
            row=8, column=0, sticky="w", pady=4
        )
        self.bitis_tarihi_entry = tk.Entry(frm)
        self.bitis_tarihi_entry.grid(row=8, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="Yoklama girme süresi\n(dakika)", justify="left").grid(
            row=9, column=0, sticky="w", pady=4
        )
        self.sure_spin = tk.Spinbox(frm, from_=1, to=180, width=6)
        self.sure_spin.delete(0, "end")
        self.sure_spin.insert(0, str(DEVAMSIZLIK_SURESI_DK))
        self.sure_spin.grid(row=9, column=1, sticky="w", pady=4)

        frm.grid_columnconfigure(1, weight=1)

        tk.Button(
            frm, text="Kaydet", bg=RENK_VURGU, fg="white",
            font=("Segoe UI", 10, "bold"), command=self._kaydet
        ).grid(row=10, column=0, columnspan=2, pady=16, ipady=6, sticky="ew")

        # Enter/Leave tabanlı fare tekeri bağlamasına (yukarıda) ek olarak,
        # tüm içeriğe DOĞRUDAN da bağlanıyor - bazı ortamlarda Enter/Leave
        # olaylarının beklendiği gibi tetiklenmemesine karşı ek güvence.
        _agaca_fare_tekeri_ekle(frm, canvas)

    def _kaydet(self):
        grup_ad = self.grup_combo.get()
        derslik_ad = self.derslik_combo.get()
        ad = self.ad_entry.get().strip()

        if not (grup_ad and derslik_ad and ad):
            messagebox.showerror(APP_ADI, "Lütfen tüm zorunlu alanları doldurun.", parent=self)
            return

        try:
            saatler = self.program.saatleri_al()
        except ValueError as e:
            messagebox.showerror(APP_ADI, str(e), parent=self)
            return

        baslangic_tarihi = tarih_metnini_isoya_cevir(self.baslangic_tarihi_entry.get())
        bitis_tarihi = tarih_metnini_isoya_cevir(self.bitis_tarihi_entry.get())
        if baslangic_tarihi is None or bitis_tarihi is None:
            messagebox.showerror(
                APP_ADI,
                "Tarihleri GG.AA.YYYY biçiminde girin (örn. 03.11.2026) ya da "
                "sınırsız olması için boş bırakın.",
                parent=self,
            )
            return
        if baslangic_tarihi and bitis_tarihi and bitis_tarihi < baslangic_tarihi:
            messagebox.showerror(
                APP_ADI, "Bitiş tarihi, başlangıç tarihinden önce olamaz.", parent=self
            )
            return
        try:
            sure_dk = int(self.sure_spin.get())
            if not (1 <= sure_dk <= 180):
                raise ValueError
        except ValueError:
            messagebox.showerror(
                APP_ADI, "Yoklama girme süresi 1 ile 180 dakika arasında olmalıdır.", parent=self
            )
            return

        grup_id = self.grup_map[grup_ad]
        derslik_id = self.derslik_map[derslik_ad]

        self.db.ders_ekle(
            grup_id, derslik_id, ad, saatler,
            baslangic_tarihi, bitis_tarihi, sure_dk,
        )
        messagebox.showinfo(APP_ADI, "Kurs eklendi.", parent=self)
        # Pencereyi kapatmak yerine, ilk açıldığındaki gibi BOŞ hale
        # getiriyoruz; böylece art arda birden fazla kurs eklenecekse
        # pencereyi kapatıp yeniden açmaya gerek kalmıyor.
        self._formu_sifirla()

    def _formu_sifirla(self):
        self.grup_combo.set("")
        # Az önce kaydedilen kurs kendi dersliğini "dolu" yaptığı için,
        # listeyi tazeleyip artık müsait olmayan dersliği listeden
        # düşürüyoruz.
        derslikler = self.db.musait_derslikleri_getir()
        self.derslik_map = {d["ad"]: d["id"] for d in derslikler}
        self.derslik_combo["values"] = list(self.derslik_map.keys())
        self.derslik_combo.set("")
        self.ad_entry.delete(0, "end")
        self.program.varsayilanlara_sifirla()
        self.baslangic_tarihi_entry.delete(0, "end")
        self.bitis_tarihi_entry.delete(0, "end")
        self.sure_spin.delete(0, "end")
        self.sure_spin.insert(0, str(DEVAMSIZLIK_SURESI_DK))
        self.grup_combo.focus_set()


# --------------------------------------------------------------------------
# Kursiyer Ekle penceresi
# --------------------------------------------------------------------------

class KursiyerEkleDialog(tk.Toplevel):
    """Bir derse birden fazla kursiyeri, kişi sayısı kadar ayrı Kod / Ad
    Soyad alanı oluşturarak tek seferde ekler."""

    def __init__(self, master, db: VeriTabani):
        super().__init__(master)
        self.db = db
        self.title("Kursiyer Ekle")
        self.geometry("480x560")
        self.minsize(440, 360)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)

        gruplar = self.db.gruplari_getir()
        self.grup_map = {g["ad"]: g["id"] for g in gruplar}
        self.ders_map = {}
        self.satir_girdileri = []  # [(kod_entry, ad_entry), ...]

        frm = tk.Frame(self, padx=16, pady=16)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Grup").grid(row=0, column=0, sticky="w", pady=4)
        self.grup_combo = ttk.Combobox(frm, values=list(self.grup_map.keys()), state="readonly")
        self.grup_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.grup_combo.bind("<<ComboboxSelected>>", self._grup_secildi)

        tk.Label(frm, text="Kurs").grid(row=1, column=0, sticky="w", pady=4)
        self.ders_combo = ttk.Combobox(frm, values=[], state="readonly")
        self.ders_combo.grid(row=1, column=1, sticky="ew", pady=4)

        frm.grid_columnconfigure(1, weight=1)

        tk.Label(frm, text="Kaç kursiyer ekleyeceksiniz?").grid(
            row=2, column=0, sticky="w", pady=(14, 4)
        )
        sayi_cerceve = tk.Frame(frm)
        sayi_cerceve.grid(row=2, column=1, sticky="w", pady=(14, 4))
        self.sayi_spin = tk.Spinbox(sayi_cerceve, from_=1, to=200, width=5)
        self.sayi_spin.delete(0, "end")
        self.sayi_spin.insert(0, "5")
        self.sayi_spin.pack(side="left")
        tk.Button(
            sayi_cerceve, text="Oluştur", command=self._satirlari_olustur
        ).pack(side="left", padx=(8, 0))

        baslik_cerceve = tk.Frame(frm)
        baslik_cerceve.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 2))
        tk.Label(baslik_cerceve, text="No", width=4, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Label(baslik_cerceve, text="Kod", width=12, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(4, 4))
        tk.Label(baslik_cerceve, text="Ad Soyad", anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(4, 0))

        liste_cerceve = tk.Frame(frm)
        liste_cerceve.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        frm.grid_rowconfigure(4, weight=1)

        canvas = tk.Canvas(liste_cerceve, highlightthickness=0)
        self.canvas = canvas
        dikey_kaydirma = ttk.Scrollbar(liste_cerceve, orient="vertical", command=canvas.yview)
        self.satirlar_frame = tk.Frame(canvas)
        self.satirlar_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.satirlar_frame, anchor="nw")
        canvas.configure(yscrollcommand=dikey_kaydirma.set)
        canvas.pack(side="left", fill="both", expand=True)
        dikey_kaydirma.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(canvas)

        tk.Button(
            frm, text="Hepsini Kaydet", bg=RENK_VURGU, fg="white",
            font=("Segoe UI", 10, "bold"), command=self._kaydet
        ).grid(row=5, column=0, columnspan=2, pady=(6, 0), ipady=8, sticky="ew")

        self._satirlari_olustur()

    def _grup_secildi(self, event=None):
        grup_id = self.grup_map[self.grup_combo.get()]
        dersler = self.db.dersleri_getir(grup_id=grup_id)
        self.ders_map = {
            f"{d['ad']} ({d['derslik_ad']})": d["id"]
            for d in dersler
        }
        self.ders_combo["values"] = list(self.ders_map.keys())
        self.ders_combo.set("")

    def _satir_ekle(self, sira_no):
        satir = tk.Frame(self.satirlar_frame)
        satir.pack(fill="x", pady=2)
        tk.Label(satir, text=f"{sira_no}.", width=4, anchor="w").pack(side="left")
        kod_entry = tk.Entry(satir, width=12)
        kod_entry.pack(side="left", padx=(4, 4))
        ad_entry = tk.Entry(satir, width=28)
        ad_entry.pack(side="left", padx=(4, 0), fill="x", expand=True)
        self.satir_girdileri.append((kod_entry, ad_entry))
        _agaca_fare_tekeri_ekle(satir, self.canvas)

    def _satirlari_olustur(self):
        """Kişi sayısı kutusuna göre satır oluşturur/kaldırır. Sayı
        ARTIRILDIĞINDA daha önce doldurulmuş satırlara dokunmaz, sadece
        eksik kalan satırları ekler; sayı AZALTILDIĞINDA ise sadece
        fazlalık satırları kaldırır (dolu satırlar varsa önce onaylatır).
        Böylece örneğin 5 kişi girip isim/kod doldurduktan sonra sayıyı
        7'ye çıkarmak, ilk 5'i silmez - sadece 6. ve 7. satırları ekler."""
        try:
            sayi = int(self.sayi_spin.get())
        except ValueError:
            messagebox.showerror(APP_ADI, "Lütfen geçerli bir sayı girin.", parent=self)
            return
        if not (1 <= sayi <= 200):
            messagebox.showerror(APP_ADI, "Kişi sayısı 1 ile 200 arasında olmalıdır.", parent=self)
            return

        mevcut_sayi = len(self.satir_girdileri)

        if sayi == mevcut_sayi:
            return

        if sayi > mevcut_sayi:
            for i in range(mevcut_sayi + 1, sayi + 1):
                self._satir_ekle(i)
            self.satir_girdileri[mevcut_sayi][0].focus_set()
            return

        # sayi < mevcut_sayi: sondaki fazlalık satırları kaldırıyoruz.
        kaldirilacaklar = self.satir_girdileri[sayi:]
        doluysa = any(
            kod_e.get().strip() or ad_e.get().strip() for kod_e, ad_e in kaldirilacaklar
        )
        if doluysa and not messagebox.askyesno(
            APP_ADI,
            f"Kişi sayısını azaltıyorsunuz; son {len(kaldirilacaklar)} satırdaki "
            "girilmiş bilgiler silinecek. Devam edilsin mi?",
            parent=self,
        ):
            # Kullanıcı vazgeçti; sayı kutusunu mevcut duruma geri al.
            self.sayi_spin.delete(0, "end")
            self.sayi_spin.insert(0, str(mevcut_sayi))
            return

        for kod_entry, _ad_entry in kaldirilacaklar:
            kod_entry.master.destroy()
        self.satir_girdileri = self.satir_girdileri[:sayi]

    def _kaydet(self):
        grup_ad = self.grup_combo.get()
        ders_secim = self.ders_combo.get()
        if not (grup_ad and ders_secim):
            messagebox.showerror(APP_ADI, "Lütfen önce grup ve kurs seçin.", parent=self)
            return
        if not self.satir_girdileri:
            messagebox.showerror(
                APP_ADI,
                "Lütfen önce kaç kursiyer ekleyeceğinizi girip \"Oluştur\" butonuna basın.",
                parent=self,
            )
            return

        ders_id = self.ders_map[ders_secim]

        eklenen = 0
        atlanan = []  # (satir_no, sebep)
        bu_turda_kullanilanlar = set()

        for i, (kod_entry, ad_entry) in enumerate(self.satir_girdileri, start=1):
            kod = kod_entry.get().strip()
            ad = ad_entry.get().strip()
            if not kod and not ad:
                continue  # boş satır, sessizce atla
            if not kod:
                atlanan.append((i, "kod girilmemiş (ad soyad yalnız yeterli değil)"))
                continue
            if kod in bu_turda_kullanilanlar or self.db.kod_kullanilmis_mi(ders_id, kod):
                atlanan.append((i, f"“{kod}” kodu bu kursta zaten kullanılıyor"))
                continue
            self.db.kursiyer_ekle(ders_id, kod, ad)
            bu_turda_kullanilanlar.add(kod)
            eklenen += 1

        mesaj = f"{eklenen} kursiyer eklendi."
        if atlanan:
            gosterilecek = atlanan[:15]
            detay = "\n".join(f"{no}. satır: {sebep}" for no, sebep in gosterilecek)
            if len(atlanan) > 15:
                detay += f"\n... ve {len(atlanan) - 15} satır daha"
            mesaj += f"\n\n{len(atlanan)} satır atlandı:\n{detay}"

        if eklenen > 0:
            messagebox.showinfo(APP_ADI, mesaj, parent=self)
            # Pencereyi kapatmak yerine, ilk açıldığındaki gibi BOŞ hale
            # getiriyoruz; böylece farklı bir kursa kursiyer eklenecekse
            # pencereyi kapatıp yeniden açmaya gerek kalmıyor.
            self._formu_sifirla()
        else:
            messagebox.showerror(APP_ADI, mesaj, parent=self)

    def _formu_sifirla(self):
        self.grup_combo.set("")
        self.ders_combo.set("")
        self.ders_combo["values"] = []
        self.ders_map = {}
        for kod_entry, _ad_entry in self.satir_girdileri:
            kod_entry.master.destroy()
        self.satir_girdileri = []
        self.sayi_spin.delete(0, "end")
        self.sayi_spin.insert(0, "5")
        self._satirlari_olustur()
        self.grup_combo.focus_set()


# --------------------------------------------------------------------------
# Yoklama alma penceresi (öğretmen)
# --------------------------------------------------------------------------

class YoklamaAlDialog(tk.Toplevel):
    """Öğretmen, yöneticinin o ders için belirlediği süre boyunca bu
    pencereyi istediği kadar açıp yoklamayı kaydedebilir/düzeltebilir;
    her kayıtta 'son güncelleme zamanı' güncellenir. Bir kursiyer 'Yok'
    olarak işaretlenirse, sağında yöneticinin tanımladığı mazeret
    listesinden bir seçim yapılabilir."""

    def __init__(self, master, db: VeriTabani, ders_row):
        super().__init__(master)
        self.db = db
        self.ders = ders_row
        self.title(f"Yoklama - {ders_row['ad']} - {ders_row['sira']}. Saat")
        pencere_genisligi = 540
        pencere_yuksekligi = 600
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.configure(bg=RENK_ARKAPLAN)
        self.grab_set()
        _grab_guvenli_baglat(self)
        # ÖNEMLİ: Bu pencerenin açıldığı DerslikDersEkrani artık "-topmost"
        # (her zaman önde) olduğu için, bu pencere de topmost YAPILMAZSA
        # Windows onu kendi (topmost olmayan) sırasına koyar ve pencere,
        # kendi ÜST penceresinin ARKASINA gizlenmiş halde açılabilirdi -
        # kilit (grab) hâlâ ona ait olduğu için de kullanıcı görünürdeki
        # (üstteki) pencereye tıkladığında hiçbir şey olmuyor, uygulama
        # "donmuş" gibi görünüyordu. Bu pencere de topmost yapılarak,
        # üstteki topmost pencerenin ÖNÜNE çıkması garanti ediliyor.
        self.transient(master)
        self.attributes("-topmost", True)
        self.lift()
        self.focus_force()
        _baslik_cubugu_rengini_ayarla(self)

        bugun_str = date.today().strftime("%Y-%m-%d")
        self.oturum = self.db.oturum_olustur_veya_getir(
            ders_row["id"], ders_row["ders_saat_id"], bugun_str
        )
        mevcut_kayitlar = {
            k["kursiyer_id"]: k
            for k in self.db.yoklama_kayitlarini_getir(self.oturum["id"])
        } if self.oturum else {}

        ust_serit = tk.Frame(self, bg=RENK_VURGU)
        ust_serit.pack(fill="x")
        tk.Label(
            ust_serit,
            text=f"{ders_row['ad']}  |  {ders_row['derslik_ad']}",
            bg=RENK_VURGU, fg="white", font=("Segoe UI", 13, "bold"),
        ).pack(pady=(14, 0))
        tk.Label(
            ust_serit,
            text=f"{GUN_ADLARI[ders_row['gun']]}  •  {ders_row['sira']}. Saat  "
                 f"{ders_row['baslangic']} - {ders_row['bitis']}",
            bg=RENK_VURGU, fg="#e0e7ff", font=("Segoe UI", 10),
        ).pack(pady=(0, 12))

        govde_dis = tk.Frame(self, bg=RENK_ARKAPLAN)
        govde_dis.pack(fill="both", expand=True, padx=14, pady=(12, 0))

        govde = tk.Frame(govde_dis, bg="white", highlightbackground="#d8dbe0", highlightthickness=1)
        govde.pack(fill="both", expand=True)

        canvas = tk.Canvas(govde, highlightthickness=0, bg="white")
        scrollbar = ttk.Scrollbar(govde, orient="vertical", command=canvas.yview)
        self.liste_frame = tk.Frame(canvas, bg="white")
        self.liste_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.liste_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        scrollbar.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(canvas)

        self.kursiyerler = self.db.kursiyerleri_getir(ders_row["id"])
        self.mazeretler = self.db.mazeretleri_getir()
        self.mazeret_metinleri = [m["metin"] for m in self.mazeretler]
        self.mazeret_id_by_metin = {m["metin"]: m["id"] for m in self.mazeretler}

        self.durum_vars = {}
        self.mazeret_vars = {}

        if not self.kursiyerler:
            tk.Label(
                self.liste_frame, text="Bu kursa henüz kursiyer eklenmemiş.",
                bg="white", fg="#777777",
            ).pack(pady=20)
        else:
            baslik = tk.Frame(self.liste_frame, bg="#eef1f6")
            baslik.pack(fill="x")
            tk.Label(baslik, text="Kod", width=8, anchor="w", bg="#eef1f6", font=("Segoe UI", 9, "bold")).pack(side="left", pady=6)
            tk.Label(baslik, text="Ad Soyad", width=16, anchor="w", bg="#eef1f6", font=("Segoe UI", 9, "bold")).pack(side="left", pady=6)
            tk.Label(baslik, text="MEVCUT", width=8, anchor="w", bg="#eef1f6", font=("Segoe UI", 9, "bold")).pack(side="left", pady=6)
            tk.Label(baslik, text="MAZERET SEBEBİ", anchor="w", bg="#eef1f6", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(20, 0), pady=6)

            for i, k in enumerate(self.kursiyerler):
                satir_bg = "#ffffff" if i % 2 == 0 else "#f5f7fa"
                satir = tk.Frame(self.liste_frame, bg=satir_bg)
                satir.pack(fill="x")
                tk.Label(satir, text=k["kod"], width=8, anchor="w", bg=satir_bg).pack(side="left", pady=5)
                tk.Label(satir, text=k["ad_soyad"] or "-", width=16, anchor="w", bg=satir_bg).pack(side="left", pady=5)

                onceki = mevcut_kayitlar.get(k["id"])
                var_baslangic = (onceki["durum"] == "var") if onceki else True
                var = tk.BooleanVar(value=var_baslangic)

                mazeret_var = tk.StringVar(value="")
                if onceki and onceki["mazeret_metin"]:
                    mazeret_var.set(onceki["mazeret_metin"])

                if self.mazeret_metinleri:
                    mazeret_widget = ttk.Combobox(
                        satir, values=self.mazeret_metinleri, textvariable=mazeret_var,
                        state="readonly", width=16,
                    )
                else:
                    mazeret_widget = tk.Label(satir, text="(mazeret tanımlı değil)", bg=satir_bg, fg="#999999")

                def _tik_degisti(var=var, widget=mazeret_widget):
                    if var.get():
                        widget.pack_forget()
                    else:
                        widget.pack(side="left", padx=(20, 0))

                tk.Checkbutton(satir, variable=var, bg=satir_bg, activebackground=satir_bg, command=_tik_degisti).pack(side="left")

                if not var_baslangic:
                    mazeret_widget.pack(side="left", padx=(20, 0))

                self.durum_vars[k["id"]] = var
                self.mazeret_vars[k["id"]] = mazeret_var

        _agaca_fare_tekeri_ekle(self.liste_frame, canvas)

        alt = tk.Frame(self, bg=RENK_ARKAPLAN)
        alt.pack(fill="x", pady=12, padx=14)
        tk.Label(
            alt, text="İşaretli = Var, işaretsiz = Yok (yanında mazeret seçilebilir)",
            bg=RENK_ARKAPLAN, fg="#555555",
        ).pack(anchor="w")
        if self.oturum and self.oturum["durum"] == "tamamlandi":
            tk.Label(
                alt, text=f"Bu oturum için daha önce kayıt yapılmış (son güncelleme: "
                          f"{self.oturum['giris_zamani']}). Tekrar kaydederseniz üzerine yazılır.",
                bg=RENK_ARKAPLAN, fg="#b45309", wraplength=460, justify="left",
            ).pack(anchor="w", pady=(4, 0))
        tk.Button(
            alt, text="Yoklamayı Kaydet", bg=RENK_YESIL, fg="white",
            activebackground="#15803d", activeforeground="white",
            font=("Segoe UI", 11, "bold"), relief="flat", bd=0, command=self._kaydet
        ).pack(fill="x", pady=(8, 0), ipady=8)

    def _kaydet(self):
        bugun_str = date.today().strftime("%Y-%m-%d")
        oturum = self.db.oturum_olustur_veya_getir(
            self.ders["id"], self.ders["ders_saat_id"], bugun_str
        )
        if oturum["durum"] == "ogretmen_girmedi":
            sure_dk = self.ders["devamsizlik_suresi_dk"] or DEVAMSIZLIK_SURESI_DK
            # NOT: Bu pencerenin (self) "-topmost" ÖZELLİĞİNİ messagebox'tan
            # önce kapatmak YANLIŞTI - bu pencereyi açan DerslikDersEkrani de
            # topmost olduğu için, self topmost'u kapatılınca hem self hem de
            # ona bağlı (owner'ı self olan) messagebox, DerslikDersEkrani'nin
            # ARKASINDA kalıyordu (kaydet sonrası "kayboluyor" gibi
            # görünmesinin sebebi buydu). Doğrusu: self'i topmost bırakmak
            # ve tekrar lift() ile diğer topmost pencerelerin (DerslikDersEkrani
            # dahil) önüne almak - messagebox zaten self'in "sahibi" (owner'ı)
            # olduğu için doğal olarak self'in önünde açılır.
            self.lift()
            self.focus_force()
            messagebox.showerror(
                APP_ADI,
                f"Bu kurs için {sure_dk} dakikalık giriş süresi doldu ve sistem "
                "otomatik olarak devamsız işaretledi. Düzeltme gerekiyorsa "
                "yöneticiye başvurun.",
                parent=self,
            )
            self.destroy()
            return
        kayitlar = []
        for kursiyer_id, var in self.durum_vars.items():
            if var.get():
                kayitlar.append((kursiyer_id, "var", None))
            else:
                mazeret_metin = self.mazeret_vars[kursiyer_id].get().strip()
                mazeret_id = self.mazeret_id_by_metin.get(mazeret_metin)
                kayitlar.append((kursiyer_id, "yok", mazeret_id))
        if kayitlar:
            self.db.yoklama_kaydet(oturum["id"], kayitlar)
        else:
            # kursiyer yoksa da oturumu tamamlandı olarak işaretle
            conn = self.db.baglan()
            try:
                conn.execute(
                    "UPDATE oturumlar SET durum='tamamlandi', giris_zamani=? WHERE id=?",
                    (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), oturum["id"]),
                )
                conn.commit()
            finally:
                conn.close()
        # bkz. yukarıdaki aynı not: topmost KAPATILMIYOR, tam tersine lift()
        # ile self (ve dolayısıyla owner'ı self olan messagebox) tekrar
        # DerslikDersEkrani'nin önüne alınıyor.
        self.lift()
        self.focus_force()
        messagebox.showinfo(APP_ADI, "Yoklama kaydedildi.", parent=self)
        self.destroy()


# --------------------------------------------------------------------------
# Grup ekranı (bir gruba şifreyle girildikten sonra açılan pencere)
# --------------------------------------------------------------------------

class GrupEkrani(tk.Toplevel):
    """Bir gruba şifreyle girildikten sonra açılan pencere: o gruba ait
    dersliklerin listesini gösterir. Öğretmen, içine girmek istediği
    dersliğe tıklar."""

    def __init__(self, master, db: VeriTabani, grup_row):
        super().__init__(master)
        self.db = db
        self.grup = grup_row
        self.title(f"{grup_row['ad']} - Derslikler")
        # Pencere konumu belirtilmezse Tk bunu ekranın SOL ÜST köşesine
        # yakın açıyordu; ekranda ortalanmış (biraz sağa alınmış) şekilde
        # açılması için x/y burada hesaplanıyor.
        pencere_genisligi = 520
        pencere_yuksekligi = 560
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.minsize(440, 420)
        self.configure(bg=RENK_ARKAPLAN)
        _baslik_cubugu_rengini_ayarla(self)
        # Öğretmen kendi grubuna girdiğinde bu pencere ana sayfanın hep
        # ÖNÜNDE kalsın istendi. Önceden bunun için grab_set()
        # kullanılmıştı, ANCAK grab_set() uygulamadaki TÜM diğer
        # pencereleri (Yönetici Paneli dahil) tıklanamaz hale getiriyor -
        # yönetici hem kendi panelini hem bir öğretmen ekranını aynı anda
        # açtığında panelinde işlem yapamaz hale geliyordu. Bunun yerine
        # "-topmost" kullanılıyor: bu pencere her zaman ana sayfanın
        # ÖNÜNDE görünür ama DİĞER pencereleri (ör. Yönetici Paneli)
        # kilitlemez, ikisinde de aynı anda çalışılabilir.
        self.transient(master)
        self.attributes("-topmost", True)
        self.lift()
        self.focus_force()

        ust_serit = tk.Frame(self, bg=RENK_SOL_PANEL)
        ust_serit.pack(fill="x")
        tk.Label(
            ust_serit, text=grup_row["ad"], bg=RENK_SOL_PANEL, fg=RENK_METIN_ACIK,
            font=("Segoe UI", 16, "bold"),
        ).pack(pady=(18, 2))
        tk.Label(
            ust_serit, text="Yoklama almak istediğiniz dersliğe tıklayın",
            bg=RENK_SOL_PANEL, fg="#c7cdd6",
        ).pack(pady=(0, 16))

        disaridaki = tk.Frame(self, bg=RENK_ARKAPLAN)
        disaridaki.pack(fill="both", expand=True)
        canvas = tk.Canvas(disaridaki, highlightthickness=0, bg=RENK_ARKAPLAN)
        kaydirma_cubugu = ttk.Scrollbar(disaridaki, orient="vertical", command=canvas.yview)
        self.liste_frame = tk.Frame(canvas, bg=RENK_ARKAPLAN)
        self.liste_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.liste_frame, anchor="nw")
        canvas.configure(yscrollcommand=kaydirma_cubugu.set)
        canvas.pack(side="left", fill="both", expand=True, padx=16, pady=16)
        kaydirma_cubugu.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(canvas)
        self.canvas = canvas

        self._doldur()
        # Pencere gün değişimi (gece yarısı) sırasında açık bırakılmış
        # olabilir; bir kursun bitiş tarihi geçtiğinde listeden düşmesi
        # için kullanıcı elle bir şey yapmasa bile periyodik olarak
        # kendiliğinden tazelenir.
        self._oto_yenile_id = None
        self._oto_yenile_planla()

    def _oto_yenile_planla(self):
        self._oto_yenile_id = self.after(60_000, self._oto_yenile)

    def _oto_yenile(self):
        if not self.winfo_exists():
            return
        self._doldur()
        self._oto_yenile_planla()

    def _doldur(self):
        for w in self.liste_frame.winfo_children():
            w.destroy()
        # Derslikler kurumun tamamına ait ORTAK bir havuz olduğu için,
        # burada her buton "Kurs adı (Derslik adı)" şeklinde, öğretmenin
        # kendi grubuna ait ve o an aktif olan kursları gösterir - sadece
        # derslik adı değil, hangi kursun orada olduğu da görünsün diye.
        bugun_str = date.today().strftime("%Y-%m-%d")
        tum_dersler = self.db.dersleri_getir(grup_id=self.grup["id"])
        aktif_dersler = [d for d in tum_dersler if ders_tarih_araliginda_mi(d, bugun_str)]
        if not aktif_dersler:
            tk.Label(
                self.liste_frame,
                text="Bu grubun şu anda aktif (tarih aralığı içinde) bir kursu yok.",
                bg=RENK_ARKAPLAN, fg="#777777", wraplength=420, justify="left",
            ).pack(pady=30)
            return
        for ders in aktif_dersler:
            tk.Button(
                self.liste_frame, text=f"{ders['ad']}   ({ders['derslik_ad']})",
                anchor="w", relief="flat", bd=0,
                cursor="hand2", bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK,
                activebackground=RENK_VURGU, activeforeground="white",
                font=("Segoe UI", 12, "bold"), padx=14,
                command=lambda d=ders: self._derslik_sec({"id": d["derslik_id"], "ad": d["derslik_ad"]}),
            ).pack(fill="x", pady=4, ipady=10)
        _agaca_fare_tekeri_ekle(self.liste_frame, self.canvas)

    def _derslik_sec(self, derslik_row):
        pencere = DerslikDersEkrani(self, self.db, self.grup, derslik_row)
        self.wait_window(pencere)
        # -topmost kullanıldığı için (bkz. __init__'teki not) burada
        # kilit yeniden kurmaya gerek yok; sadece pencere tekrar görünür/
        # öne gelsin diye lift/focus yapılıyor.
        try:
            if self.winfo_exists():
                self.lift()
                self.focus_force()
        except tk.TclError:
            pass


class DerslikDersEkrani(tk.Toplevel):
    """Bir dersliğe girildikten sonra açılan pencere: gün sekmeleri ve
    seçilen güne ait periyot (saat) kartlarını gösterir. Sadece bugüne
    ait kartlarda yoklama alma işlemi yapılabilir."""

    def __init__(self, master, db: VeriTabani, grup_row, derslik_row):
        super().__init__(master)
        self.db = db
        self.grup = grup_row
        self.derslik = derslik_row
        self.title(f"{derslik_row['ad']} - {grup_row['ad']}")
        # 6 periyot kartının tamamı (özellikle 5. ve 6. saat) sabit 580px
        # yükseklikte sığmıyordu; pencere artık ekran yüksekliğine göre
        # olabildiğince aşağı doğru genişler (ama ekrandan taşmaz - bkz.
        # Yoklama Paneli / Kurs Ekle'de daha önce yaşanan aynı hata).
        pencere_genisligi = 640
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        pencere_yuksekligi = min(820, ekran_yuksekligi - 80)
        pencere_yuksekligi = max(460, pencere_yuksekligi)
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.minsize(560, 460)
        self.configure(bg=RENK_ARKAPLAN)
        _baslik_cubugu_rengini_ayarla(self)
        # Bu pencere de (yoklama alınan asıl ekran) ana sayfanın hep
        # ÖNÜNDE kalsın istendi - bkz. GrupEkrani'ndeki aynı not (grab_set
        # yerine -topmost: diğer pencereleri - ör. Yönetici Paneli -
        # kilitlemeden ana sayfanın önünde kalmayı sağlıyor).
        self.transient(master)
        self.attributes("-topmost", True)
        self.lift()
        self.focus_force()
        self.secili_gun_idx = bugun_gun_index()
        if self.secili_gun_idx >= 5:  # hafta sonu girişse haftanın ilk gününü göster
            self.secili_gun_idx = 0

        ust_serit = tk.Frame(self, bg=RENK_SOL_PANEL)
        ust_serit.pack(fill="x")
        tk.Label(
            ust_serit, text=f"{derslik_row['ad']}  ({grup_row['ad']})",
            bg=RENK_SOL_PANEL, fg=RENK_METIN_ACIK, font=("Segoe UI", 16, "bold"),
        ).pack(pady=(18, 2))
        self.alt_baslik_label = tk.Label(ust_serit, text="", bg=RENK_SOL_PANEL, fg="#c7cdd6")
        self.alt_baslik_label.pack(pady=(0, 16))

        gun_cerceve = tk.Frame(self, bg=RENK_ARKAPLAN)
        gun_cerceve.pack(pady=(14, 10))
        self.gun_butonlari = {}
        for idx, gun_ad in enumerate(HAFTA_ICI_GUNLERI):
            b = tk.Button(
                gun_cerceve, text=gun_ad, width=9, relief="flat", bd=0, cursor="hand2",
                bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK, activebackground=RENK_VURGU, activeforeground="white",
                command=lambda i=idx: self._gun_sec(i),
            )
            b.pack(side="left", padx=2, ipady=4)
            self.gun_butonlari[idx] = b

        disaridaki = tk.Frame(self, bg=RENK_ARKAPLAN)
        disaridaki.pack(fill="both", expand=True, padx=16, pady=(0, 4))
        canvas = tk.Canvas(disaridaki, highlightthickness=0, bg=RENK_ARKAPLAN)
        kaydirma_cubugu = ttk.Scrollbar(disaridaki, orient="vertical", command=canvas.yview)
        self.liste_frame = tk.Frame(canvas, bg=RENK_ARKAPLAN)
        self.liste_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.liste_frame, anchor="nw")
        canvas.configure(yscrollcommand=kaydirma_cubugu.set)
        canvas.pack(side="left", fill="both", expand=True)
        kaydirma_cubugu.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(canvas)
        self.canvas = canvas

        tk.Button(
            self, text="Yenile", command=self._yenile,
            bg="#e5e7eb", fg="#1f2937", relief="flat", bd=0, cursor="hand2",
        ).pack(pady=(0, 16), ipadx=14, ipady=5)

        self._gun_sec(self.secili_gun_idx)
        # Bu pencere açık kalmışken bir kursun bitiş tarihi geçerse (ör.
        # gece yarısını geçince), elle "Yenile" basılmasa bile kurs
        # periyotlarının listeden kendiliğinden kalkması için periyodik
        # otomatik tazeleme.
        self._oto_yenile_id = None
        self._oto_yenile_planla()

    def _oto_yenile_planla(self):
        self._oto_yenile_id = self.after(60_000, self._oto_yenile)

    def _oto_yenile(self):
        if not self.winfo_exists():
            return
        self._yenile()
        self._oto_yenile_planla()

    def _gun_sec(self, gun_idx):
        self.secili_gun_idx = gun_idx
        for idx, b in self.gun_butonlari.items():
            if idx == gun_idx:
                b.configure(bg=RENK_VURGU, fg="white")
            else:
                b.configure(bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK)
        self._yenile()

    def _yenile(self):
        for w in self.liste_frame.winfo_children():
            w.destroy()

        self.db.otomatik_devamsizlik_kontrolu()

        bugun_idx = bugun_gun_index()
        bugun_str = date.today().strftime("%Y-%m-%d")
        bugun_mu = (self.secili_gun_idx == bugun_idx)

        # ÖNEMLİ: Gün sekmeleri (Pzt..Cuma) hep İÇİNDE BULUNULAN HAFTAyı
        # gösterir. Bir kurs sadece haftanın BİR KISMINDA aktifse (ör.
        # başlangıcı Salı, bitişi Çarşamba), "bugün aktif mi" kontrolü tek
        # başına yeterli değildir - o zaman kurs, aktif olduğu haftanın
        # Pazartesi/Perşembe/Cuma sekmelerinde de (aslında aktif olmadığı
        # günlerde) yanlışlıkla görünmeye devam ederdi. Bunun yerine
        # SEÇİLİ SEKMENİN kendi takvim tarihi hesaplanır ve kursun o
        # GÜN aktif olup olmadığına bakılır.
        haftanin_pazartesi = date.today() - timedelta(days=bugun_idx)
        secili_tarih = haftanin_pazartesi + timedelta(days=self.secili_gun_idx)
        secili_tarih_str = secili_tarih.strftime("%Y-%m-%d")

        if bugun_mu:
            self.alt_baslik_label.config(
                text=f"{GUN_ADLARI[self.secili_gun_idx]}  -  {date.today().strftime('%d.%m.%Y')}  (bugün)"
            )
        else:
            self.alt_baslik_label.config(
                text=f"{GUN_ADLARI[self.secili_gun_idx]}  -  {secili_tarih.strftime('%d.%m.%Y')}"
            )

        # Derslikler ortak bir havuz olduğu için, bu dersliği kullanan ama
        # BAŞKA bir gruba ait olabilecek kurslar burada gösterilmez. Kursun
        # bu sekmede gösterilip gösterilmeyeceği, o sekmenin GERÇEK takvim
        # tarihine göre belirlenir (bkz. yukarıdaki not) - "bugün" değil.
        aktif_dersler = [
            d for d in self.db.derslik_aktif_dersleri(self.derslik["id"], secili_tarih_str)
            if d["grup_id"] == self.grup["id"]
        ]
        gunun_saatleri = []
        for ders in aktif_dersler:
            for s in self.db.ders_saatleri_getir(ders["id"]):
                if s["gun"] == self.secili_gun_idx:
                    gunun_saatleri.append({
                        "id": ders["id"],
                        "ders_saat_id": s["id"],
                        "ad": ders["ad"],
                        "derslik_ad": ders["derslik_ad"],
                        "devamsizlik_suresi_dk": ders["devamsizlik_suresi_dk"],
                        "gun": s["gun"],
                        "sira": s["sira"],
                        "baslangic": s["baslangic"],
                        "bitis": s["bitis"],
                    })
        gunun_saatleri.sort(key=lambda x: x["sira"])

        if not gunun_saatleri:
            tk.Label(
                self.liste_frame, text="Bu derslikte bu güne ait planlanmış kurs yok.",
                bg=RENK_ARKAPLAN, fg="#777777",
            ).pack(pady=30)
            return

        simdi = datetime.now()
        for saat in gunun_saatleri:
            kutu = tk.Frame(
                self.liste_frame, bg="white",
                highlightbackground="#d8dbe0", highlightthickness=1,
            )
            kutu.pack(fill="x", pady=6, padx=2)

            ust = tk.Frame(kutu, bg="white")
            ust.pack(fill="x", padx=12, pady=(10, 2))
            tk.Label(
                ust, text=f"{saat['sira']}. Saat   {saat['ad']}",
                bg="white", font=("Segoe UI", 11, "bold"),
            ).pack(side="left")
            tk.Label(ust, text=f"{saat['baslangic']} - {saat['bitis']}", bg="white").pack(side="right")

            alt = tk.Frame(kutu, bg="white")
            alt.pack(fill="x", padx=12, pady=(2, 10))

            if not bugun_mu:
                continue

            sure_dk = saat["devamsizlik_suresi_dk"] or DEVAMSIZLIK_SURESI_DK
            oturum = self.db.oturum_getir(saat["ders_saat_id"], bugun_str)
            bas_dt = datetime.combine(date.today(), saat_str_to_time(saat["baslangic"]))
            son_giris_dt = bas_dt + timedelta(minutes=sure_dk)

            if oturum and oturum["durum"] == "tamamlandi":
                if simdi <= son_giris_dt:
                    tk.Label(
                        alt, text=f"Yoklama alındı (son güncelleme: {oturum['giris_zamani']})",
                        bg="white", fg=RENK_YESIL,
                    ).pack(side="left")
                    tk.Button(
                        alt, text="Yoklamayı Düzenle", bg=RENK_VURGU, fg="white",
                        relief="flat", bd=0, cursor="hand2",
                        command=lambda s=saat: self._yoklama_al(s),
                    ).pack(side="right", ipadx=6, ipady=2)
                else:
                    tk.Label(
                        alt,
                        text=f"Yoklama alındı (son güncelleme: {oturum['giris_zamani']}) - "
                             "düzenleme süresi doldu",
                        bg="white", fg=RENK_YESIL,
                    ).pack(side="left")
            elif oturum and oturum["durum"] == "ogretmen_girmedi":
                tk.Label(
                    alt, text="Öğretmen girmedi - devamsız işaretlendi", bg="white", fg=RENK_KIRMIZI,
                ).pack(side="left")
            elif simdi < bas_dt:
                tk.Label(
                    alt, text=f"Bu ders henüz başlamadı (saat {saat['baslangic']})",
                    bg="white", fg="#555555",
                ).pack(side="left")
            elif simdi <= son_giris_dt:
                kalan = int((son_giris_dt - simdi).total_seconds() // 60) + 1
                tk.Label(
                    alt, text=f"Yoklama girme süresi: yaklaşık {kalan} dk kaldı",
                    bg="white", fg="#b45309",
                ).pack(side="left")
                tk.Button(
                    alt, text="Yoklama Al", bg=RENK_VURGU, fg="white",
                    relief="flat", bd=0, cursor="hand2",
                    command=lambda s=saat: self._yoklama_al(s),
                ).pack(side="right", ipadx=6, ipady=2)
            else:
                tk.Label(alt, text="Giriş süresi doldu", bg="white", fg=RENK_KIRMIZI).pack(side="left")

        _agaca_fare_tekeri_ekle(self.liste_frame, self.canvas)

    def _yoklama_al(self, saat_row):
        pencere = YoklamaAlDialog(self, self.db, saat_row)
        self.wait_window(pencere)
        # -topmost kalıcı bir pencere özelliği olduğu için (grab_set'in
        # aksine) burada yeniden kurulmasına gerek yok; sadece pencere
        # tekrar görünür/öne gelsin diye lift/focus yapılıyor.
        try:
            if self.winfo_exists():
                self.lift()
                self.focus_force()
        except tk.TclError:
            pass
        self._yenile()


# --------------------------------------------------------------------------
# Yönetici paneli
# --------------------------------------------------------------------------

def _tarih_secici_olustur(parent, baslangic_tarih=None):
    """tkcalendar mevcutsa üzerine tıklanınca modern, uygulamanın renk
    paletiyle uyumlu bir takvim açılan DateEntry, değilse (paket kurulu
    değilse) elle yazılan düz bir metin kutusu döndürür."""
    baslangic_tarih = baslangic_tarih or date.today()
    if TKCALENDAR_MEVCUT:
        ortak_ayarlar = dict(
            width=11, date_pattern="yyyy-mm-dd",
            font=("Segoe UI", 10),
            background=RENK_VURGU, foreground="white", borderwidth=1,
            headersbackground=RENK_VURGU, headersforeground="white",
            selectbackground=RENK_VURGU, selectforeground="white",
            normalbackground="white", normalforeground="#1f2937",
            weekendbackground="#f4f5f7", weekendforeground="#1f2937",
            othermonthforeground="#c3c7cf", othermonthbackground="#fafafa",
            othermonthweforeground="#c3c7cf", othermonthwebackground="#fafafa",
            showweeknumbers=False,
        )
        try:
            # Ay/gün adlarını Türkçe göstermek için; bilgisayarda Türkçe
            # yerel ayarı bulunamazsa (locale hatası) sessizce varsayılan
            # (İngilizce) yerel ayara düşer, uygulama çökmez.
            widget = DateEntry(parent, locale="tr_TR", **ortak_ayarlar)
        except Exception:
            widget = DateEntry(parent, **ortak_ayarlar)
        widget.set_date(baslangic_tarih)
    else:
        widget = tk.Entry(
            parent, width=12, font=("Segoe UI", 10), relief="flat",
            highlightthickness=1, highlightbackground="#d8dbe0", highlightcolor=RENK_VURGU,
        )
        widget.insert(0, baslangic_tarih.strftime("%Y-%m-%d"))
    return widget


def _tarih_secici_ayarla(widget, tarih_obj):
    if TKCALENDAR_MEVCUT and hasattr(widget, "set_date"):
        widget.set_date(tarih_obj)
    else:
        widget.delete(0, "end")
        widget.insert(0, tarih_obj.strftime("%Y-%m-%d"))


class YoneticiPaneli(tk.Toplevel):
    def __init__(self, master, db: VeriTabani, salt_okunur=False):
        super().__init__(master)
        self.db = db
        self.salt_okunur = salt_okunur
        self.title("Yoklama Paneli" + (" (İzleyici)" if salt_okunur else " - Yönetici"))
        # Tablodaki kolonların toplam genişliği (~1220px) + kaydırma çubuğu +
        # kenar boşlukları 1180px'e sığmıyordu, bu yüzden sağ taraftaki
        # kolonlar (Yok/Toplam vb.) kırpılıyordu; pencere hem izleyici hem
        # yönetici için ekran genişliğine göre olabildiğince geniş açılır.
        # ÖNEMLİ: alt sınır (eski "max(1370, ...)" / "max(600, ...)") ASLA
        # ekranın gerçek genişliğinden/yüksekliğinden büyük olmamalı - aksi
        # halde (ör. 1366x768 gibi çok yaygın bir ekranda) pencere ekranın
        # dışına taşar ve sağ/alt kısım GÖRÜNMEZ olur; bu önceki hatanın
        # asıl sebebiydi. Bu yüzden önce ekrana göre üst sınır konur, sonra
        # küçük ve güvenli bir alt sınırla makul bir minimum sağlanır; tablo
        # zaten yatay/dikey kaydırma çubuklarına sahip olduğu için pencere
        # tüm sütunları aynı anda göstermese bile kaydırarak ulaşılabilir.
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        pencere_genisligi = min(1440, ekran_genisligi - 40)
        pencere_genisligi = max(900, pencere_genisligi)
        pencere_yuksekligi = min(760, ekran_yuksekligi - 80)
        pencere_yuksekligi = max(500, pencere_yuksekligi)
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.minsize(900, 480)
        self.configure(bg=RENK_ARKAPLAN)
        _baslik_cubugu_rengini_ayarla(self)
        # Ana sayfanın (Yönetici Ana Paneli) arkasında kalmaması için:
        # sadece "transient" (+ lift/focus_force) yeterli olmuyordu -
        # kullanıcı ana sayfaya TIKLADIĞINDA çoğu pencere yöneticisinde
        # (özellikle Windows'ta) ana sayfa yine de öne gelebiliyordu, çünkü
        # transient yalnızca bir ipucu, kesin bir garanti değil. Kesin
        # çözüm: uygulamadaki diğer tüm alt pencerelerde (Kurs Ekle,
        # Kursiyer Ekle, Düzenle vb.) zaten kullanılan "grab_set" - bu,
        # bu pencere açıkken ana sayfanın TIKLANAMAZ/etkileşimsiz olmasını
        # sağlar, dolayısıyla ana sayfaya tıklayıp onu öne getirmek hiç
        # mümkün olmaz.
        self.transient(master)
        self.grab_set()
        _grab_guvenli_baglat(self)
        self.lift()
        self.focus_force()

        self.mevcut_satirlar = []
        self.mevcut_dosya_etiketi = date.today().strftime("%Y-%m-%d")

        BUTON_KOYU = dict(
            bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK, activebackground=RENK_VURGU, activeforeground="white",
            font=("Segoe UI", 9, "bold"), relief="flat", bd=0, cursor="hand2",
        )
        BUTON_VURGU = dict(
            bg=RENK_VURGU, fg="white", activebackground="#1d4ed8", activeforeground="white",
            font=("Segoe UI", 9, "bold"), relief="flat", bd=0, cursor="hand2",
        )
        BUTON_YESIL = dict(
            bg=RENK_YESIL, fg="white", activebackground="#15803d", activeforeground="white",
            font=("Segoe UI", 9, "bold"), relief="flat", bd=0, cursor="hand2",
        )
        ETIKET_STIL = dict(bg=RENK_ARKAPLAN, fg="#374151", font=("Segoe UI", 10))

        ust = tk.Frame(self, bg=RENK_ARKAPLAN)
        ust.pack(fill="x", padx=14, pady=(14, 6))

        tk.Button(ust, text="◀ Önceki Gün", command=self._onceki_gun, **BUTON_KOYU).pack(side="left", ipadx=6, ipady=4)
        tk.Label(ust, text="Tarih:" if TKCALENDAR_MEVCUT else "Tarih (YYYY-AA-GG):", **ETIKET_STIL).pack(side="left", padx=(12, 6))
        self.tarih_entry = _tarih_secici_olustur(ust)
        self.tarih_entry.pack(side="left", padx=(0, 10))
        tk.Button(ust, text="Göster", command=self._yenile, **BUTON_VURGU).pack(side="left", ipadx=8, ipady=4)
        tk.Button(ust, text="Bugün", command=self._bugune_don, **BUTON_KOYU).pack(side="left", padx=(6, 0), ipadx=6, ipady=4)
        tk.Button(ust, text="Sonraki Gün ▶", command=self._sonraki_gun, **BUTON_KOYU).pack(side="left", padx=(6, 0), ipadx=6, ipady=4)

        tk.Button(
            ust, text="Excel'e Aktar", command=self._disa_aktar, **BUTON_YESIL
        ).pack(side="right", ipadx=8, ipady=4)

        aralik = tk.Frame(self, bg=RENK_ARKAPLAN)
        aralik.pack(fill="x", padx=14, pady=(0, 10))
        tk.Label(aralik, text="Tarih aralığı:", **ETIKET_STIL).pack(side="left")
        tk.Label(aralik, text="Başlangıç", **ETIKET_STIL).pack(side="left", padx=(10, 6))
        self.aralik_baslangic_entry = _tarih_secici_olustur(aralik)
        self.aralik_baslangic_entry.pack(side="left")
        tk.Label(aralik, text="Bitiş", **ETIKET_STIL).pack(side="left", padx=(10, 6))
        self.aralik_bitis_entry = _tarih_secici_olustur(aralik)
        self.aralik_bitis_entry.pack(side="left")
        tk.Button(
            aralik, text="Aralığı Göster", command=self._araligi_goster, **BUTON_VURGU
        ).pack(side="left", padx=(10, 0), ipadx=8, ipady=4)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=(0, 8))

        ozet = tk.Frame(self, bg=RENK_ARKAPLAN)
        ozet.pack(fill="x", padx=14)
        self.ozet_label = tk.Label(ozet, text="", bg=RENK_ARKAPLAN, fg="#111827", font=("Segoe UI", 11, "bold"))
        self.ozet_label.pack(anchor="w", pady=(0, 8))

        agac_cercevesi = tk.Frame(self, bg=RENK_ARKAPLAN)
        agac_cercevesi.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        kolonlar = ("tarih", "saat", "grup", "derslik", "ders_adi", "durum", "giris_zamani", "toplam", "var", "yok")
        self.tree = ttk.Treeview(agac_cercevesi, columns=kolonlar, show="headings", height=18)
        basliklar = {
            "tarih": "Tarih", "saat": "Saat", "grup": "Grup", "derslik": "Derslik", "ders_adi": "Kurs",
            "durum": "Durum", "giris_zamani": "Giriş Zamanı", "toplam": "Toplam Kişi",
            "var": "Hazır (Var)", "yok": "Yok (Devamsız)",
        }
        # "Saat" kolonu artık "1. Saat  08:00 - 08:50" gibi periyot bilgisi de
        # içerdiği için genişletildi.
        genislikler = {
            "tarih": 90, "saat": 170, "grup": 110, "derslik": 130, "ders_adi": 140,
            "durum": 200, "giris_zamani": 150, "toplam": 90, "var": 90, "yok": 120,
        }
        self.kolon_basliklari = basliklar
        self.siralama_sutunu = None
        self.siralama_tersten = False
        for k in kolonlar:
            self.tree.heading(k, text=basliklar[k], command=lambda col=k: self._sutuna_gore_sirala(col))
            self.tree.column(k, width=genislikler[k], anchor="center" if k not in ("grup", "derslik", "ders_adi") else "w")
        dikey_kaydirma = ttk.Scrollbar(agac_cercevesi, orient="vertical", command=self.tree.yview)
        yatay_kaydirma = ttk.Scrollbar(agac_cercevesi, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=dikey_kaydirma.set, xscrollcommand=yatay_kaydirma.set)
        agac_cercevesi.grid_rowconfigure(0, weight=1)
        agac_cercevesi.grid_columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew")
        dikey_kaydirma.grid(row=0, column=1, sticky="ns")
        yatay_kaydirma.grid(row=1, column=0, sticky="ew")

        self._yenile()

    def _bugune_don(self):
        _tarih_secici_ayarla(self.tarih_entry, date.today())
        self._yenile()

    def _gun_kaydir(self, delta_gun):
        tarih_str = self.tarih_entry.get().strip()
        try:
            mevcut = datetime.strptime(tarih_str, "%Y-%m-%d").date()
        except ValueError:
            mevcut = date.today()
        yeni = mevcut + timedelta(days=delta_gun)
        _tarih_secici_ayarla(self.tarih_entry, yeni)
        self._yenile()

    def _onceki_gun(self):
        self._gun_kaydir(-1)

    def _sonraki_gun(self):
        self._gun_kaydir(1)

    def _tabloyu_doldur(self, satirlar, baslik_etiketi):
        self.mevcut_baslik_etiketi = baslik_etiketi
        # Bir sütun başlığına daha önce tıklanarak sıralama seçildiyse,
        # yeni gelen veri (gün değiştirme, "Göster" vb.) de aynı sıralamayla
        # gösterilmeye devam eder.
        if getattr(self, "siralama_sutunu", None):
            satirlar = sorted(
                satirlar, key=self._siralama_anahtari(self.siralama_sutunu),
                reverse=self.siralama_tersten,
            )

        for i in self.tree.get_children():
            self.tree.delete(i)

        self.mevcut_satirlar = satirlar
        self._basliklari_guncelle()
        # NOT: Excel dışa aktarımıyla (bkz. _xlsx_olustur) AYNI mantık.
        # Bir kursun aynı gün içinde birden çok periyodu (ör. 6 saat)
        # olabiliyor ve her periyot burada AYRI bir satır olarak geliyor,
        # ama hepsi AYNI kursiyer listesine (aynı kişilere) ait. Bu yüzden
        # "Toplam kişi" her SATIRDA değil, her KURS ilk göründüğünde BİR
        # KERE sayılır - aksi halde ör. 6 periyotlu bir kursun kişi sayısı
        # toplamda 6 katına çıkmış gibi görünürdü. "Hazır"/"Devamsız"
        # toplamları da (Excel'deki gibi) BİLEREK gösterilmiyor - bunlar
        # farklı periyotlardaki sayıların toplamı olduğu için (özellikle
        # öğretmenin yoklama girmediği periyotlarla karışınca) yanıltıcı
        # bir "genel toplam" ortaya çıkarıyordu.
        gorulen_ders_idleri = set()
        toplam_kisi = 0
        for s in satirlar:
            if s["ders_id"] not in gorulen_ders_idleri:
                gorulen_ders_idleri.add(s["ders_id"])
                toplam_kisi += s["toplam"]
        self.ozet_label.config(
            text=f"{baslik_etiketi}  |  Toplam kişi: {toplam_kisi}"
        )

        for s in satirlar:
            # Öğretmen yoklamayı KENDİSİ hiç girmediyse - ister süre henüz
            # dolmadığı için "bekliyor" olsun, ister süre dolup sistem
            # otomatik olarak "öğretmen girmedi" diye işaretlemiş olsun -
            # hem Hazır (Var) hem de Yok (Devamsız) sütununda "0"/sayı
            # göstermek "gerçekten yoklama alındı" gibi yanlış bir izlenim
            # veriyordu; bu iki durumda da her ikisi de "-" gösterilir.
            # (Excel dışa aktarımıyla aynı mantık - bkz. _xlsx_olustur.)
            henuz_alinmadi = (s["durum"] == "bekliyor")
            ogretmen_girmedi_mi = (s["durum"] == "ogretmen_girmedi")
            yoklama_alinmadi = henuz_alinmadi or ogretmen_girmedi_mi
            var_gosterim = "-" if yoklama_alinmadi else s["var"]
            yok_gosterim = "-" if yoklama_alinmadi else s["yok"]
            self.tree.insert(
                "", "end",
                values=(
                    s["tarih"], s["saat"], s["grup"], s["derslik"], s["ders_adi"],
                    DURUM_METIN.get(s["durum"], s["durum"]), s["giris_zamani"],
                    s["toplam"], var_gosterim, yok_gosterim,
                ),
            )

    def _sutuna_gore_sirala(self, col):
        """Bir sütun başlığına tıklandığında, tabloyu o sütuna göre sıralar;
        aynı başlığa tekrar tıklanırsa artan/azalan sırayı tersine çevirir
        (küçükten büyüğe / büyükten küçüğe)."""
        if self.siralama_sutunu == col:
            self.siralama_tersten = not self.siralama_tersten
        else:
            self.siralama_sutunu = col
            self.siralama_tersten = False
        self._tabloyu_doldur(self.mevcut_satirlar, self.mevcut_baslik_etiketi)

    def _siralama_anahtari(self, col):
        # "Saat" sütunu "1. Saat  08:00 - 08:50" gibi metin olarak
        # göründüğü için METİN olarak sıralanırsa "10. Saat" "2. Saat"tan
        # önce gelir; bunun yerine gerçek periyot numarasına (sira) göre
        # sıralanır. "Toplam/Hazır/Yok" her zaman sayı olarak saklanır
        # ("-" yalnızca EKRANDA gösterilen bir metindir, bkz. _tabloyu_doldur),
        # bu yüzden doğrudan sayısal olarak sıralanabilir.
        if col == "saat":
            return lambda s: s.get("sira", 0)
        if col == "durum":
            return lambda s: DURUM_METIN.get(s["durum"], s["durum"])
        if col in ("toplam", "var", "yok"):
            return lambda s: s[col]
        return lambda s: (s.get(col) or "")

    def _basliklari_guncelle(self):
        """Aktif sıralama sütununun başlığına küçük bir ok ekler (▲ artan /
        ▼ azalan), diğer başlıklar düz metin olarak kalır."""
        for k, metin in self.kolon_basliklari.items():
            if k == self.siralama_sutunu:
                ok = " ▼" if self.siralama_tersten else " ▲"
                self.tree.heading(k, text=metin + ok)
            else:
                self.tree.heading(k, text=metin)

    def _yenile(self):
        self.db.otomatik_devamsizlik_kontrolu()
        tarih_str = self.tarih_entry.get().strip()
        try:
            datetime.strptime(tarih_str, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror(APP_ADI, "Tarihi YYYY-AA-GG formatında girin.", parent=self)
            return

        satirlar = self.db.gunluk_ozet(tarih_str)
        self._tabloyu_doldur(satirlar, tarih_str)
        self.mevcut_dosya_etiketi = tarih_str

    def _araligi_goster(self):
        self.db.otomatik_devamsizlik_kontrolu()
        baslangic_str = self.aralik_baslangic_entry.get().strip()
        bitis_str = self.aralik_bitis_entry.get().strip()
        try:
            b_tarih = datetime.strptime(baslangic_str, "%Y-%m-%d").date()
            s_tarih = datetime.strptime(bitis_str, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror(APP_ADI, "Tarihleri YYYY-AA-GG formatında girin.", parent=self)
            return
        if s_tarih < b_tarih:
            messagebox.showerror(APP_ADI, "Bitiş tarihi, başlangıç tarihinden önce olamaz.", parent=self)
            return
        if (s_tarih - b_tarih).days > 400:
            messagebox.showerror(APP_ADI, "En fazla 400 günlük bir aralık seçebilirsiniz.", parent=self)
            return

        satirlar = []
        gun = b_tarih
        while gun <= s_tarih:
            satirlar.extend(self.db.gunluk_ozet(gun.strftime("%Y-%m-%d")))
            gun += timedelta(days=1)

        self._tabloyu_doldur(satirlar, f"{baslangic_str}  →  {bitis_str}")
        self.mevcut_dosya_etiketi = f"{baslangic_str}_{bitis_str}"

    def _disa_aktar(self):
        if not OPENPYXL_MEVCUT:
            messagebox.showerror(
                APP_ADI,
                "Excel'e aktarma için gereken 'openpyxl' kütüphanesi bulunamadı.\n\n"
                "Eğer bu programı .py dosyası olarak çalıştırıyorsanız:\n"
                "  pip install openpyxl\n\n"
                "komutunu çalıştırın. EXE olarak kullanıyorsanız, build_exe.bat "
                "ile yeniden derleyin (bu betik openpyxl'i otomatik kurar).",
                parent=self,
            )
            return

        satirlar = self.mevcut_satirlar
        if not satirlar:
            messagebox.showinfo(APP_ADI, "Aktarılacak kayıt yok.", parent=self)
            return
        dosya_yolu = filedialog.asksaveasfilename(
            parent=self,
            title="Excel olarak kaydet",
            defaultextension=".xlsx",
            initialfile=f"yoklama_{self.mevcut_dosya_etiketi}.xlsx",
            filetypes=[("Excel dosyası", "*.xlsx")],
        )
        if not dosya_yolu:
            return
        try:
            self._xlsx_olustur(dosya_yolu, satirlar)
            messagebox.showinfo(APP_ADI, f"Dışa aktarıldı:\n{dosya_yolu}", parent=self)
        except Exception as e:
            messagebox.showerror(APP_ADI, f"Dışa aktarma başarısız:\n{e}", parent=self)

    def _xlsx_olustur(self, dosya_yolu, ozet_satirlar):
        wb = Workbook()

        baslik_font = Font(bold=True, color="FFFFFF")
        baslik_dolgu = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        kirmizi_font = Font(color="B91C1C")
        kalin_font = Font(bold=True)
        kurs_baslik_font = Font(bold=True, color="1F2937")
        kurs_baslik_dolgu = PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid")

        # Satırlar veritabanından saate göre (tüm kursların aynı periyodu
        # bir arada) geldiği için, dışa aktarımda okunması çok karışık
        # görünüyordu. Bunun yerine kurs kurs AYRI bloklar halinde, her
        # kursun kendi içinde tarih/saat sırasına göre yazdırılır.
        ozet_satirlar = sorted(
            ozet_satirlar,
            key=lambda s: (
                s["grup"], s["ders_adi"], s["ders_id"], s["tarih"], s.get("sira", 0),
            ),
        )

        # ---- 1. sayfa: Özet ----
        ws1 = wb.active
        ws1.title = "Özet"
        ozet_basliklar = [
            "Tarih", "Saat", "Grup", "Derslik", "Kurs", "Kurs Başlangıç", "Kurs Bitiş",
            "Durum", "Giriş Zamanı (son güncelleme)",
            "Toplam Kişi", "Hazır (Var)", "Yok (Devamsız)",
            "Devamsız Kursiyerler (Kod - Ad Soyad - Mazeret)",
        ]
        ws1.append(ozet_basliklar)
        for col_idx in range(1, len(ozet_basliklar) + 1):
            hucre = ws1.cell(row=1, column=col_idx)
            hucre.font = baslik_font
            hucre.fill = baslik_dolgu
            hucre.alignment = Alignment(horizontal="center")

        # Her oturumun kayıtlarını bir kere çekip hem bu sayfada hem de
        # Detay sayfasında kullanmak için önbelleğe alıyoruz.
        oturum_kayitlari = {}
        for satir in ozet_satirlar:
            if satir["oturum_id"]:
                oturum_kayitlari[satir["oturum_id"]] = self.db.yoklama_kayitlarini_getir(satir["oturum_id"])

        # NOT: Bir kursun aynı gün içinde birden çok periyodu (ör. 6 saat)
        # olabiliyor ve her periyot gunluk_ozet()'te AYRI bir satır olarak
        # geliyor - ama hepsi AYNI kursiyer listesine (aynı kişilere) ait.
        # Bu yüzden "Toplam Kişi" toplamını her SATIRDA değil, her KURS
        # ilk göründüğünde BİR KERE sayıyoruz - aksi halde ör. 6 periyotlu
        # bir kursun kişi sayısı toplamda 6 katına çıkmış gibi görünürdü.
        toplam_kisi_genel = 0
        onceki_ders_id_ozet = None
        ilk_kurs_blogu_mu = True
        for satir in ozet_satirlar:
            if satir["ders_id"] != onceki_ders_id_ozet:
                onceki_ders_id_ozet = satir["ders_id"]
                toplam_kisi_genel += satir["toplam"]
                if not ilk_kurs_blogu_mu:
                    # Her derslik/kurs bloğu birbirinden ayırt edilsin diye
                    # aralarına 2 boş satır bırakılıyor (ilk bloktan önce
                    # gerek yok, o zaten sayfanın başlığının hemen altında).
                    ws1.append([])
                    ws1.append([])
                ilk_kurs_blogu_mu = False
                ws1.append([
                    f"KURS: {satir['ders_adi']}   |   Grup: {satir['grup']}   |   Derslik: {satir['derslik']}"
                ] + [""] * (len(ozet_basliklar) - 1))
                kurs_baslik_satir_no = ws1.max_row
                # NOT: Bu satır bilerek BİRLEŞTİRİLMİYOR (merge_cells
                # kullanılmıyor) - Excel'in kendi Sırala/Filtrele
                # özelliği, aralarında farklı boyutlarda birleştirilmiş
                # hücre bulunan bir aralıkta hata veriyor. Bunun yerine her
                # hücre ayrı ayrı aynı gri/kalın görünümle boyanır; böylece
                # kurs blokları görsel olarak yine ayrılır AMA sayfa tam
                # olarak sıralanabilir/filtrelenebilir kalır.
                for c in range(1, len(ozet_basliklar) + 1):
                    kurs_baslik_hucre = ws1.cell(row=kurs_baslik_satir_no, column=c)
                    kurs_baslik_hucre.font = kurs_baslik_font
                    kurs_baslik_hucre.fill = kurs_baslik_dolgu
                    kurs_baslik_hucre.alignment = Alignment(horizontal="left", vertical="center")

            # Öğretmen yoklamayı KENDİSİ hiç girmediyse - "bekliyor" (süre
            # henüz dolmadı) ya da "ogretmen_girmedi" (süre doldu, sistem
            # otomatik devamsız işaretledi) - hem Hazır (Var) hem de Yok
            # (Devamsız) sütununda "0"/sayı göstermek "gerçekten yoklama
            # alındı" gibi yanlış bir izlenim veriyordu; bu iki durumda da
            # her ikisi de "-" gösterilir.
            henuz_alinmadi = (satir["durum"] == "bekliyor")
            ogretmen_girmedi_mi = (satir["durum"] == "ogretmen_girmedi")
            yoklama_alinmadi = henuz_alinmadi or ogretmen_girmedi_mi
            var_gosterim = "-" if yoklama_alinmadi else satir["var"]
            yok_gosterim = "-" if yoklama_alinmadi else satir["yok"]
            # Öğretmen yoklamayı hiç girmediyse (yukarıdaki "-" durumu),
            # sistemin otomatik "herkes yok" işaretlemesi GERÇEK bir
            # devamsızlık kaydı DEĞİLDİR - bu yüzden "Devamsız Kursiyerler"
            # sütununda kimsenin kodu/adı YAZILMAZ, sütun boş kalır.
            # Sadece öğretmenin GERÇEKTEN kaydettiği (tamamlandi) bir
            # oturumdaki devamsızlar burada listelenir.
            if yoklama_alinmadi:
                devamsiz_metin = ""
                devamsizlar = []
            else:
                kayitlar = oturum_kayitlari.get(satir["oturum_id"], [])
                devamsizlar = [k for k in kayitlar if k["durum"] == "yok"]
                # Devamsız kursiyerleri alt alta, aralarında boşluk olacak
                # şekilde ayrı satırlara yazıyoruz (çok kişi olsa da düzgün
                # görünmesi için hücreyi kaydırmalı yapıp satır yüksekliğini
                # kişi sayısına göre otomatik ayarlıyoruz).
                devamsiz_metin = "\n".join(_devamsiz_kursiyer_satiri(k) for k in devamsizlar)
            kurs_baslangic_gosterim = satir.get("kurs_baslangic") or "Sınırsız"
            kurs_bitis_gosterim = satir.get("kurs_bitis") or "Sınırsız"
            ws1.append([
                satir["tarih"], satir["saat"], satir["grup"], satir["derslik"], satir["ders_adi"],
                kurs_baslangic_gosterim, kurs_bitis_gosterim,
                DURUM_METIN.get(satir["durum"], satir["durum"]),
                satir["giris_zamani"], satir["toplam"], var_gosterim, yok_gosterim,
                devamsiz_metin,
            ])
            satir_no = ws1.max_row
            if (not yoklama_alinmadi) and satir["yok"] > 0:
                ws1.cell(row=satir_no, column=12).font = kirmizi_font
                ws1.cell(row=satir_no, column=13).font = kirmizi_font

            devamsiz_hucre = ws1.cell(row=satir_no, column=13)
            devamsiz_hucre.alignment = Alignment(
                wrap_text=True, vertical="top", horizontal="left"
            )
            # "Toplam Kişi" / "Hazır (Var)" / "Yok (Devamsız)" sütunları
            # (10, 11, 12) - hem sayı hem de "-" değeri olsun - hücre içinde
            # ORTALANARAK gösterilir; diğer sütunlar varsayılan hizalamada
            # kalır.
            ORTALANACAK_SUTUNLAR = (10, 11, 12)
            for c in range(1, 13):
                if c in ORTALANACAK_SUTUNLAR:
                    ws1.cell(row=satir_no, column=c).alignment = Alignment(
                        vertical="top", horizontal="center"
                    )
                else:
                    ws1.cell(row=satir_no, column=c).alignment = Alignment(vertical="top")
            satir_sayisi = max(1, len(devamsizlar))
            if satir_sayisi > 1:
                # Excel'in izin verdiği azami satır yüksekliği ~409 puandır.
                ws1.row_dimensions[satir_no].height = min(15 * satir_sayisi + 4, 400)

        # "Giriş Zamanı" sütununu METİN olarak işaretle (Excel'in tarih
        # sanip ##### göstermesini engellemek için) ve sütunları genişlet.
        giris_zamani_kolon = 9
        for satir_no in range(2, ws1.max_row + 1):
            ws1.cell(row=satir_no, column=giris_zamani_kolon).number_format = "@"

        # "Toplam" satırını eklemeden ÖNCE son gerçek veri satırını
        # işaretliyoruz; aşağıdaki auto_filter (Excel'in kendi sırala/
        # filtrele ok işaretleri) sadece gerçek veriyi kapsasın, alttaki
        # boş satır + TOPLAM özet satırı bu aralığın dışında kalsın.
        ozet_veri_son_satir = ws1.max_row

        ws1.append([])
        # Hazır (Var) / Yok (Devamsız) toplamları BİLEREK burada
        # gösterilmiyor: bunlar farklı kurs/periyotlardaki sayıların
        # toplamı olduğu için (özellikle öğretmenin yoklama girmediği,
        # "-" gösterilen periyotlarla karışınca) yanıltıcı bir "genel
        # toplam" ortaya çıkarıyordu. Sadece (her kurs bir kez sayılan,
        # bkz. yukarıdaki not) toplam kişi sayısı gösterilir.
        toplam_satir = [
            "", "", "", "", "", "", "", "", "TOPLAM KİŞİ",
            toplam_kisi_genel, "", "", "",
        ]
        ws1.append(toplam_satir)
        for c in range(9, 14):
            ws1.cell(row=ws1.max_row, column=c).font = kalin_font
        # "Toplam Kişi" sayısı varsayılanda sağa yaslı görünüyordu (Excel
        # sayıları sağa yaslar); üstteki "Toplam Kişi" başlığıyla ve diğer
        # satırlardaki sayı sütunlarıyla tutarlı olsun diye ortalanıyor.
        ws1.cell(row=ws1.max_row, column=10).alignment = Alignment(horizontal="center")

        genislikler_ozet = [14, 14, 18, 24, 24, 15, 15, 28, 20, 13, 13, 16, 34]
        for i, genislik in enumerate(genislikler_ozet, start=1):
            ws1.column_dimensions[get_column_letter(i)].width = genislik
        # Excel'in kendi başlık ok işaretleri (filtrele + sırala A-Z/Z-A);
        # kurs başlık satırları artık birleştirilmiş hücre OLMADIĞI için bu
        # sayfa tam olarak sıralanabilir/filtrelenebilir.
        ws1.auto_filter.ref = f"A1:{get_column_letter(len(ozet_basliklar))}{ozet_veri_son_satir}"
        ws1.freeze_panes = "A2"

        # ---- 2. sayfa: Detay (kursiyer bazlı - kim gelmiş/gelmemiş) ----
        ws2 = wb.create_sheet("Detay (Kursiyer)")
        detay_basliklar = [
            "Tarih", "Saat", "Grup", "Derslik", "Kurs", "Kurs Başlangıç", "Kurs Bitiş",
            "Kursiyer Kodu", "Ad Soyad", "Durum", "Mazeret",
        ]
        ws2.append(detay_basliklar)
        for col_idx in range(1, len(detay_basliklar) + 1):
            hucre = ws2.cell(row=1, column=col_idx)
            hucre.font = baslik_font
            hucre.fill = baslik_dolgu
            hucre.alignment = Alignment(horizontal="center")

        onceki_ders_id_detay = None
        for satir in ozet_satirlar:
            if not satir["oturum_id"]:
                continue
            if satir["durum"] == "ogretmen_girmedi":
                # bkz. Özet sayfasındaki aynı not: öğretmen yoklamayı hiç
                # girmediyse sistemin otomatik "herkes yok" işaretlemesi
                # gerçek bir kayıt değildir - bu kursiyerler burada da
                # tek tek listelenmez.
                continue
            if satir["ders_id"] != onceki_ders_id_detay:
                onceki_ders_id_detay = satir["ders_id"]
                ws2.append([
                    f"KURS: {satir['ders_adi']}   |   Grup: {satir['grup']}   |   Derslik: {satir['derslik']}"
                ] + [""] * (len(detay_basliklar) - 1))
                kurs_baslik_satir_no2 = ws2.max_row
                # bkz. Özet sayfasındaki aynı not: hücreler bilerek
                # birleştirilmiyor, sayfa tam sıralanabilir/filtrelenebilir
                # kalsın diye.
                for c in range(1, len(detay_basliklar) + 1):
                    kurs_baslik_hucre2 = ws2.cell(row=kurs_baslik_satir_no2, column=c)
                    kurs_baslik_hucre2.font = kurs_baslik_font
                    kurs_baslik_hucre2.fill = kurs_baslik_dolgu
                    kurs_baslik_hucre2.alignment = Alignment(horizontal="left", vertical="center")

            kayitlar = oturum_kayitlari.get(satir["oturum_id"], [])
            kurs_baslangic_gosterim2 = satir.get("kurs_baslangic") or "Sınırsız"
            kurs_bitis_gosterim2 = satir.get("kurs_bitis") or "Sınırsız"
            for k in kayitlar:
                durum_metni = "Var" if k["durum"] == "var" else "Yok"
                ws2.append([
                    satir["tarih"], satir["saat"], satir["grup"], satir["derslik"], satir["ders_adi"],
                    kurs_baslangic_gosterim2, kurs_bitis_gosterim2,
                    k["kod"], k["ad_soyad"], durum_metni, k["mazeret_metin"] or "",
                ])
                if durum_metni == "Yok":
                    for c in range(1, len(detay_basliklar) + 1):
                        ws2.cell(row=ws2.max_row, column=c).font = kirmizi_font

        genislikler_detay = [14, 14, 18, 24, 24, 15, 15, 16, 28, 10, 20]
        for i, genislik in enumerate(genislikler_detay, start=1):
            ws2.column_dimensions[get_column_letter(i)].width = genislik
        ws2.freeze_panes = "A2"
        if ws2.max_row > 1:
            ws2.auto_filter.ref = f"A1:{get_column_letter(len(detay_basliklar))}{ws2.max_row}"

        wb.save(dosya_yolu)


class SifreYonetimDialog(tk.Toplevel):
    def __init__(self, master, db: VeriTabani):
        super().__init__(master)
        self.db = db
        self.title("Şifreleri Yönet")
        self.geometry("380x480")
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)

        tk.Label(self, text="Şifreleri Yönet", font=("Segoe UI", 13, "bold")).pack(pady=(14, 10))

        frm = tk.Frame(self, padx=16)
        frm.pack(fill="both", expand=True)

        self.grup_girisler = {}
        for grup in self.db.gruplari_getir():
            satir = tk.Frame(frm)
            satir.pack(fill="x", pady=4)
            tk.Label(satir, text=grup["ad"], width=16, anchor="w").pack(side="left")
            girdi = tk.Entry(satir, show="*")
            girdi.pack(side="left", fill="x", expand=True)
            self.grup_girisler[grup["id"]] = girdi

        ttk.Separator(self).pack(fill="x", pady=10, padx=16)

        admin_frame = tk.Frame(frm)
        admin_frame.pack(fill="x", pady=4)
        tk.Label(admin_frame, text="Yönetici şifresi", width=16, anchor="w").pack(side="left")
        self.admin_girdi = tk.Entry(admin_frame, show="*")
        self.admin_girdi.pack(side="left", fill="x", expand=True)

        izleyici_frame = tk.Frame(frm)
        izleyici_frame.pack(fill="x", pady=4)
        tk.Label(izleyici_frame, text="İzleyici şifresi", width=16, anchor="w").pack(side="left")
        self.izleyici_girdi = tk.Entry(izleyici_frame, show="*")
        self.izleyici_girdi.pack(side="left", fill="x", expand=True)

        tk.Label(
            self, text="Yalnızca değiştirmek istediğiniz alanları doldurun,\ndiğerlerini boş bırakabilirsiniz.",
            fg="#555555", justify="center",
        ).pack(pady=8)

        tk.Button(
            self, text="Kaydet", bg=RENK_VURGU, fg="white",
            font=("Segoe UI", 10, "bold"), command=self._kaydet
        ).pack(pady=10, ipady=6, fill="x", padx=16)

    def _kaydet(self):
        degisiklik = False
        for grup_id, girdi in self.grup_girisler.items():
            deger = girdi.get().strip()
            if deger:
                self.db.grup_sifre_ayarla(grup_id, deger)
                degisiklik = True
        admin_sifre = self.admin_girdi.get().strip()
        if admin_sifre:
            self.db.ayar_kaydet("admin_sifre_hash", sifre_hashle(admin_sifre))
            degisiklik = True
        izleyici_sifre = self.izleyici_girdi.get().strip()
        if izleyici_sifre:
            self.db.ayar_kaydet("izleyici_sifre_hash", sifre_hashle(izleyici_sifre))
            degisiklik = True
        if degisiklik:
            messagebox.showinfo(APP_ADI, "Şifreler güncellendi.", parent=self)
        self.destroy()


# --------------------------------------------------------------------------
# Ders / Kursiyer Düzenleme paneli (yanlış girilen kayıtları düzeltmek için)
# --------------------------------------------------------------------------

class DuzenlemePaneli(tk.Toplevel):
    """Yönetici, yanlış girilmiş bir dersi veya kursiyeri burada
    düzenleyebilir ya da silebilir."""

    def __init__(self, master, db: VeriTabani):
        super().__init__(master)
        self.db = db
        self.title("Kurs ve Kursiyer Düzenle")
        # İçerik (kurs bilgileri + haftalık program + kursiyer listesi +
        # alt kısımdaki "Yeni Kursiyer Ekle" / "Seçileni Sil" butonları)
        # sabit 960px yükseklikte, birçok ekranda (ör. 1366x768) sığmıyor
        # ve en alttaki butonlar hiç görünmüyordu; pencere artık ekran
        # yüksekliğine göre olabildiğince açılır (ama ekrandan taşmaz) VE
        # tüm içerik kaydırılabilir (fare tekerleği dahil) bir alana
        # alındı - böylece hangi ekran boyutunda olursa olsun tüm butonlara
        # aşağı kaydırarak ulaşılabilir.
        pencere_genisligi = 820
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        pencere_yuksekligi = min(960, ekran_yuksekligi - 80)
        pencere_yuksekligi = max(480, pencere_yuksekligi)
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.minsize(720, 420)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)

        self.grup_map = {g["ad"]: g["id"] for g in self.db.gruplari_getir()}
        # Derslikler kurumun tamamına ait ORTAK bir havuz olduğu için, bir
        # kursun dersliği Grup'tan bağımsız olarak tüm dersliklerin
        # listesinden değiştirilebilir - ANCAK başka bir kursa bitiş
        # tarihine kadar ayrılmış (dolu) derslikler bu listede gözükmez.
        # Burada henüz bir kurs SEÇİLMEDİĞİ için hariç tutulacak bir kurs
        # yok; bir kurs seçildiğinde liste _ders_secildi içinde, o kursun
        # kendi dersliğini de kapsayacak şekilde yeniden hesaplanır.
        self.derslik_map = {d["ad"]: d["id"] for d in self.db.musait_derslikleri_getir()}
        self.ders_map = {}
        self.secili_ders_id = None
        self.secili_kursiyer_id = None

        disaridaki = tk.Frame(self)
        disaridaki.pack(fill="both", expand=True)
        canvas = tk.Canvas(disaridaki, highlightthickness=0)
        kaydirma_cubugu = ttk.Scrollbar(disaridaki, orient="vertical", command=canvas.yview)
        icerik = tk.Frame(canvas)
        icerik.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas_pencere_id = canvas.create_window((0, 0), window=icerik, anchor="nw")
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfigure(canvas_pencere_id, width=e.width),
        )
        canvas.configure(yscrollcommand=kaydirma_cubugu.set)
        canvas.pack(side="left", fill="both", expand=True)
        kaydirma_cubugu.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(canvas)
        self.canvas = canvas

        ust = tk.Frame(icerik, padx=16, pady=12)
        ust.pack(fill="x")
        ust.grid_columnconfigure(1, weight=1)

        tk.Label(ust, text="Grup").grid(row=0, column=0, sticky="w", pady=4)
        self.grup_combo = ttk.Combobox(ust, values=list(self.grup_map.keys()), state="readonly")
        self.grup_combo.grid(row=0, column=1, sticky="ew", pady=4, padx=(6, 0))
        self.grup_combo.bind("<<ComboboxSelected>>", self._grup_secildi)

        tk.Label(ust, text="Kurs").grid(row=1, column=0, sticky="w", pady=4)
        self.ders_combo = ttk.Combobox(ust, values=[], state="readonly")
        self.ders_combo.grid(row=1, column=1, sticky="ew", pady=4, padx=(6, 0))
        self.ders_combo.bind("<<ComboboxSelected>>", self._ders_secildi)

        # ---- Ders bilgileri ----
        ders_grup = tk.LabelFrame(icerik, text="Kurs Bilgileri", padx=12, pady=10)
        ders_grup.pack(fill="x", padx=16, pady=(4, 8))
        ders_grup.grid_columnconfigure(1, weight=1)

        tk.Label(ders_grup, text="Kurs adı").grid(row=0, column=0, sticky="w", pady=4)
        self.ad_entry = tk.Entry(ders_grup)
        self.ad_entry.grid(row=0, column=1, sticky="ew", pady=4)

        tk.Label(ders_grup, text="Derslik").grid(row=1, column=0, sticky="w", pady=4)
        self.derslik_combo = ttk.Combobox(
            ders_grup, values=list(self.derslik_map.keys()), state="readonly"
        )
        self.derslik_combo.grid(row=1, column=1, sticky="ew", pady=4)

        tk.Label(
            ders_grup,
            text="Haftalık ders programı (Pazartesi-Perşembe aynı, Cuma ayrı)",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 4))
        self.program = HaftalikProgramCercevesi(ders_grup)
        self.program.grid(row=3, column=0, columnspan=2, sticky="ew")

        tk.Label(ders_grup, text="Başlangıç tarihi (GG.AA.YYYY)").grid(
            row=4, column=0, sticky="w", pady=4
        )
        self.baslangic_tarihi_entry = tk.Entry(ders_grup, width=14)
        self.baslangic_tarihi_entry.grid(row=4, column=1, sticky="w", pady=4)

        tk.Label(ders_grup, text="Bitiş tarihi (GG.AA.YYYY)").grid(
            row=5, column=0, sticky="w", pady=4
        )
        self.bitis_tarihi_entry = tk.Entry(ders_grup, width=14)
        self.bitis_tarihi_entry.grid(row=5, column=1, sticky="w", pady=4)

        tk.Label(ders_grup, text="Yoklama girme süresi (dk)").grid(row=6, column=0, sticky="w", pady=4)
        self.sure_spin = tk.Spinbox(ders_grup, from_=1, to=180, width=6)
        self.sure_spin.grid(row=6, column=1, sticky="w", pady=4)

        buton_cerceve = tk.Frame(ders_grup)
        buton_cerceve.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        tk.Button(
            buton_cerceve, text="Kurs Bilgilerini Güncelle", bg=RENK_VURGU, fg="white",
            command=self._ders_guncelle,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            buton_cerceve, text="Bu Kursu Sil", bg=RENK_KIRMIZI, fg="white",
            command=self._ders_sil,
        ).pack(side="left")

        # NOT: Bu kursa özel ara tatil ekleme/düzenleme artık burada değil -
        # Yönetici Panelindeki "Ara Tatil Yönetimi" ekranı (bkz.
        # AraTatilYonetimPaneli) aynı işi (tek bir kursu işaretleyip ekleme,
        # listeleme, düzenleme, silme dahil) tamamıyla ve daha fazlasıyla
        # (toplu ekleme + düzenleme) yapıyor; burada tekrarına gerek yoktu.

        # ---- Kursiyerler ----
        kursiyer_grup = tk.LabelFrame(icerik, text="Kursiyerler", padx=12, pady=10)
        kursiyer_grup.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        tree_cerceve = tk.Frame(kursiyer_grup)
        tree_cerceve.pack(fill="both", expand=True)
        self.kursiyer_tree = ttk.Treeview(
            tree_cerceve, columns=("kod", "ad_soyad"), show="headings", height=8
        )
        self.kursiyer_tree.heading("kod", text="Kod")
        self.kursiyer_tree.heading("ad_soyad", text="Ad Soyad")
        self.kursiyer_tree.column("kod", width=100, anchor="w")
        self.kursiyer_tree.column("ad_soyad", width=280, anchor="w")
        kaydirma = ttk.Scrollbar(tree_cerceve, orient="vertical", command=self.kursiyer_tree.yview)
        self.kursiyer_tree.configure(yscrollcommand=kaydirma.set)
        self.kursiyer_tree.pack(side="left", fill="both", expand=True)
        kaydirma.pack(side="right", fill="y")
        self.kursiyer_tree.bind("<<TreeviewSelect>>", self._kursiyer_secildi)

        duzenle_cerceve = tk.Frame(kursiyer_grup)
        duzenle_cerceve.pack(fill="x", pady=(10, 0))
        tk.Label(duzenle_cerceve, text="Kod").grid(row=0, column=0, sticky="w")
        self.kursiyer_kod_entry = tk.Entry(duzenle_cerceve, width=14)
        self.kursiyer_kod_entry.grid(row=0, column=1, sticky="w", padx=(4, 20))
        tk.Label(duzenle_cerceve, text="Ad Soyad").grid(row=0, column=2, sticky="w")
        self.kursiyer_ad_entry = tk.Entry(duzenle_cerceve, width=28)
        self.kursiyer_ad_entry.grid(row=0, column=3, sticky="ew", padx=(4, 0))
        duzenle_cerceve.grid_columnconfigure(3, weight=1)

        kursiyer_buton_cerceve = tk.Frame(kursiyer_grup)
        kursiyer_buton_cerceve.pack(fill="x", pady=(8, 0))
        tk.Button(
            kursiyer_buton_cerceve, text="Yeni Kursiyer Ekle", bg=RENK_YESIL, fg="white",
            command=self._kursiyer_ekle,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            kursiyer_buton_cerceve, text="Seçileni Güncelle", bg=RENK_VURGU, fg="white",
            command=self._kursiyer_guncelle,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            kursiyer_buton_cerceve, text="Seçileni Sil", bg=RENK_KIRMIZI, fg="white",
            command=self._kursiyer_sil,
        ).pack(side="left")

        # Enter/Leave tabanlı fare tekeri bağlamasına (yukarıda) ek olarak,
        # tüm içeriğe DOĞRUDAN da bağlanıyor - bazı ortamlarda Enter/Leave
        # olaylarının beklendiği gibi tetiklenmemesine karşı ek güvence.
        _agaca_fare_tekeri_ekle(icerik, self.canvas)

        self._ders_alanlarini_temizle()

    def _grup_secildi(self, event=None):
        grup_id = self.grup_map[self.grup_combo.get()]
        dersler = self.db.dersleri_getir(grup_id=grup_id)
        self.ders_map = {
            f"{d['ad']} ({d['derslik_ad']})": d["id"]
            for d in dersler
        }
        self.ders_combo["values"] = list(self.ders_map.keys())
        self.ders_combo.set("")
        self._ders_alanlarini_temizle()

    def _ders_secildi(self, event=None):
        ders_secim = self.ders_combo.get()
        if not ders_secim:
            return
        ders_id = self.ders_map[ders_secim]
        self.secili_ders_id = ders_id
        ders = self.db.ders_getir(ders_id)

        # Derslik listesini bu kursu hariç tutarak yeniden hesaplıyoruz;
        # böylece kursun ŞU AN kullandığı derslik "dolu" diye listeden
        # düşmüyor ve kullanıcı isterse aynı dersliği tekrar seçebiliyor,
        # ama başka kurslara ayrılmış (gerçekten dolu) derslikler yine de
        # gözükmüyor.
        derslikler = self.db.musait_derslikleri_getir(haric_ders_id=ders_id)
        self.derslik_map = {d["ad"]: d["id"] for d in derslikler}
        if ders["derslik_ad"] not in self.derslik_map:
            # Güvenlik önlemi: bir veri tutarsızlığı yüzünden kursun kendi
            # dersliği listede yoksa bile yine de eklenip gösterilsin.
            self.derslik_map[ders["derslik_ad"]] = ders["derslik_id"]
        self.derslik_combo["values"] = list(self.derslik_map.keys())

        self.ad_entry.delete(0, "end")
        self.ad_entry.insert(0, ders["ad"])
        self.derslik_combo.set(ders["derslik_ad"])

        self.program.saatleri_yukle(self.db.ders_saatleri_getir(ders_id))

        self.baslangic_tarihi_entry.delete(0, "end")
        self.baslangic_tarihi_entry.insert(0, iso_tarihi_goruntu_metnine_cevir(ders["baslangic_tarihi"]))
        self.bitis_tarihi_entry.delete(0, "end")
        self.bitis_tarihi_entry.insert(0, iso_tarihi_goruntu_metnine_cevir(ders["bitis_tarihi"]))
        self.sure_spin.delete(0, "end")
        self.sure_spin.insert(0, str(ders["devamsizlik_suresi_dk"] or DEVAMSIZLIK_SURESI_DK))

        self._kursiyerleri_yenile()

    def _ders_alanlarini_temizle(self):
        self.secili_ders_id = None
        self.ad_entry.delete(0, "end")
        # Artık hiçbir kurs seçili olmadığı için hariç tutulacak bir kurs
        # da yok - listeyi genel müsait derslik listesine döndürüyoruz.
        derslikler = self.db.musait_derslikleri_getir()
        self.derslik_map = {d["ad"]: d["id"] for d in derslikler}
        self.derslik_combo["values"] = list(self.derslik_map.keys())
        self.derslik_combo.set("")
        self.baslangic_tarihi_entry.delete(0, "end")
        self.bitis_tarihi_entry.delete(0, "end")
        self.sure_spin.delete(0, "end")
        self.sure_spin.insert(0, str(DEVAMSIZLIK_SURESI_DK))
        self._kursiyerleri_yenile()

    def _kursiyerleri_yenile(self):
        for i in self.kursiyer_tree.get_children():
            self.kursiyer_tree.delete(i)
        self.kursiyer_kod_entry.delete(0, "end")
        self.kursiyer_ad_entry.delete(0, "end")
        self.secili_kursiyer_id = None
        if not self.secili_ders_id:
            return
        for k in self.db.kursiyerleri_getir(self.secili_ders_id):
            self.kursiyer_tree.insert("", "end", iid=str(k["id"]), values=(k["kod"], k["ad_soyad"] or "-"))

    def _kursiyer_secildi(self, event=None):
        secim = self.kursiyer_tree.selection()
        if not secim:
            return
        kursiyer_id = int(secim[0])
        self.secili_kursiyer_id = kursiyer_id
        kursiyer = self.db.kursiyer_getir(kursiyer_id)
        if kursiyer:
            self.kursiyer_kod_entry.delete(0, "end")
            self.kursiyer_kod_entry.insert(0, kursiyer["kod"])
            self.kursiyer_ad_entry.delete(0, "end")
            self.kursiyer_ad_entry.insert(0, kursiyer["ad_soyad"])

    def _ders_guncelle(self):
        if not self.secili_ders_id:
            messagebox.showerror(APP_ADI, "Lütfen önce düzenlenecek kursu seçin.", parent=self)
            return
        ad = self.ad_entry.get().strip()
        derslik_ad = self.derslik_combo.get()

        if not (ad and derslik_ad):
            messagebox.showerror(APP_ADI, "Lütfen kurs adı ve derslik alanlarını doldurun.", parent=self)
            return

        try:
            saatler = self.program.saatleri_al()
        except ValueError as e:
            messagebox.showerror(APP_ADI, str(e), parent=self)
            return

        baslangic_tarihi = tarih_metnini_isoya_cevir(self.baslangic_tarihi_entry.get())
        bitis_tarihi = tarih_metnini_isoya_cevir(self.bitis_tarihi_entry.get())
        if baslangic_tarihi is None or bitis_tarihi is None:
            messagebox.showerror(
                APP_ADI,
                "Tarihleri GG.AA.YYYY biçiminde girin ya da sınırsız olması için boş bırakın.",
                parent=self,
            )
            return
        if baslangic_tarihi and bitis_tarihi and bitis_tarihi < baslangic_tarihi:
            messagebox.showerror(
                APP_ADI, "Bitiş tarihi, başlangıç tarihinden önce olamaz.", parent=self
            )
            return
        try:
            sure_dk = int(self.sure_spin.get())
            if not (1 <= sure_dk <= 180):
                raise ValueError
        except ValueError:
            messagebox.showerror(
                APP_ADI, "Yoklama girme süresi 1 ile 180 dakika arasında olmalıdır.", parent=self
            )
            return

        derslik_id = self.derslik_map[derslik_ad]
        self.db.ders_guncelle(
            self.secili_ders_id, ad, saatler, derslik_id=derslik_id,
            baslangic_tarihi=baslangic_tarihi, bitis_tarihi=bitis_tarihi,
            devamsizlik_suresi_dk=sure_dk,
        )
        messagebox.showinfo(APP_ADI, "Kurs güncellendi.", parent=self)

        secili_id = self.secili_ders_id
        self._grup_secildi()
        for etiket, did in self.ders_map.items():
            if did == secili_id:
                self.ders_combo.set(etiket)
                self._ders_secildi()
                break

    def _ders_sil(self):
        if not self.secili_ders_id:
            messagebox.showerror(APP_ADI, "Lütfen önce silinecek kursu seçin.", parent=self)
            return
        kursiyer_sayisi = self.db.ders_kursiyer_sayisi(self.secili_ders_id)
        uyari = "Bu kursu silmek istediğinizden emin misiniz? Bu işlem geri alınamaz."
        if kursiyer_sayisi > 0:
            uyari += (
                f"\n\nBu kursta kayıtlı {kursiyer_sayisi} kursiyer ve onlara ait "
                "tüm yoklama kayıtları da kalıcı olarak silinecektir."
            )
        if not messagebox.askyesno(APP_ADI, uyari, parent=self):
            return
        self.db.ders_sil(self.secili_ders_id)
        messagebox.showinfo(APP_ADI, "Kurs silindi.", parent=self)
        self._grup_secildi()

    def _kursiyer_ekle(self):
        if not self.secili_ders_id:
            messagebox.showerror(APP_ADI, "Lütfen önce kursiyer eklenecek kursu seçin.", parent=self)
            return
        kod = self.kursiyer_kod_entry.get().strip()
        ad = self.kursiyer_ad_entry.get().strip()
        if not kod:
            messagebox.showerror(APP_ADI, "Lütfen kod alanını doldurun (ad soyad opsiyoneldir).", parent=self)
            return
        if self.db.kod_kullanilmis_mi(self.secili_ders_id, kod):
            messagebox.showerror(
                APP_ADI, "Bu kod bu kursta başka bir kursiyer tarafından zaten kullanılıyor.",
                parent=self,
            )
            return
        self.db.kursiyer_ekle(self.secili_ders_id, kod, ad)
        messagebox.showinfo(APP_ADI, "Kursiyer eklendi.", parent=self)
        self._kursiyerleri_yenile()

    def _kursiyer_guncelle(self):
        if not self.secili_ders_id:
            messagebox.showerror(APP_ADI, "Lütfen önce bir kurs seçin.", parent=self)
            return
        if not self.secili_kursiyer_id:
            messagebox.showerror(
                APP_ADI, "Lütfen listeden düzenlenecek kursiyeri seçin.", parent=self
            )
            return
        kod = self.kursiyer_kod_entry.get().strip()
        ad = self.kursiyer_ad_entry.get().strip()
        if not kod:
            messagebox.showerror(APP_ADI, "Lütfen kod alanını doldurun (ad soyad opsiyoneldir).", parent=self)
            return
        if self.db.kod_kullanilmis_mi(self.secili_ders_id, kod, haric_kursiyer_id=self.secili_kursiyer_id):
            messagebox.showerror(
                APP_ADI, "Bu kod bu kursta başka bir kursiyer tarafından zaten kullanılıyor.",
                parent=self,
            )
            return
        self.db.kursiyer_guncelle(self.secili_kursiyer_id, kod, ad)
        messagebox.showinfo(APP_ADI, "Kursiyer güncellendi.", parent=self)
        self._kursiyerleri_yenile()

    def _kursiyer_sil(self):
        if not self.secili_kursiyer_id:
            messagebox.showerror(APP_ADI, "Lütfen listeden silinecek kursiyeri seçin.", parent=self)
            return
        kursiyer = self.db.kursiyer_getir(self.secili_kursiyer_id)
        ad_gosterim = kursiyer["ad_soyad"] if kursiyer else ""
        if not messagebox.askyesno(
            APP_ADI,
            f"\"{ad_gosterim}\" adlı kursiyeri ve yoklama kayıtlarını silmek "
            "istediğinizden emin misiniz? Bu işlem geri alınamaz.",
            parent=self,
        ):
            return
        self.db.kursiyer_sil(self.secili_kursiyer_id)
        messagebox.showinfo(APP_ADI, "Kursiyer silindi.", parent=self)
        self._kursiyerleri_yenile()


# --------------------------------------------------------------------------
# Derslik Yönetimi paneli
# --------------------------------------------------------------------------

class DerslikYonetimPaneli(tk.Toplevel):
    def __init__(self, master, db: VeriTabani, gruplar_degisti_geri_cagirma=None):
        super().__init__(master)
        self.db = db
        self.gruplar_degisti_geri_cagirma = gruplar_degisti_geri_cagirma
        self.title("Derslik Yönetimi")
        self.geometry("560x680")
        self.minsize(480, 560)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)

        self.grup_map = {g["ad"]: g["id"] for g in self.db.gruplari_getir()}
        self.secili_derslik_id = None

        grup_duzenle = tk.LabelFrame(self, text="Grup Düzenle", padx=16, pady=10)
        grup_duzenle.pack(fill="x", padx=16, pady=(12, 8))

        ust = tk.Frame(grup_duzenle)
        ust.pack(fill="x")
        tk.Label(ust, text="Grup").pack(side="left")
        self.grup_combo = ttk.Combobox(ust, values=list(self.grup_map.keys()), state="readonly")
        self.grup_combo.pack(side="left", fill="x", expand=True, padx=(6, 0))

        grup_butonlar = tk.Frame(grup_duzenle)
        grup_butonlar.pack(fill="x", pady=(8, 0))
        tk.Button(
            grup_butonlar, text="Yeni Grup Ekle", bg=RENK_VURGU, fg="white",
            relief="flat", bd=0, cursor="hand2", command=self._grup_ekle,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            grup_butonlar, text="Grubu Yeniden Adlandır", bg=RENK_VURGU, fg="white",
            relief="flat", bd=0, cursor="hand2", command=self._grubu_yeniden_adlandir,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            grup_butonlar, text="Grubu Sil", bg=RENK_KIRMIZI, fg="white",
            relief="flat", bd=0, cursor="hand2", command=self._grubu_sil,
        ).pack(side="left")

        self.sayac_label = tk.Label(self, text="", fg="#555555")
        self.sayac_label.pack(anchor="w", padx=16)

        tree_cerceve = tk.Frame(self, padx=16)
        tree_cerceve.pack(fill="both", expand=True, pady=(8, 4))
        self.tree = ttk.Treeview(tree_cerceve, columns=("ad",), show="headings", height=12)
        self.tree.heading("ad", text="Derslik Adı")
        self.tree.column("ad", width=380, anchor="w")
        kaydirma = ttk.Scrollbar(tree_cerceve, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=kaydirma.set)
        self.tree.pack(side="left", fill="both", expand=True)
        kaydirma.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._secildi)

        duzenle = tk.Frame(self, padx=16)
        duzenle.pack(fill="x", pady=(4, 0))
        tk.Label(duzenle, text="Derslik adı").pack(side="left")
        self.ad_entry = tk.Entry(duzenle)
        self.ad_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

        butonlar = tk.Frame(self, padx=16, pady=10)
        butonlar.pack(fill="x")
        tk.Button(
            butonlar, text="Yeni Derslik Ekle", bg=RENK_VURGU, fg="white",
            command=self._ekle,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            butonlar, text="Seçileni Yeniden Adlandır", bg=RENK_VURGU, fg="white",
            command=self._yeniden_adlandir,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            butonlar, text="Seçileni Sil", bg=RENK_KIRMIZI, fg="white",
            command=self._sil,
        ).pack(side="left")

        self._yenile()

    def _grup_map_yenile(self, secili_ad=None):
        self.grup_map = {g["ad"]: g["id"] for g in self.db.gruplari_getir()}
        self.grup_combo.configure(values=list(self.grup_map.keys()))
        if secili_ad is not None and secili_ad in self.grup_map:
            self.grup_combo.set(secili_ad)

    def _grup_ekle(self):
        ad = simpledialog.askstring(APP_ADI, "Yeni grup adı:", parent=self)
        if ad is None:
            return
        ad = ad.strip()
        if not ad:
            messagebox.showerror(APP_ADI, "Lütfen grup adını girin.", parent=self)
            return
        try:
            self.db.grup_ekle(ad)
        except sqlite3.IntegrityError:
            messagebox.showerror(APP_ADI, "Bu isimde bir grup zaten var.", parent=self)
            return
        self._grup_map_yenile(secili_ad=ad)
        self._yenile()
        if self.gruplar_degisti_geri_cagirma:
            self.gruplar_degisti_geri_cagirma()
        messagebox.showinfo(
            APP_ADI,
            "Grup eklendi. Bu gruba giriş şifresini \"Şifreleri Yönet\" "
            "ekranından belirleyebilirsiniz.",
            parent=self,
        )

    def _grubu_yeniden_adlandir(self):
        grup_id = self._secili_grup_id()
        if grup_id is None:
            messagebox.showerror(APP_ADI, "Lütfen önce bir grup seçin.", parent=self)
            return
        eski_ad = self.grup_combo.get()
        yeni_ad = simpledialog.askstring(
            APP_ADI, "Yeni grup adı:", initialvalue=eski_ad, parent=self
        )
        if yeni_ad is None:
            return
        yeni_ad = yeni_ad.strip()
        if not yeni_ad:
            messagebox.showerror(APP_ADI, "Lütfen grup adını girin.", parent=self)
            return
        try:
            self.db.grup_guncelle(grup_id, yeni_ad)
        except sqlite3.IntegrityError:
            messagebox.showerror(APP_ADI, "Bu isimde bir grup zaten var.", parent=self)
            return
        self._grup_map_yenile(secili_ad=yeni_ad)
        self._yenile()
        if self.gruplar_degisti_geri_cagirma:
            self.gruplar_degisti_geri_cagirma()
        messagebox.showinfo(APP_ADI, "Grup adı güncellendi.", parent=self)

    def _grubu_sil(self):
        grup_id = self._secili_grup_id()
        if grup_id is None:
            messagebox.showerror(APP_ADI, "Lütfen önce bir grup seçin.", parent=self)
            return
        grup_ad = self.grup_combo.get()
        ders_sayisi = self.db.grup_ders_sayisi(grup_id)
        if ders_sayisi > 0:
            messagebox.showerror(
                APP_ADI,
                f"Bu grupta {ders_sayisi} kurs tanımlı; grup silinemez. Önce bu "
                "gruba ait kursları \"Kurs/Kursiyer Düzenle\" ekranından silin.",
                parent=self,
            )
            return
        uyari = f"\"{grup_ad}\" grubunu silmek istediğinizden emin misiniz? Bu işlem geri alınamaz."
        if not messagebox.askyesno(APP_ADI, uyari, parent=self):
            return
        self.db.grup_sil(grup_id)
        self.grup_combo.set("")
        self._grup_map_yenile()
        if self.gruplar_degisti_geri_cagirma:
            self.gruplar_degisti_geri_cagirma()
        messagebox.showinfo(APP_ADI, "Grup silindi.", parent=self)

    def _secili_grup_id(self):
        grup_ad = self.grup_combo.get()
        if not grup_ad:
            return None
        return self.grup_map[grup_ad]

    def _yenile(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.ad_entry.delete(0, "end")
        self.secili_derslik_id = None
        derslikler = self.db.derslikleri_getir()
        for d in derslikler:
            self.tree.insert("", "end", iid=str(d["id"]), values=(d["ad"],))
        self.sayac_label.config(text=f"{len(derslikler)} derslik")

    def _secildi(self, event=None):
        secim = self.tree.selection()
        if not secim:
            return
        derslik_id = int(secim[0])
        self.secili_derslik_id = derslik_id
        degerler = self.tree.item(secim[0], "values")
        self.ad_entry.delete(0, "end")
        self.ad_entry.insert(0, degerler[0])

    def _ekle(self):
        ad = self.ad_entry.get().strip()
        if not ad:
            messagebox.showerror(APP_ADI, "Lütfen derslik adını girin.", parent=self)
            return
        try:
            self.db.derslik_ekle(ad)
        except sqlite3.IntegrityError:
            messagebox.showerror(APP_ADI, "Bu isimde bir derslik zaten var.", parent=self)
            return
        messagebox.showinfo(APP_ADI, "Derslik eklendi.", parent=self)
        self._yenile()

    def _yeniden_adlandir(self):
        if not self.secili_derslik_id:
            messagebox.showerror(APP_ADI, "Lütfen listeden bir derslik seçin.", parent=self)
            return
        ad = self.ad_entry.get().strip()
        if not ad:
            messagebox.showerror(APP_ADI, "Lütfen derslik adını girin.", parent=self)
            return
        try:
            self.db.derslik_guncelle(self.secili_derslik_id, ad)
        except sqlite3.IntegrityError:
            messagebox.showerror(APP_ADI, "Bu isimde bir derslik zaten var.", parent=self)
            return
        messagebox.showinfo(APP_ADI, "Derslik güncellendi.", parent=self)
        self._yenile()

    def _sil(self):
        if not self.secili_derslik_id:
            messagebox.showerror(APP_ADI, "Lütfen listeden bir derslik seçin.", parent=self)
            return
        kullanim = self.db.derslik_kullanim_sayisi(self.secili_derslik_id)
        if kullanim > 0:
            messagebox.showerror(
                APP_ADI,
                f"Bu derslik {kullanim} kurs tarafından kullanılıyor; silinemez. "
                "Önce bu dersliği kullanan kursları \"Kurs/Kursiyer Düzenle\" "
                "ekranından silin veya başka bir dersliğe taşıyın.",
                parent=self,
            )
            return
        if not messagebox.askyesno(
            APP_ADI, "Bu dersliği silmek istediğinizden emin misiniz?", parent=self
        ):
            return
        self.db.derslik_sil(self.secili_derslik_id)
        messagebox.showinfo(APP_ADI, "Derslik silindi.", parent=self)
        self._yenile()


# --------------------------------------------------------------------------
# Mazeret Yönetimi paneli
# --------------------------------------------------------------------------

class MazeretYonetimPaneli(tk.Toplevel):
    """Yönetici, devamsız kursiyerler için seçilebilecek mazeret listesini
    (raporlu, revir, sevk vb.) burada yönetir. Önce kaç mazeret
    ekleneceği sayı olarak seçilir, sonra o kadar metin kutusu açılıp
    hepsi birden kaydedilir (Kursiyer Ekle ekranındaki gibi); ayrıca
    mevcut mazeretler tek tek yeniden adlandırılabilir/silinebilir."""

    def __init__(self, master, db: VeriTabani):
        super().__init__(master)
        self.db = db
        self.title("Mazeret Yönetimi")
        self.geometry("480x700")
        self.minsize(440, 560)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)
        self.secili_mazeret_id = None
        self.satir_girdileri = []  # [Entry, ...]

        tk.Label(
            self, text="Devamsızlık mazeretleri (öğretmen yoklama alırken\n"
                       "kursiyeri 'Yok' işaretlediğinde buradan seçer).",
            justify="left", fg="#555555",
        ).pack(anchor="w", padx=16, pady=(14, 8))

        # ---- Toplu mazeret ekleme ----
        ekle_baslik = tk.LabelFrame(self, text="Yeni Mazeret(ler) Ekle", padx=12, pady=10)
        ekle_baslik.pack(fill="x", padx=16)

        sayi_cerceve = tk.Frame(ekle_baslik)
        sayi_cerceve.pack(fill="x")
        tk.Label(sayi_cerceve, text="Kaç mazeret ekleyeceksiniz?").pack(side="left")
        self.sayi_spin = tk.Spinbox(sayi_cerceve, from_=1, to=50, width=5)
        self.sayi_spin.delete(0, "end")
        self.sayi_spin.insert(0, "3")
        self.sayi_spin.pack(side="left", padx=(8, 8))
        tk.Button(
            sayi_cerceve, text="Oluştur", command=self._satirlari_olustur
        ).pack(side="left")

        liste_cerceve = tk.Frame(ekle_baslik)
        liste_cerceve.pack(fill="both", expand=False, pady=(10, 4))
        canvas = tk.Canvas(liste_cerceve, highlightthickness=0, height=140)
        self.canvas = canvas
        dikey_kaydirma = ttk.Scrollbar(liste_cerceve, orient="vertical", command=canvas.yview)
        self.satirlar_frame = tk.Frame(canvas)
        self.satirlar_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.satirlar_frame, anchor="nw")
        canvas.configure(yscrollcommand=dikey_kaydirma.set)
        canvas.pack(side="left", fill="both", expand=True)
        dikey_kaydirma.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(canvas)

        tk.Button(
            ekle_baslik, text="Hepsini Kaydet", bg=RENK_VURGU, fg="white",
            font=("Segoe UI", 10, "bold"), command=self._hepsini_kaydet,
        ).pack(fill="x", pady=(4, 0), ipady=6)

        self._satirlari_olustur()

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=16, pady=12)

        # ---- Mevcut mazeretler ----
        tk.Label(self, text="Mevcut Mazeretler", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", padx=16
        )

        tree_cerceve = tk.Frame(self, padx=16)
        tree_cerceve.pack(fill="both", expand=True, pady=(6, 4))
        self.tree = ttk.Treeview(tree_cerceve, columns=("metin",), show="headings", height=10)
        self.tree.heading("metin", text="Mazeret")
        self.tree.column("metin", width=340, anchor="w")
        kaydirma = ttk.Scrollbar(tree_cerceve, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=kaydirma.set)
        self.tree.pack(side="left", fill="both", expand=True)
        kaydirma.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._secildi)

        duzenle = tk.Frame(self, padx=16)
        duzenle.pack(fill="x", pady=(8, 0))
        tk.Label(duzenle, text="Mazeret metni").pack(side="left")
        self.metin_entry = tk.Entry(duzenle)
        self.metin_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

        butonlar = tk.Frame(self, padx=16, pady=10)
        butonlar.pack(fill="x")
        tk.Button(
            butonlar, text="Seçileni Güncelle", bg=RENK_VURGU, fg="white",
            command=self._guncelle,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            butonlar, text="Seçileni Sil", bg=RENK_KIRMIZI, fg="white",
            command=self._sil,
        ).pack(side="left")

        self._yenile()

    def _satir_ekle(self, sira_no):
        satir = tk.Frame(self.satirlar_frame)
        satir.pack(fill="x", pady=2)
        tk.Label(satir, text=f"{sira_no}.", width=3, anchor="w").pack(side="left")
        entry = tk.Entry(satir)
        entry.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.satir_girdileri.append(entry)
        _agaca_fare_tekeri_ekle(satir, self.canvas)

    def _satirlari_olustur(self):
        """Mazeret sayısı kutusuna göre satır oluşturur/kaldırır. Sayı
        ARTIRILDIĞINDA daha önce doldurulmuş satırlara dokunmaz, sadece
        eksik kalan satırları ekler; sayı AZALTILDIĞINDA ise sadece
        fazlalık satırları kaldırır (dolu satırlar varsa önce onaylatır)."""
        try:
            sayi = int(self.sayi_spin.get())
        except ValueError:
            messagebox.showerror(APP_ADI, "Lütfen geçerli bir sayı girin.", parent=self)
            return
        if not (1 <= sayi <= 50):
            messagebox.showerror(APP_ADI, "Mazeret sayısı 1 ile 50 arasında olmalıdır.", parent=self)
            return

        mevcut_sayi = len(self.satir_girdileri)

        if sayi == mevcut_sayi:
            return

        if sayi > mevcut_sayi:
            for i in range(mevcut_sayi + 1, sayi + 1):
                self._satir_ekle(i)
            self.satir_girdileri[mevcut_sayi].focus_set()
            return

        # sayi < mevcut_sayi: sondaki fazlalık satırları kaldırıyoruz.
        kaldirilacaklar = self.satir_girdileri[sayi:]
        doluysa = any(e.get().strip() for e in kaldirilacaklar)
        if doluysa and not messagebox.askyesno(
            APP_ADI,
            f"Mazeret sayısını azaltıyorsunuz; son {len(kaldirilacaklar)} satırdaki "
            "girilmiş bilgiler silinecek. Devam edilsin mi?",
            parent=self,
        ):
            self.sayi_spin.delete(0, "end")
            self.sayi_spin.insert(0, str(mevcut_sayi))
            return

        for entry in kaldirilacaklar:
            entry.master.destroy()
        self.satir_girdileri = self.satir_girdileri[:sayi]

    def _hepsini_kaydet(self):
        if not self.satir_girdileri:
            messagebox.showerror(
                APP_ADI,
                "Lütfen önce kaç mazeret ekleyeceğinizi girip \"Oluştur\" butonuna basın.",
                parent=self,
            )
            return
        eklenen = 0
        for entry in self.satir_girdileri:
            metin = entry.get().strip()
            if not metin:
                continue  # boş satır, sessizce atla
            self.db.mazeret_ekle(metin)
            eklenen += 1
        if eklenen == 0:
            messagebox.showerror(APP_ADI, "Kaydedilecek bir mazeret metni girmediniz.", parent=self)
            return
        messagebox.showinfo(APP_ADI, f"{eklenen} mazeret eklendi.", parent=self)
        for widget in self.satirlar_frame.winfo_children():
            widget.destroy()
        self.satir_girdileri = []
        self._yenile()

    def _yenile(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.metin_entry.delete(0, "end")
        self.secili_mazeret_id = None
        for m in self.db.mazeretleri_getir():
            self.tree.insert("", "end", iid=str(m["id"]), values=(m["metin"],))

    def _secildi(self, event=None):
        secim = self.tree.selection()
        if not secim:
            return
        mazeret_id = int(secim[0])
        self.secili_mazeret_id = mazeret_id
        degerler = self.tree.item(secim[0], "values")
        self.metin_entry.delete(0, "end")
        self.metin_entry.insert(0, degerler[0])

    def _guncelle(self):
        if not self.secili_mazeret_id:
            messagebox.showerror(APP_ADI, "Lütfen listeden bir mazeret seçin.", parent=self)
            return
        metin = self.metin_entry.get().strip()
        if not metin:
            messagebox.showerror(APP_ADI, "Lütfen mazeret metnini girin.", parent=self)
            return
        self.db.mazeret_guncelle(self.secili_mazeret_id, metin)
        messagebox.showinfo(APP_ADI, "Mazeret güncellendi.", parent=self)
        self._yenile()

    def _sil(self):
        if not self.secili_mazeret_id:
            messagebox.showerror(APP_ADI, "Lütfen listeden bir mazeret seçin.", parent=self)
            return
        kullanim = self.db.mazeret_kullanim_sayisi(self.secili_mazeret_id)
        uyari = "Bu mazereti silmek istediğinizden emin misiniz?"
        if kullanim > 0:
            uyari += (
                f"\n\nBu mazeret {kullanim} yoklama kaydında kullanılıyor; "
                "silinirse bu kayıtlarda mazeret boş görünecektir (kayıtların "
                "kendisi silinmez)."
            )
        if not messagebox.askyesno(APP_ADI, uyari, parent=self):
            return
        self.db.mazeret_sil(self.secili_mazeret_id)
        messagebox.showinfo(APP_ADI, "Mazeret silindi.", parent=self)
        self._yenile()


# --------------------------------------------------------------------------
# Genel Ayarlar (uygulama başlığı vb.)
# --------------------------------------------------------------------------

class GenelAyarlarDialog(tk.Toplevel):
    """Yönetici, ana ekranda görünen metinleri (üstteki uygulama başlığı,
    ortadaki açıklama yazısı ve sağdaki İzleyici/Yönetici bölümünün
    üstündeki başlık) burada değiştirebilir."""

    def __init__(self, master, db: VeriTabani, baslik_degisti_geri_cagirma=None):
        super().__init__(master)
        self.db = db
        self.baslik_degisti_geri_cagirma = baslik_degisti_geri_cagirma
        self.title("Genel Ayarlar")
        self.geometry("480x480")
        self.minsize(440, 440)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)

        frm = tk.Frame(self, padx=16, pady=16)
        frm.pack(fill="both", expand=True)

        tk.Label(
            frm, text="Uygulama başlığı (ana ekranın üstünde büyük yazar; "
                      "birden fazla satır olabilir)", anchor="w",
        ).pack(fill="x")
        self.baslik_text = tk.Text(frm, height=2, wrap="word")
        self.baslik_text.insert("1.0", self.db.ayar_getir("app_basligi", VARSAYILAN_APP_BASLIGI))
        self.baslik_text.pack(fill="x", pady=(4, 12))

        tk.Label(
            frm, text="Sağ panel başlığı (İzleyici / Yönetici butonlarının üstünde)",
            anchor="w",
        ).pack(fill="x")
        self.sag_panel_entry = tk.Entry(frm)
        self.sag_panel_entry.insert(
            0, self.db.ayar_getir("sag_panel_basligi", VARSAYILAN_SAG_PANEL_BASLIGI)
        )
        self.sag_panel_entry.pack(fill="x", pady=(4, 12))

        tk.Label(
            frm, text="Ana ekran açıklama metni", anchor="w",
        ).pack(fill="x")
        aciklama_cerceve = tk.Frame(frm)
        aciklama_cerceve.pack(fill="both", expand=True, pady=(4, 12))
        self.aciklama_text = tk.Text(aciklama_cerceve, height=8, wrap="word")
        aciklama_kaydirma = ttk.Scrollbar(
            aciklama_cerceve, orient="vertical", command=self.aciklama_text.yview
        )
        self.aciklama_text.configure(yscrollcommand=aciklama_kaydirma.set)
        self.aciklama_text.pack(side="left", fill="both", expand=True)
        aciklama_kaydirma.pack(side="right", fill="y")
        self.aciklama_text.insert(
            "1.0", self.db.ayar_getir("ana_aciklama", VARSAYILAN_ANA_ACIKLAMA)
        )

        tk.Button(
            frm, text="Kaydet", bg=RENK_VURGU, fg="white",
            font=("Segoe UI", 10, "bold"), command=self._kaydet
        ).pack(fill="x", ipady=6)

    def _kaydet(self):
        baslik = self.baslik_text.get("1.0", "end").strip()
        sag_panel_basligi = self.sag_panel_entry.get().strip()
        aciklama = self.aciklama_text.get("1.0", "end").strip()
        if not baslik:
            messagebox.showerror(APP_ADI, "Lütfen uygulama başlığını girin.", parent=self)
            return
        if not sag_panel_basligi:
            messagebox.showerror(APP_ADI, "Lütfen sağ panel başlığını girin.", parent=self)
            return
        if not aciklama:
            messagebox.showerror(APP_ADI, "Lütfen açıklama metnini girin.", parent=self)
            return
        self.db.ayar_kaydet("app_basligi", baslik)
        self.db.ayar_kaydet("sag_panel_basligi", sag_panel_basligi)
        self.db.ayar_kaydet("ana_aciklama", aciklama)
        if self.baslik_degisti_geri_cagirma:
            self.baslik_degisti_geri_cagirma()
        messagebox.showinfo(APP_ADI, "Genel ayarlar güncellendi.", parent=self)
        self.destroy()


# --------------------------------------------------------------------------
# Resmi Tatiller paneli
# --------------------------------------------------------------------------

class ResmiTatilYonetimPaneli(tk.Toplevel):
    """Yönetici, KURUMUN TAMAMINI etkileyen resmi tatil günlerini burada
    yönetir (ör. 29 Ekim Cumhuriyet Bayramı). Buraya eklenen bir tarihte,
    o gün TÜM kurslar için (hangi grupta/derslikte olursa olsun) ders
    yapılmaz - ayrıca bkz. DuzenlemePaneli'ndeki "Ara Tatiller" (sadece
    TEK bir kursu etkileyen, kursa özel tatil aralıkları için)."""

    def __init__(self, master, db: VeriTabani):
        super().__init__(master)
        self.db = db
        self.title("Resmi Tatiller")
        self.geometry("460x560")
        self.minsize(420, 460)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)
        self.secili_tatil_id = None

        tk.Label(
            self,
            text="Buraya eklenen tarihlerde HİÇBİR kursta ders yapılmaz\n"
                 "(tüm gruplar/derslikler için geçerlidir).",
            justify="left", fg="#555555",
        ).pack(anchor="w", padx=16, pady=(14, 8))

        ekle_cerceve = tk.LabelFrame(self, text="Yeni Resmi Tatil Ekle", padx=12, pady=10)
        ekle_cerceve.pack(fill="x", padx=16)

        tarih_satiri = tk.Frame(ekle_cerceve)
        tarih_satiri.pack(fill="x")
        tk.Label(tarih_satiri, text="Tarih").pack(side="left")
        self.tarih_secici = _tarih_secici_olustur(tarih_satiri)
        self.tarih_secici.pack(side="left", padx=(6, 0))

        ad_satiri = tk.Frame(ekle_cerceve)
        ad_satiri.pack(fill="x", pady=(8, 0))
        tk.Label(ad_satiri, text="Açıklama (opsiyonel)").pack(side="left")
        self.ad_entry = tk.Entry(ad_satiri)
        self.ad_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

        tk.Button(
            ekle_cerceve, text="Ekle", bg=RENK_YESIL, fg="white",
            font=("Segoe UI", 10, "bold"), command=self._ekle,
        ).pack(fill="x", pady=(10, 0), ipady=5)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=16, pady=12)

        tk.Label(self, text="Mevcut Resmi Tatiller", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", padx=16
        )

        tree_cerceve = tk.Frame(self, padx=16)
        tree_cerceve.pack(fill="both", expand=True, pady=(6, 4))
        self.tree = ttk.Treeview(
            tree_cerceve, columns=("tarih", "ad"), show="headings", height=10
        )
        self.tree.heading("tarih", text="Tarih")
        self.tree.heading("ad", text="Açıklama")
        self.tree.column("tarih", width=110, anchor="center")
        self.tree.column("ad", width=260, anchor="w")
        kaydirma = ttk.Scrollbar(tree_cerceve, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=kaydirma.set)
        self.tree.pack(side="left", fill="both", expand=True)
        kaydirma.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._secildi)

        butonlar = tk.Frame(self, padx=16, pady=10)
        butonlar.pack(fill="x")
        tk.Button(
            butonlar, text="Seçileni Sil", bg=RENK_KIRMIZI, fg="white",
            command=self._sil,
        ).pack(side="left")

        self._yenile()

    def _yenile(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.secili_tatil_id = None
        for t in self.db.resmi_tatilleri_getir():
            self.tree.insert(
                "", "end", iid=str(t["id"]),
                values=(iso_tarihi_goruntu_metnine_cevir(t["tarih"]), t["ad"] or "-"),
            )

    def _secildi(self, event=None):
        secim = self.tree.selection()
        self.secili_tatil_id = int(secim[0]) if secim else None

    def _ekle(self):
        tarih_str = self.tarih_secici.get().strip()
        try:
            datetime.strptime(tarih_str, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror(APP_ADI, "Lütfen geçerli bir tarih seçin.", parent=self)
            return
        self.db.resmi_tatil_ekle(tarih_str, self.ad_entry.get().strip())
        self.ad_entry.delete(0, "end")
        self._yenile()

    def _sil(self):
        if not self.secili_tatil_id:
            messagebox.showerror(APP_ADI, "Lütfen listeden silinecek tatili seçin.", parent=self)
            return
        self.db.resmi_tatil_sil(self.secili_tatil_id)
        self._yenile()


# --------------------------------------------------------------------------
# Ara Tatil Yönetimi (birden fazla kursa TEK SEFERDE ortak ara tatil ekleme)
# --------------------------------------------------------------------------

class AraTatilYonetimPaneli(tk.Toplevel):
    """Kurs ve Kursiyer Düzenle ekranındaki (tek kurslu) ara tatil girişine
    ek olarak: burada TÜM kurslar işaretlenebilir bir liste olarak
    gösterilir ("Hepsini İşaretle" / "Hepsini Kaldır" ile toplu, ya da
    tek tek işaretlenip kaldırılabilir - ör. hepsini işaretleyip birkaçını
    tekrar kaldırmak gibi), ortak BİR tarih aralığı girilip "Seçili
    Kurslara Ekle" ile hepsine birden tek seferde uygulanır - onlarca
    kursa tek tek girmektense çok daha hızlıdır. Alttaki listede DAHA
    ÖNCE (bu ekrandan ya da Kurs Düzenle ekranından) eklenmiş TÜM ara
    tatiller görünür; yanlış girilmiş bir tarih varsa buradan seçilip
    düzenlenebilir ya da silinebilir."""

    def __init__(self, master, db: VeriTabani):
        super().__init__(master)
        self.db = db
        self.title("Ara Tatil Yönetimi")
        pencere_genisligi = 720
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        pencere_yuksekligi = min(820, ekran_yuksekligi - 80)
        pencere_yuksekligi = max(480, pencere_yuksekligi)
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.minsize(600, 460)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)

        self.kurs_vars = {}  # ders_id -> tk.BooleanVar
        self.secili_ara_tatil_id = None

        disaridaki = tk.Frame(self)
        disaridaki.pack(fill="both", expand=True)
        canvas = tk.Canvas(disaridaki, highlightthickness=0)
        kaydirma_cubugu = ttk.Scrollbar(disaridaki, orient="vertical", command=canvas.yview)
        icerik = tk.Frame(canvas)
        icerik.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas_pencere_id = canvas.create_window((0, 0), window=icerik, anchor="nw")
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfigure(canvas_pencere_id, width=e.width),
        )
        canvas.configure(yscrollcommand=kaydirma_cubugu.set)
        canvas.pack(side="left", fill="both", expand=True)
        kaydirma_cubugu.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(canvas)
        self.canvas = canvas

        tk.Label(
            icerik,
            text="Aynı ara tatili birden fazla kursa TEK SEFERDE eklemek için: "
                 "aşağıdan kursları işaretleyin, ortak tarih aralığını girin "
                 "ve \"Seçili Kurslara Ekle\"ye basın.",
            justify="left", fg="#555555", wraplength=660,
        ).pack(anchor="w", padx=16, pady=(14, 8))

        # ---- Kurs seçimi ----
        kurs_grup = tk.LabelFrame(icerik, text="Kursları Seç", padx=12, pady=10)
        kurs_grup.pack(fill="x", padx=16)

        toplu_buton_cerceve = tk.Frame(kurs_grup)
        toplu_buton_cerceve.pack(fill="x", pady=(0, 8))
        tk.Button(
            toplu_buton_cerceve, text="Hepsini İşaretle", command=self._hepsini_isaretle,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            toplu_buton_cerceve, text="Hepsini Kaldır", command=self._hepsini_kaldir,
        ).pack(side="left")

        kurs_liste_cerceve = tk.Frame(kurs_grup)
        kurs_liste_cerceve.pack(fill="both", expand=False)
        kurs_canvas = tk.Canvas(kurs_liste_cerceve, highlightthickness=0, height=180)
        self.kurs_canvas = kurs_canvas
        kurs_kaydirma = ttk.Scrollbar(kurs_liste_cerceve, orient="vertical", command=kurs_canvas.yview)
        self.kurs_liste_frame = tk.Frame(kurs_canvas)
        self.kurs_liste_frame.bind(
            "<Configure>", lambda e: kurs_canvas.configure(scrollregion=kurs_canvas.bbox("all"))
        )
        kurs_canvas.create_window((0, 0), window=self.kurs_liste_frame, anchor="nw")
        kurs_canvas.configure(yscrollcommand=kurs_kaydirma.set)
        kurs_canvas.pack(side="left", fill="both", expand=True)
        kurs_kaydirma.pack(side="right", fill="y")
        _canvasa_fare_tekeri_ekle(kurs_canvas)

        # ---- Ortak tarih aralığı + toplu ekle ----
        tarih_grup = tk.LabelFrame(icerik, text="Ortak Ara Tatil Tarihi", padx=12, pady=10)
        tarih_grup.pack(fill="x", padx=16, pady=(10, 8))
        tarih_satiri = tk.Frame(tarih_grup)
        tarih_satiri.pack(fill="x")
        tk.Label(tarih_satiri, text="Başlangıç").pack(side="left")
        self.ortak_bas_secici = _tarih_secici_olustur(tarih_satiri)
        self.ortak_bas_secici.pack(side="left", padx=(4, 14))
        tk.Label(tarih_satiri, text="Bitiş").pack(side="left")
        self.ortak_bit_secici = _tarih_secici_olustur(tarih_satiri)
        self.ortak_bit_secici.pack(side="left", padx=(4, 0))
        tk.Button(
            tarih_grup, text="Seçili Kurslara Ekle", bg=RENK_YESIL, fg="white",
            font=("Segoe UI", 10, "bold"), command=self._secililere_ekle,
        ).pack(fill="x", pady=(10, 0), ipady=6)

        ttk.Separator(icerik, orient="horizontal").pack(fill="x", padx=16, pady=12)

        # ---- Mevcut ara tatiller (düzenle / sil) ----
        mevcut_grup = tk.LabelFrame(icerik, text="Mevcut Ara Tatiller (Düzenle / Sil)", padx=12, pady=10)
        mevcut_grup.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        tree_cerceve = tk.Frame(mevcut_grup)
        tree_cerceve.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            tree_cerceve, columns=("kurs", "grup", "derslik", "baslangic", "bitis"),
            show="headings", height=8,
        )
        baslik_metinleri = {
            "kurs": "Kurs", "grup": "Grup", "derslik": "Derslik",
            "baslangic": "Başlangıç", "bitis": "Bitiş",
        }
        sutun_genislikleri = {"kurs": 160, "grup": 90, "derslik": 90, "baslangic": 90, "bitis": 90}
        for k in ("kurs", "grup", "derslik", "baslangic", "bitis"):
            self.tree.heading(k, text=baslik_metinleri[k])
            self.tree.column(
                k, width=sutun_genislikleri[k],
                anchor="w" if k in ("kurs", "grup", "derslik") else "center",
            )
        kaydirma2 = ttk.Scrollbar(tree_cerceve, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=kaydirma2.set)
        self.tree.pack(side="left", fill="both", expand=True)
        kaydirma2.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._secildi)

        duzenle_cerceve = tk.Frame(mevcut_grup)
        duzenle_cerceve.pack(fill="x", pady=(10, 0))
        tk.Label(duzenle_cerceve, text="Başlangıç").pack(side="left")
        self.duzenle_bas_secici = _tarih_secici_olustur(duzenle_cerceve)
        self.duzenle_bas_secici.pack(side="left", padx=(4, 14))
        tk.Label(duzenle_cerceve, text="Bitiş").pack(side="left")
        self.duzenle_bit_secici = _tarih_secici_olustur(duzenle_cerceve)
        self.duzenle_bit_secici.pack(side="left", padx=(4, 0))

        duzenle_buton_cerceve = tk.Frame(mevcut_grup)
        duzenle_buton_cerceve.pack(fill="x", pady=(8, 0))
        tk.Button(
            duzenle_buton_cerceve, text="Seçileni Güncelle", bg=RENK_VURGU, fg="white",
            command=self._secileni_guncelle,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            duzenle_buton_cerceve, text="Seçileni Sil", bg=RENK_KIRMIZI, fg="white",
            command=self._secileni_sil,
        ).pack(side="left")

        _agaca_fare_tekeri_ekle(icerik, self.canvas)

        self._kurslari_yenile()
        self._ara_tatilleri_yenile()

    def _kurslari_yenile(self):
        for w in self.kurs_liste_frame.winfo_children():
            w.destroy()
        self.kurs_vars = {}
        dersler = self.db.dersleri_getir()
        if not dersler:
            tk.Label(
                self.kurs_liste_frame, text="Henüz eklenmiş bir kurs yok.", fg="#777777",
            ).pack(anchor="w", pady=10)
            return
        for d in dersler:
            var = tk.BooleanVar(value=False)
            self.kurs_vars[d["id"]] = var
            tk.Checkbutton(
                self.kurs_liste_frame,
                text=f"{d['ad']}  ({d['grup_ad']} - {d['derslik_ad']})",
                variable=var, anchor="w",
            ).pack(fill="x", anchor="w")
        _agaca_fare_tekeri_ekle(self.kurs_liste_frame, self.kurs_canvas)

    def _hepsini_isaretle(self):
        for var in self.kurs_vars.values():
            var.set(True)

    def _hepsini_kaldir(self):
        for var in self.kurs_vars.values():
            var.set(False)

    def _secililere_ekle(self):
        secili_ders_idler = [did for did, var in self.kurs_vars.items() if var.get()]
        if not secili_ders_idler:
            messagebox.showerror(APP_ADI, "Lütfen en az bir kurs işaretleyin.", parent=self)
            return
        bas = self.ortak_bas_secici.get().strip()
        bit = self.ortak_bit_secici.get().strip()
        try:
            datetime.strptime(bas, "%Y-%m-%d")
            datetime.strptime(bit, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror(APP_ADI, "Lütfen geçerli bir başlangıç ve bitiş tarihi seçin.", parent=self)
            return
        if bit < bas:
            messagebox.showerror(
                APP_ADI, "Bitiş tarihi, başlangıç tarihinden önce olamaz.", parent=self
            )
            return
        for ders_id in secili_ders_idler:
            self.db.ders_ara_tatil_ekle(ders_id, bas, bit)
        messagebox.showinfo(
            APP_ADI, f"{len(secili_ders_idler)} kursa ara tatil eklendi.", parent=self
        )
        self._ara_tatilleri_yenile()

    def _ara_tatilleri_yenile(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.secili_ara_tatil_id = None
        for t in self.db.tum_ara_tatilleri_getir():
            self.tree.insert(
                "", "end", iid=str(t["id"]),
                values=(
                    t["ders_adi"], t["grup_ad"], t["derslik_ad"],
                    iso_tarihi_goruntu_metnine_cevir(t["baslangic_tarihi"]),
                    iso_tarihi_goruntu_metnine_cevir(t["bitis_tarihi"]),
                ),
            )

    def _secildi(self, event=None):
        secim = self.tree.selection()
        if not secim:
            self.secili_ara_tatil_id = None
            return
        ara_tatil_id = int(secim[0])
        self.secili_ara_tatil_id = ara_tatil_id
        t = self.db.ders_ara_tatil_getir(ara_tatil_id)
        if not t:
            return
        _tarih_secici_ayarla(
            self.duzenle_bas_secici, datetime.strptime(t["baslangic_tarihi"], "%Y-%m-%d").date()
        )
        _tarih_secici_ayarla(
            self.duzenle_bit_secici, datetime.strptime(t["bitis_tarihi"], "%Y-%m-%d").date()
        )

    def _secileni_guncelle(self):
        if not self.secili_ara_tatil_id:
            messagebox.showerror(APP_ADI, "Lütfen listeden düzenlenecek ara tatili seçin.", parent=self)
            return
        bas = self.duzenle_bas_secici.get().strip()
        bit = self.duzenle_bit_secici.get().strip()
        try:
            datetime.strptime(bas, "%Y-%m-%d")
            datetime.strptime(bit, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror(APP_ADI, "Lütfen geçerli bir başlangıç ve bitiş tarihi seçin.", parent=self)
            return
        if bit < bas:
            messagebox.showerror(
                APP_ADI, "Bitiş tarihi, başlangıç tarihinden önce olamaz.", parent=self
            )
            return
        self.db.ders_ara_tatil_guncelle(self.secili_ara_tatil_id, bas, bit)
        messagebox.showinfo(APP_ADI, "Ara tatil güncellendi.", parent=self)
        self._ara_tatilleri_yenile()

    def _secileni_sil(self):
        if not self.secili_ara_tatil_id:
            messagebox.showerror(APP_ADI, "Lütfen listeden silinecek ara tatili seçin.", parent=self)
            return
        self.db.ders_ara_tatil_sil(self.secili_ara_tatil_id)
        self._ara_tatilleri_yenile()


# --------------------------------------------------------------------------
# Tatiller (ara menü) - Resmi Tatiller ve Ara Tatil Yönetimi ekranlarına
# TEK bir yerden erişim sağlar.
# --------------------------------------------------------------------------

class TatillerPaneli(tk.Toplevel):
    """Yönetici Panelinde ayrı ayrı iki buton yerine tek bir "Tatiller"
    butonu olsun istendiği için eklenen küçük ara menü: buradan hem
    Resmi Tatiller hem de Ara Tatil Yönetimi ekranlarına gidilebilir."""

    def __init__(self, master, db: VeriTabani):
        super().__init__(master)
        self.db = db
        self.title("Tatiller")
        pencere_genisligi = 360
        pencere_yuksekligi = 260
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.minsize(320, 220)
        self.configure(bg=RENK_ARKAPLAN)
        self.grab_set()
        _grab_guvenli_baglat(self)
        _baslik_cubugu_rengini_ayarla(self)

        tk.Label(
            self, text="Tatiller", font=("Segoe UI", 14, "bold"),
            bg=RENK_ARKAPLAN, fg="#111827",
        ).pack(pady=(20, 4))
        tk.Label(
            self, text="Hangi tatil türünü yönetmek istiyorsunuz?",
            bg=RENK_ARKAPLAN, fg="#555555",
        ).pack(pady=(0, 14))

        tk.Button(
            self, text="Resmi Tatiller", bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK,
            activebackground=RENK_VURGU, activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 10), command=self._resmi_tatiller,
        ).pack(fill="x", padx=24, pady=6, ipady=10)
        tk.Button(
            self, text="Ara Tatil Yönetimi", bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK,
            activebackground=RENK_VURGU, activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 10), command=self._ara_tatil_yonetimi,
        ).pack(fill="x", padx=24, pady=6, ipady=10)

    def _alt_pencere_ac(self, pencere):
        self.wait_window(pencere)
        try:
            if self.winfo_exists():
                self.lift()
                self.focus_force()
        except tk.TclError:
            pass

    def _resmi_tatiller(self):
        self._alt_pencere_ac(ResmiTatilYonetimPaneli(self, self.db))

    def _ara_tatil_yonetimi(self):
        self._alt_pencere_ac(AraTatilYonetimPaneli(self, self.db))


# --------------------------------------------------------------------------
# Yönetici Ana Paneli (hub) - yönetici bir kez şifre girer, sonra tüm
# yönetim bölümlerine tekrar şifre istenmeden erişebilir.
# --------------------------------------------------------------------------

class YoneticiAnaPaneli(tk.Toplevel):
    def __init__(self, master, db: VeriTabani, baslik_degisti_geri_cagirma=None,
                 gruplar_degisti_geri_cagirma=None):
        super().__init__(master)
        self.db = db
        self.baslik_degisti_geri_cagirma = baslik_degisti_geri_cagirma
        self.gruplar_degisti_geri_cagirma = gruplar_degisti_geri_cagirma
        self.title("Yönetici Paneli")
        # Pencere konumu belirtilmezse Tk bunu ekranın SOL ÜST köşesine
        # yakın açıyordu; ekranda ortalanmış (biraz sağa alınmış) şekilde
        # açılması için x/y burada hesaplanıyor.
        pencere_genisligi = 400
        # NOT: yeni butonlar eklendikçe yükseklik buna göre ayarlandı (660
        # -> 710 -> 760 -> 710) - "Resmi Tatiller" ve "Ara Tatil Yönetimi"
        # artık ayrı iki buton değil, tek bir "Tatiller" butonu altında.
        pencere_yuksekligi = 710
        ekran_genisligi = self.winfo_screenwidth()
        ekran_yuksekligi = self.winfo_screenheight()
        x = max(0, (ekran_genisligi - pencere_genisligi) // 2)
        y = max(0, (ekran_yuksekligi - pencere_yuksekligi) // 2)
        self.geometry(f"{pencere_genisligi}x{pencere_yuksekligi}+{x}+{y}")
        self.minsize(360, 600)
        self.configure(bg=RENK_ARKAPLAN)
        _baslik_cubugu_rengini_ayarla(self)
        # NOT: Bu pencere ("Yönetici Paneli") BİLEREK kilitlenmiyor
        # (grab_set YOK) - yönetici bu panel açıkken ana sayfaya dönüp
        # oradan başka bir işlem de yapabilsin istendi (ör. bir grubun
        # dersine öğretmen gibi girmek). transient(master) yine de
        # kalıyor; bu, pencereyi ana sayfayla aynı görev çubuğu
        # grubunda tutar ama TIKLANMASINI engellemez. İlk açıldığında
        # görünür/öne gelsin diye bir kerelik lift/focus yeterli - eğer
        # kullanıcı sonradan ana sayfaya tıklarsa (ki artık buna izin
        # veriliyor), bu pencere normal pencere sıralamasına göre arkada
        # kalabilir; gerekirse görev çubuğundan/panele tıklayarak tekrar
        # öne getirilebilir.
        self.transient(master)
        self.lift()
        self.focus_force()

        tk.Label(
            self, text="Yönetici Paneli", font=("Segoe UI", 15, "bold"),
            bg=RENK_ARKAPLAN, fg="#111827",
        ).pack(pady=(20, 16))

        butonlar = [
            ("Yoklama (Görüntüle / Excel'e Aktar)", self._yoklama),
            ("Kurs Ekle", self._ders_ekle),
            ("Kursiyer Ekle", self._kursiyer_ekle),
            ("Kurs / Kursiyer Düzenle", self._duzenle),
            ("Derslik Yönetimi", self._derslik_yonetimi),
            ("Mazeret Yönetimi", self._mazeret_yonetimi),
            ("Tatiller", self._tatiller),
            ("Şifreleri Yönet", self._sifreleri_yonet),
            ("Genel Ayarlar", self._genel_ayarlar),
        ]
        for etiket, komut in butonlar:
            tk.Button(
                self, text=etiket, bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK,
                activebackground=RENK_VURGU, activeforeground="white",
                relief="flat", bd=0, cursor="hand2",
                font=("Segoe UI", 10), command=komut,
            ).pack(fill="x", padx=24, pady=5, ipady=8)

    def _alt_pencere_ac(self, pencere):
        """Bir alt pencere (Kurs Ekle, Yoklama Paneli vb.) açıldığında bu
        pencere ("Yönetici Paneli") kapanana kadar bekler. NOT: Bu pencere
        artık BİLEREK kilitlenmiyor (bkz. __init__'teki not) - alt pencere
        kapandığında burada grab_set() TEKRAR kurulmuyor, sadece panel
        tekrar görünür/öne gelsin diye lift/focus yapılıyor."""
        self.wait_window(pencere)
        try:
            if self.winfo_exists():
                self.lift()
                self.focus_force()
        except tk.TclError:
            pass

    def _yoklama(self):
        self._alt_pencere_ac(YoneticiPaneli(self, self.db))

    def _ders_ekle(self):
        self._alt_pencere_ac(DersEkleDialog(self, self.db))

    def _kursiyer_ekle(self):
        self._alt_pencere_ac(KursiyerEkleDialog(self, self.db))

    def _duzenle(self):
        self._alt_pencere_ac(DuzenlemePaneli(self, self.db))

    def _derslik_yonetimi(self):
        self._alt_pencere_ac(
            DerslikYonetimPaneli(self, self.db, gruplar_degisti_geri_cagirma=self.gruplar_degisti_geri_cagirma)
        )

    def _mazeret_yonetimi(self):
        self._alt_pencere_ac(MazeretYonetimPaneli(self, self.db))

    def _tatiller(self):
        self._alt_pencere_ac(TatillerPaneli(self, self.db))

    def _sifreleri_yonet(self):
        self._alt_pencere_ac(SifreYonetimDialog(self, self.db))

    def _genel_ayarlar(self):
        self._alt_pencere_ac(GenelAyarlarDialog(self, self.db, self.baslik_degisti_geri_cagirma))


# --------------------------------------------------------------------------
# Ana pencere
# --------------------------------------------------------------------------

# --------------------------------------------------------------------------
# Kurum logosu (MSÜ arması), ana ekranın ortasında gösterilir; PNG
# olarak base64 kodlanmış biçimde doğrudan kaynak dosyasına gömülüdür,
# böylece EXE tek dosya halinde kalır ve ayrı bir resim dosyası
# taşımaya/gerekmez.
ANA_LOGO_PNG_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAHIAAACWCAYAAAACCNzpAAAKMGlDQ1BJQ0MgUHJvZmlsZQAAeJydlndUVNcWh8+9d3qhzTAUKUPv"
    "vQ0gvTep0kRhmBlgKAMOMzSxIaICEUVEBBVBgiIGjIYisSKKhYBgwR6QIKDEYBRRUXkzslZ05eW9l5ffH2d9a5+99z1n733WugCQ"
    "vP25vHRYCoA0noAf4uVKj4yKpmP7AQzwAAPMAGCyMjMCQj3DgEg+Hm70TJET+CIIgDd3xCsAN428g+h08P9JmpXBF4jSBInYgs3J"
    "ZIm4UMSp2YIMsX1GxNT4FDHDKDHzRQcUsbyYExfZ8LPPIjuLmZ3GY4tYfOYMdhpbzD0i3pol5IgY8RdxURaXky3iWyLWTBWmcUX8"
    "VhybxmFmAoAiie0CDitJxKYiJvHDQtxEvBQAHCnxK47/igWcHIH4Um7pGbl8bmKSgK7L0qOb2doy6N6c7FSOQGAUxGSlMPlsult6"
    "WgaTlwvA4p0/S0ZcW7qoyNZmttbWRubGZl8V6r9u/k2Je7tIr4I/9wyi9X2x/ZVfej0AjFlRbXZ8scXvBaBjMwDy97/YNA8CICnq"
    "W/vAV/ehieclSSDIsDMxyc7ONuZyWMbigv6h/+nwN/TV94zF6f4oD92dk8AUpgro4rqx0lPThXx6ZgaTxaEb/XmI/3HgX5/DMIST"
    "wOFzeKKIcNGUcXmJonbz2FwBN51H5/L+UxP/YdiftDjXIlEaPgFqrDGQGqAC5Nc+gKIQARJzQLQD/dE3f3w4EL+8CNWJxbn/LOjf"
    "s8Jl4iWTm/g5zi0kjM4S8rMW98TPEqABAUgCKlAAKkAD6AIjYA5sgD1wBh7AFwSCMBAFVgEWSAJpgA+yQT7YCIpACdgBdoNqUAsa"
    "QBNoASdABzgNLoDL4Dq4AW6DB2AEjIPnYAa8AfMQBGEhMkSBFCBVSAsygMwhBuQIeUD+UAgUBcVBiRAPEkL50CaoBCqHqqE6qAn6"
    "HjoFXYCuQoPQPWgUmoJ+h97DCEyCqbAyrA2bwAzYBfaDw+CVcCK8Gs6DC+HtcBVcDx+D2+EL8HX4NjwCP4dnEYAQERqihhghDMQN"
    "CUSikQSEj6xDipFKpB5pQbqQXuQmMoJMI+9QGBQFRUcZoexR3qjlKBZqNWodqhRVjTqCakf1oG6iRlEzqE9oMloJbYC2Q/ugI9GJ"
    "6Gx0EboS3YhuQ19C30aPo99gMBgaRgdjg/HGRGGSMWswpZj9mFbMecwgZgwzi8ViFbAGWAdsIJaJFWCLsHuxx7DnsEPYcexbHBGn"
    "ijPHeeKicTxcAa4SdxR3FjeEm8DN46XwWng7fCCejc/Fl+Eb8F34Afw4fp4gTdAhOBDCCMmEjYQqQgvhEuEh4RWRSFQn2hKDiVzi"
    "BmIV8TjxCnGU+I4kQ9InuZFiSELSdtJh0nnSPdIrMpmsTXYmR5MF5O3kJvJF8mPyWwmKhLGEjwRbYr1EjUS7xJDEC0m8pJaki+Qq"
    "yTzJSsmTkgOS01J4KW0pNymm1DqpGqlTUsNSs9IUaTPpQOk06VLpo9JXpSdlsDLaMh4ybJlCmUMyF2XGKAhFg+JGYVE2URoolyjj"
    "VAxVh+pDTaaWUL+j9lNnZGVkLWXDZXNka2TPyI7QEJo2zYeWSiujnaDdob2XU5ZzkePIbZNrkRuSm5NfIu8sz5Evlm+Vvy3/XoGu"
    "4KGQorBToUPhkSJKUV8xWDFb8YDiJcXpJdQl9ktYS4qXnFhyXwlW0lcKUVqjdEipT2lWWUXZSzlDea/yReVpFZqKs0qySoXKWZUp"
    "VYqqoypXtUL1nOozuizdhZ5Kr6L30GfUlNS81YRqdWr9avPqOurL1QvUW9UfaRA0GBoJGhUa3RozmqqaAZr5ms2a97XwWgytJK09"
    "Wr1ac9o62hHaW7Q7tCd15HV8dPJ0mnUe6pJ1nXRX69br3tLD6DH0UvT2693Qh/Wt9JP0a/QHDGADawOuwX6DQUO0oa0hz7DecNiI"
    "ZORilGXUbDRqTDP2Ny4w7jB+YaJpEm2y06TX5JOplWmqaYPpAzMZM1+zArMus9/N9c1Z5jXmtyzIFp4W6y06LV5aGlhyLA9Y3rWi"
    "WAVYbbHqtvpobWPNt26xnrLRtImz2WczzKAyghiljCu2aFtX2/W2p23f2VnbCexO2P1mb2SfYn/UfnKpzlLO0oalYw7qDkyHOocR"
    "R7pjnONBxxEnNSemU73TE2cNZ7Zzo/OEi55Lsssxlxeupq581zbXOTc7t7Vu590Rdy/3Yvd+DxmP5R7VHo891T0TPZs9Z7ysvNZ4"
    "nfdGe/t57/Qe9lH2Yfk0+cz42viu9e3xI/mF+lX7PfHX9+f7dwXAAb4BuwIeLtNaxlvWEQgCfQJ3BT4K0glaHfRjMCY4KLgm+GmI"
    "WUh+SG8oJTQ29GjomzDXsLKwB8t1lwuXd4dLhseEN4XPRbhHlEeMRJpEro28HqUYxY3qjMZGh0c3Rs+u8Fixe8V4jFVMUcydlTor"
    "c1ZeXaW4KnXVmVjJWGbsyTh0XETc0bgPzEBmPXM23id+X/wMy421h/Wc7cyuYE9xHDjlnIkEh4TyhMlEh8RdiVNJTkmVSdNcN241"
    "92Wyd3Jt8lxKYMrhlIXUiNTWNFxaXNopngwvhdeTrpKekz6YYZBRlDGy2m717tUzfD9+YyaUuTKzU0AV/Uz1CXWFm4WjWY5ZNVlv"
    "s8OzT+ZI5/By+nL1c7flTuR55n27BrWGtaY7Xy1/Y/7oWpe1deugdfHrutdrrC9cP77Ba8ORjYSNKRt/KjAtKC94vSliU1ehcuGG"
    "wrHNXpubiySK+EXDW+y31G5FbeVu7d9msW3vtk/F7OJrJaYllSUfSlml174x+6bqm4XtCdv7y6zLDuzA7ODtuLPTaeeRcunyvPKx"
    "XQG72ivoFcUVr3fH7r5aaVlZu4ewR7hnpMq/qnOv5t4dez9UJ1XfrnGtad2ntG/bvrn97P1DB5wPtNQq15bUvj/IPXi3zquuvV67"
    "vvIQ5lDWoacN4Q293zK+bWpUbCxp/HiYd3jkSMiRniabpqajSkfLmuFmYfPUsZhjN75z/66zxailrpXWWnIcHBcef/Z93Pd3Tvid"
    "6D7JONnyg9YP+9oobcXtUHtu+0xHUsdIZ1Tn4CnfU91d9l1tPxr/ePi02umaM7Jnys4SzhaeXTiXd272fMb56QuJF8a6Y7sfXIy8"
    "eKsnuKf/kt+lK5c9L1/sdek9d8XhyumrdldPXWNc67hufb29z6qv7Sern9r6rfvbB2wGOm/Y3ugaXDp4dshp6MJN95uXb/ncun57"
    "2e3BO8vv3B2OGR65y747eS/13sv7WffnH2x4iH5Y/EjqUeVjpcf1P+v93DpiPXJm1H2070nokwdjrLHnv2T+8mG88Cn5aeWE6kTT"
    "pPnk6SnPqRvPVjwbf57xfH666FfpX/e90H3xw2/Ov/XNRM6Mv+S/XPi99JXCq8OvLV93zwbNPn6T9mZ+rvitwtsj7xjvet9HvJ+Y"
    "z/6A/VD1Ue9j1ye/Tw8X0hYW/gUDmPP8uaxzGQAAWABJREFUeNrtvXd4XWeZ7v1731V2Ve/Fkntvce+KY6f3BENC7wPMwAwzHM7A"
    "cMgwDZg5AwPDCW0IJLQkJnF6nGZbLrEd9yZ3WZLVe9l1tff7Y23ZTkixnebwsa5rX7Il7a211r2edj8N/kSPz87GALhlceE3bllc"
    "+I1zv/fn4z1yzB4GcVnRbTcvKfJuXlLk3bKs6LZzf/andsg/RRB37cK+ZVnRba6jVmdFDZUVNXAdtfqWZUW37dqF/acIpvYndC2i"
    "pgZ961ac22qKPuK66ndZUUOMrsoWeTkB4klHpNPuqimjI6fWbU3sqalBb2xE/VkiL6Hjrrv866itxbl1adFXXIf7sqMmY6qzkRIh"
    "JWJMdTbZURPX5b5baor+V20tzrnvfc8/xe/1C6ipQa+txampQc9XRf+N4nM52aZbVRmVKISnQIrhy1SqqSXmDQxamhL8tE90/dXw"
    "e4eB/bNqfZdUaW0tzqqlJaNML/yIFOL2osKQM6I8qnsewtAlhq6wbQfDEIAUuTkB6SmcZNKdF/TCy6dVZ214amO8972uat+TQK5a"
    "hVZXh9fYiPe+mqJVluOtMU1tUmV5xCktDOm2qwgFJO1dCRraACOf5vY00ZCHpglyskxpmtJJJJ1Radv94LTRkYanaxMHz/ns9xyg"
    "7ynVOmzPvvUtvFtWlBeQtr+j4NPRsM6IiogbCuqaZSlCQcGhE3FCuSNZtGA840blc/xUP48/spYF06PEEh6mKUimHPd0S1yLJRwE"
    "/A8B4+8feaG159y/82cg30oAQdatQqxejQtw2+UlH3Bs9zumLkcW5AXdkuKQFEIIKQWhgKChJUlBxTQM0+SXv32JxpZ+vvmVlby0"
    "4wjzJtoYho7tKJQCpZTq6Ex6PX0pzXK8Bt3Q/v7hDR0PDEvn5NWob3HpA6pd6gAWr0LeXYdXV4e6fVnhZeMrIz9Hqa9HI0ZuVWXU"
    "KcgL6kIIIXDp7E7S0W1xsB7yCgr47o/W0T+UYu7MEcyfVclDTx0nnUqjSw/HcYmGNUCI7KghIxHdsSwv30q775swIjJn6qjwkT88"
    "nWitBbVqFdqqOkQtl67KvRQlUqxa5au2YQlcVVM01lZ8xXXUJ8NBzSgoCLqF+UEhEFIIiMXTtPUGmDZjMiMqCkBAPGERDOiMqs6j"
    "oamfnt44Q3ELgWDc6Dz27m/k+NHjjKkQBIMmAvBQXndvSvX0pLREyrU1XdxjCP7v6tquE8MSmjkvj0sMVLFqFdqlcGJ33YXcsAF5"
    "bhhw+/KSKa7jfsF11ccCph7JzTEpLw27hiY1y/aQUuA4DmlRyqc/eTXjRuUiANfzcF2F5yl+v2Yfu/a28J1vXsOmrQ386v5dPHjv"
    "h3FSLrv2t3Hfb56nINSHkBpKKQKmhmW7bkt7QhsYtEhbTlzTxL2art390PqOQ+eGPZdfjncJ2FGxahVSnOsJdnYiMien3gFgxapV"
    "yM5OxLng3VWDfkAUrXAdPuW66uZQUDOzoiYVpSHXNITs7U8L05AEAhLTkDy3LUlpxWiUZ5NMOQSDOtWVeWRnB6hv6CUaCfDlzy0m"
    "LzfE1/7lGR58ZD/XrJhAOm1TXpqFYQY5enAfSy4Lkkx5JNMurqsozA+qVNrzWtqT2lDMIpl2LU2KRzWdX0xTXS9865xzrqlBLy5G"
    "vUMCIe66C7FhAzLzN10AcdOS4pUBx9y9eltz7ytd/M5ORHExavJkVAZcLvJExV13IerqEJ2diMtr8V7pQNx+RfF01/JucT1WaVJM"
    "DQU1srJMSgoDbjSsy/rmlGjuDjB2TDndPQPkRx16+9PsPuKgaYLO7hjdvQkqSrMZPTKf0dV5FBdGGVGRQ0vbIC1tgxw80sHugy2M"
    "KM3ldHs/5UXZ5OeFAcXcyQaRSICuAZ2ykmwOHGxkbIVi9IiAiiUcr6M7rQ0NWSRTLq6nDmqS1ZopH3loXef+V9r1DTX+TT7nvqmL"
    "NXvn3rdzgTuD04LK/LRuzRI3LynyUKrbU2zWNLHOE2pLwAoeWb2tOfl6HmRn5+vb1+Ji1Ot5fB+7uTo31p+a6TjeStdTVwOzQ0FN"
    "hEM6WVHDK8wPKE2TsqHFEb1Dik07B/n2N2+gZtFI7n1gD9FIgPy8IO3tAwQCOsfre+joipGfG+KGqybiKUUwoLNgdhWrHzsAwNjR"
    "BRyoa+fg4Q66uuOMHV1AWUkWjutRUpxDV3cc13X59Ifn8LNf7+K/f1bLVYtyyY5AdZmh0pbr9fanxVDMlomkQzLlKiHYKYV4Vtfl"
    "89Hc4N57H23sfzvv26oFlaG0mZoolVjsuuoKKViCEIXixkWFA9GInq1rklTaBRSepxpTltoPao+uy/26IY6Hhd5y3/OtvUJc2NMl"
    "BNx5XVWePZSqdDw1wfXULE8xR6Gm65osCZqSYFAnEtbJzjKcoKlLXRcylbJY99IgN99yJXfcNo0n1h6iMD9MXm6IlrZBfrN6D8dP"
    "dXLDVVNRSlFdkYuHQiLQDY0T9d186++v5BNf/APjRhfwT1+7kg9+5n6uvHwsQ7E0hqHhuorTLf04rseTzx9iwaxRvO/GqZQUR2lr"
    "HyRte1y7chLfu3sTe3dsZ9mcXKSu4zrKS1mONzhk6/GEQyrlkLI8HNfrEIj9UoodmmCPLsVRIyvY/PunmvrUBcqkUoiPrizPTyin"
    "wrHVOMfxpoO4LGSK6QhRLRAEgxqO6xGLO4Pi+gUFQ2Ul4cjIEVGvpzelDp2I6xJFVZmJZXvEEi6ptAdK9SlEh4J2IUQnSnUDg1Ij"
    "gScsACWUjlJhENkC8j1FCagyBCVSiALTkBi6xDQlwYBGMKh7kbDuGbomNYnwFEIK6O6zUMEKpk0djes6VJTlsPdAGx1dMQ4f7+JD"
    "75vJnBkVbHmpkdb2QYoKIqx5qo6v/83l/PzXOzAMyff++XrWPFnHF/7+D+giwOO//SiFBWG++Z3nGVmVx/LFo/n5r3dw/ZUT6OlL"
    "MKIil0Vzq3j6hWOsfeEokyeUUF6axZQJxTSc7icUNNi9+wgh0UV2lolSIATK81C243rxhCNTKUem0i6W5WE7Hpbt4SnVg6IDRJsU"
    "dCjoBTWIEAmhRMbOKtNThIFshChUShULKBWoEoTICwUk0YiOJgWnWtOYhsbU8REnJysgGptjsrUjEdd9qRHCSjta+0AWn/7MtWzf"
    "ecp7afteFQ1palx1SEiJlkw6eY6r8pTnTfQ8FyH8J9p2PDx1VqkLAUIIhBDomkDTBaYhMQyNQEC6AUNTui4ECGnoQirlSinBdkAp"
    "0HRFSzd8+W+WETAcDhxup6oyj5lTyzlwuJ333zKdBbNH8PQLx4jF0hTmR0ilHY6d6KaoIMLieVVs3t5EcWGU8tIsFs4ajaZJQkGd"
    "8WMK6RtI8Zmasei65PCxTm68eiI52UH6B5K0tg/yxc8sZOrEEgxDMm50IR1dMVrbh1i6cCS5ebn87jdrmDPFxLJBSoShIQxdyEgo"
    "iO0oJVCe7SiVtl2RTnuabbsFlu0VuI6a7LgKpYZfGcOpFFL690jTBJ7nomkaCP//WRGdZFq5R+ot5SrEyhXLRHVVgXx4zUZ95njv"
    "jHrUFaBJQSzhkJWbz6pbp7F0wQi5Y/FYjtb38dgj67lqURiRF1Ce66mBhFA9g1J59iCRqEFOloGU5walAtPwX54SQpNCpCxPCBAK"
    "tOGLCAago9umuVsjnUoyboTvhQZMk/EjJF/5+n0k7BCf//g8OjqHOFDXzq59LRw72c3yxaMZP7aQUSPzmT65lGMnuykvy2Lnvmaq"
    "R+Tzo//ZwmNPHeKalROZOK6If/3eemZMLePXD+yhta2PosIo6zefpHpEHpdNKycrK8CJ+h6GYhbf/UEt23aeZuqkEvr6kwSDBvGk"
    "zY0fvJfRFTB3cgClBLF4mljCo7EDAgGTUaU2JQW6SKbRdF1gGAbRsC+1AUMq11NKCqXSlsJ2fP9nODNju4qePouBIRcznE9Jjivy"
    "s4QQUohEMi22H/K0j374WspKwiydP4JgMMDqR3aRTA6gSd/WaeMrw18Ph3SzuiLCyfp2GlpcqipzCAY1Fs6uYN/hftp7FK6TFihb"
    "dMaL5Jf+8ibZF9Nlmjx54mSnHDPCkLYrpBRCGoaUPQOOPN5oSdt2ZXNHUhQXmML1FK7rS6ynYO+RFC/ud/m3f1xFIJzLxp09aGYO"
    "/b296MFsysqKycvPYde+Vrq641w2rZxbrp3MypqxTBhXxLRJpfz4nm0IIXj4iUN88LYZ6IbOvb/fzrxxKerqeznZlOKm6ydz4mQ3"
    "v/zdboTqJ9foY/PuAZYuGMXIEXk89fxRXFex/1Abn//EfMpKsrjpmsnMn11FR1eMR58+zN5DbUydVEJeTpBk0qW5dYg+q4STrfDx"
    "D1/OwgWT+c+f7sHQBfnZGkKA64KUAlDiaH1MpCwlT5y2ZDAgZSSsSRBS06QMmcitBxw5avxUOWpMtfzoBy+Xqx8/Lotz06K9V4gj"
    "LWGmTR3NZz86i3TKxnXh+z/eiD10iqqKMN29aeJx29LGV4a/nhUxzOwsg0Ra52hDmg+//zKefPYo3/6vWv71H65i9qxx/ObhYwS1"
    "JN1DJrk5QYqLs7Ftj0072skKKwrzdExDEk9YdMcLufKqRazf1sWU6ZPYt7+JojydUFDD0AUHj6e48dZruOaK8RTkBjh8rIMPvW82"
    "1109nV+tPsHNNy7mS59bzjVXjOPq5eM4caqHg4c7KC2JUpAXJi8njOd5bNvVREv7EMWFEd5/8zSqKgs4efQkd9wxnaxkgNJx5bR0"
    "xfjNw/uwkzafWVrNhIWVtLak+ORHFlGYH+TYyR5Otw5QWBBm/qwRhEIGybTNyYZennjmCNMmlfKNv7uCm66ZzDUrp9Pc6bHr0BD3"
    "/OgOSgqz6OgaZNyoPJYvHU8kt4RDh05QURLE0H0tV3fSYeyUWew7kuCDd65g4/ZWCrMtDEPH0AQHjqdp6jBYsWwM5SVZNDT1snvf"
    "aaRK0hXP5wffvp2RVTl8+q/XMGl8ESOr8/npr7ZSnCeIBGEwZhOL25Y2rjL89dwc01RK0Zsu5kf/fjsvbDiGaeo88MgBykqyaOsY"
    "wLY99h9q5pabFgDwxa8/xojybP7PV69h/aajpJMJ2rotYnGHlJfN1Mml3H7DVIIBnT2HE+w5PIhrp4iGIBIS7D08xKwZVTzx3FHu"
    "e2APj649TDCgEY0E+Yd/XYtlOSyYXUUoqDN7RiWJhEV9Qw/P1Z5EeYof3fMSfQNpvvKFJVSW53Dv73ayZfsxrrp+PDN299P9X9u5"
    "uytG9fc3MKe+h33FWSy+bw9Tx5eRnFbEvb/eQWPjIDdfP5nF86r5xe/2cPREN7F4miefPULacpg5tZzbbpiCpgkSSZu/+8bjfO/H"
    "m1h183S27mzkX/5zPQePdBANmxQXRnny6V2MKXdIJB32H0uy74RGXmEZN18zkZU14xgYTHH8RCtZgSTJlEvasuiJ5/DD776Pn/xy"
    "K/f8bic3XzuZ7Kwgm7fVM378SIRQbNrWwIG6dhbOrWJgIMknP7KI3685QnFOmmTSYyiWkchIyDBLi8Mkhvq5b/UBOnssAqZkRGUO"
    "R451cfREN5/44Bye39TCFUvGMG/2CHr74pxs6CMc1PjF7w9wyy3LqBo5itVPn+brX17B0RPd/Pf/vMjEscWsunkq7V1pVt2+lF2H"
    "+jlwdIARlaXMnF6Brgvmz65izMh8nnj2CNt2NnHT1ZPZc6CVR9ceoqoij2BQY9G8kcy9rIoJ44p46InDnDhxmu7OVsrKSpgxqYzn"
    "t9XxmVvHMOqBo7g/2UEhkin1PSzXJGM1QVVTH2OEQG2oZ1x2gMrrR7J+fwcfu3UW9z24l8OH6rAsBzMQ4fOfWMiyRaMZMyqf0y39"
    "7NjTwue/+gjptMv0KWU89MRBUkmHO2+bwfKlYyjMj1BWmsVLe1o52jBEdn4Fd96xgs6eJHf9r+W8sPEkP71vO9MmljBz+gh+/3gD"
    "n/zYlQylovxuTR2jRuazcWsD162cwO03TuPEqR46um0++oHZ/Or+XSSTNnMvqySddli/5RS//f0mxpSnKMwz6e23fIkcXxn+elbU"
    "MKMRg5ICne7eNMtqZvHXf7mUudMrsWyX+x7YzbTJZXhKcPxUD0ODaQoLIuza10rD6T5WXj6ebTub+H//fjOnWwbp6U1QUZbDbx7a"
    "xYI5I4mETb7z3y9wy7VT+PRHljDrsnFs3HqKex/czYiyHK69cgL7DrYzZ2Yld946AyEEN1w9kebWQb79gw309KW57/6XeGHDYZ54"
    "ej9dbU3MHAvF2S4nWxx6XJfrc3Wqf7QHd91JRDCILqDC0IgDloJqXcMTgKHj7mmmrD5GyYIynm0eYt/O40wqT1Ba4HGyvo3aF0+x"
    "9oXD/Hr1Pnbta+ef/uN5br1uKtdfNYFwyOBjd8wiOxpAATdfO5k1Tx7il/fvYPTIIv7pazdx640z2bmniR/8bDMfft9lbNrWwEt7"
    "m7j1umls3t7EwnljueWG6Xz1n55j1c3TeWDNfvJyQsyZWcmR411s3XmayRNK6e1L8NTzx/jy55fysTtmc+UV4zl0tIehngamjguR"
    "shRDw6p1GMicbBPL8knjzu4YGzcdY/e+ZppbBxgYsnji2SNMnlBMWUk2jz5ziJOn+vjy5xYTDpscOd7Fhi31fPD2GbR1DPLY2jq+"
    "+JlFCKXx+4f3MmViCZu3NbJh8ynuvG06Y6pz6R9MMWt6BZ3dcXbva2HpwpHk54b59eo93H7TVA7UtdPWPsjieaMoLxKk+k5QnGNR"
    "lJ2mulhimgazlk5h2bQyJjx7nBG/2IfVPIAIBRmOhyzlV5cJwBr2qhUI08Rp7qdiVzsFhsu8D00mWp5Df9cAeRGHsJkkrMexEr1M"
    "mTKKivJC+voTFBZEGD+mkHsf2MP7bppKNBLggUf2U1wUZc70EVw2rZxF80bQeLqPD3z6fqor8ygqiPCb1Xv5/McXsnBuFd/+wQY+"
    "cMs04vE0//bDFxhdXcC8WSO48epJ/Mf/20TfQIJpk0rp6Utwz293MmlCCRKbA4caefLpffT3dlKY7fihmhQMDr0SyCwT14NwEDQ1"
    "iJPqY6ivg90HuzGCUW68agJzZlTwnR/U0t0Tx7Jc+gdSjBtdQDxuMTiU5s7bZnCqqY8NW07S3hHDNDUqK3LYu7+NSeOKOXC4g2UL"
    "R1FakkUiabPvUBvf+faNzJ9ezoxZI3jsyYMMDqVpON3Ptp1N/OK/bqO7L83+3buZPDaCnXYJhQJUTapmwfypjDmRouD7O8ja3ISN"
    "QJj6GRBfmaMTr6RNDB3H8cg51ElkdxelY8soWzoBW9eJ98VxHZfK0hAHj3Rx7TXz+OiqGXztX54lKxqgsytG9Yhc7rxjFsvmVXPV"
    "ivHc//A+Ll8ympzsIOs3n+L52hOsWDaGQ0c6mDurknTKYd3mek409HDtigkEAjq1WxqYMrGUkuIojz5VR0vbIA1NfdQ39vGXn1xA"
    "OGTQ3hmjob6RvEA3mjdA2EgRCmp4nu8VvyqQnucHqZrUyYroHGuW/P3/+gBzZpaTSPnAfeC2GfT3p/CUYnAwRXdPAqkJunvifOj2"
    "mTSc7qOnN8nPvncbSxeOpKoil+27T5OfGyKZcmjvGOK266eQmx3k+Y0n2bOnmdKSLNavP47jKj7/8fmcbhngf3+pht8+tJ/Ghk6s"
    "oQ6KS7IYPXUU8+ZMZHKPIPtHu9F/uw9607ghA3Gh1LTyiQvX1BGdCbT1J4kcH2DE1BGULRiDlhUgNpAgMRhnIGHQ3B7nf39xGaeb"
    "/XPbtK2J/t44Sim+d/cmsrICXLdyApomueu7z1NUGKEgL4KuS/7280u5/qoJLF88mieeOcryJaMJBjSefO4IhXkRDh/rpKMzRk52"
    "kFnTK/jrv1hK3dFOykpyuePWGcy8bAIbNh1jwqgwjivPXObrAikEw/QTfQMunT1p2lo7eGnnEdq7knh2gsefPca//+P1fPj9l3Hi"
    "VDdbdzRR3zTIkgWj+ND7ZrBt52m27mhi38E2vvmd5xhdXcAP/u0GdE3ywsYTpG3FrOmVLF88mtb2QbbuauaHP3uRtrZe5l5WhSYE"
    "u/a28sxzO5k8AVZeO4MZo6uZ0GCR89O9aPfuRbYO+axChtu76JyQAnQNDB3R1I989iRZRwcYUVVC+ZwxVEwqoa2lid2720AEGV2d"
    "y1Asxd3/s5EXNjchpcaUicV84oNzSFkeP/qfbdQdbedLn1nM5z8xn1/8dif3PbibZNrhiWePMml8MZ/9zEIeffIIv394P1LAnJmV"
    "/Mc/Xks4EuD+h3cwojTIxq2N9HW3IFSaE/UdaGqQaEjieYrh6s43BHI4aC/OkyQGO1BWP1UlGgXRNB1tHQgzn0njS0jFLbbtaSE3"
    "P8Sc8Rovbj1OIqUxcWwBDzx6gNKSLC6bVs7Jhl6+/+PNjByRx203TON7P3yatrZu2jqT2JbL/gNNlOcMkUqlmDh9Apt2n0DJIT77"
    "/mksys6n+IU2Ij/ZjfboYURbDBEwwdBepkbf9KF8R0hoGqK5H7nhFOFtbRS6BpfNH8OMZVVsPtREa59NJBKh5cQBxlZouJ5GUX6E"
    "fQda+e39m9n04nH++Rs3sH7TCb7x7WcpLc5mxdIxNLcNUHekg1uvncTTa4/w8MObWDavCEeLYGqC4oIIvQMpmpt7CYs2xlYqinM9"
    "Bvs6UVY/RfkmtnsWxPMCcvjwFIRDJoah43o+79fUZrNw8SzyckOESrJwBlPUjA1z463jGDMhj1NtHXS1n2agJ82nPrIEPJcT9T0o"
    "TXLFsvH84eFNrFgYRNMSNLd10DvUx8RJEa6+ajxjR5Vz7KVGbszTuX7AIfTLA4hf7kfuaUYM2BA0fek5Q1S+xcfwZxo66DqiJ4nc"
    "1wJP1pO9r4uVVbkEPIuXDrdz7U2TmTGnBE8maWpppbu7lcIcm7EjNDa91MHcuePZd7CN4oIwNQtHMWVCGTt3HAJ9AE9Pc9vN41ix"
    "uBrVb5FVUUTRiBzCuoauGwz1dVJeHMB2fPrPNDWcV4D4SiCHsx/REeURHOePf/mVh+d59Mc1IoEi8tpt5r/UyBQdkiPyYFQueZML"
    "eGEwSV1BEcvmVfPY88eoKorSdPg0pQUGi2cWEdI1NEsRdjzMQQta4zgn+wg09WO0DOG0x0l7FhIdTB004Wfn1DtcjSKE7/a6CiwH"
    "D5eQbiKKQlgV2TjVecjRuaiyCHaOSdKQqKCkezDNE7XNhKJhikaX0t4d54arJ/Hsc3XcZKapdgTxo73IxgH0Uz08Y5ocnlJGzvgI"
    "g6kODGkj3gAIpUDXBadb47R1JGIXBKQQ4DoeeUVZ3Dp3JsYHHsF0FGkNhOui8DAQNOFx98QKRk4tY+zGk5idAyyuyiMSMYgP2Xgp"
    "B5F2UEkHHAfwEIBCooSGMKWfz1K8terzzRxSgADlKpTtIZSLyOR+FdKX4JCBCmhoIZ2sLJPu/hQvtg5gVhVwfP5Iktsa+MTpbkwE"
    "CgVIpK5jOh4JU8P7/c089NIeBrpjaLp83ef2lUDqF3NNactjMGZRkBUgPZRG6AIMnyy2gJEKvnWkk9iRFso1ExEMEGsaZBCFzKS4"
    "kMIHK2ieiQ2EAjGsNt1LrPIw80AJQJgShPay88ZTkLQRcQtPKXqVwhSSG4MBrNP9zGjaQY40cYNBXHE2plUoUmjoEZPumIVtX9x1"
    "6xetcaTvEZ1x+zOPjwTSgB7QyZcGg56H8BQyoCHPRHPq7Je3QuKGXe0LljIJnndxtvTV/p4m/BcCDR+kfk+BqROVBpan/Hv2isoc"
    "oRTK8/OS4iILVN+WljIBKKVwXQ8tw66cufhhiXsrHU7LAk1y3nchoxFUKvHWO0uvuE4JSKXwXA+h1NtWSPze7w0UAr2yApUc9G+e"
    "fINL0jRwHVQqiTlvLudUhL6nj/c2kJqGZ/UT/sSHyP2Pf0fZcV/K9NewGLqOSsZA8yj4zT1E7lyFl+p97d//M5Dv0KEUQhmkHnmK"
    "rK98ieLn16KPH4uX6PIlT8izqlTX8RJdGFMmUbxxHeEPvZ/E/Q8hMN/5sObPQP5RUIsIRLAO7iP94nYCV9RQsq2W6Kf+Ai/Zh7It"
    "hGGA5+Iluol+9JMUb1mPOW82zpFjWPv2I4wweO6fgbwU1KvykqQefRI8D5mbRf7/3E3h736HVpyHG+8AUyP/xz8l/96fI7Mi4Lqk"
    "1m301aph8KcwEuK9D6TnIUSI5GNPo1zXB8VxCN/5Pkq2byL66c9S/MLTRD/3Kb8iSinQNFJPPuNHX+pPY7DHnwaQZhj7aB3W9p2+"
    "15opY9OqKsn/+U8w588ZLmsDTcPt6CL94nakFr64OPLPQL5d6lWiVIrkmifO8lea5oPkuv5XTfP/DaRrN+MOdIBp/lkiLzXvVWpZ"
    "JH77ACoW98OJ4ZhS087GlhnCIHH/H1DK+pMIO/6knB0cB8+NEbzuKp9Wei11KSQoRfQvPok5ZgreUEcmTBF/BvJdPXQDlRwCU5L/"
    "s/8h/567EZHIWTv5R1frfy949QrfEfrIJ/wwxXFA0/8M5Dt+CAGajpfoxJgxneJNLxD9zMfP2MA3fK/rIgvyyL/v5xTedy8yL4qX"
    "zDA871HplO9FEJVl4SU7/AB/w7OYl01HpVK+XXQyOc5Xc2KGnR/8EEVZFuGP3EnJixsIrbwaL9EBrvOeBFP+8X0Sl7YkKhdtRClF"
    "Dz3qB/i52f6PgkFfooZfr6paM86PpvmJYNP0NfS4MRQ99zj5P/wxIicCyrskwHw9LF75I/3lzp/CcTwMQ0ORSUcBAoHneXiXRMyl"
    "CK5cidfXz9B3vodSypcyx/UJAcsCTSf763+LiEZ9ycyEH4nfPUji0TUILQyW5f++46Jsy49HI1FEMIoaiPlpsXdLuqREKYXneei6"
    "zsupJx8L1/H8PspzgRRC4DguBYW5RKJhjtadwjB0POUhhcTzPCLRMMFQiIG+oXc1zEDoxO/5FbF77n4NyyBROJjzZxO66TofZM0v"
    "2Br6jx+S2rsFQYCzLfrnPtoeUkZAN97V+DIRT5KdG8UMmHR39qHp8gyWnlIEgyY5eVnEY7EzZ6+fK42arnHTqis4fqSRVDJFKmkx"
    "NBinv2+IiVNG0dzUQXdH3zljM98dMEUwgpBZr+HJanjxHpJ/eITQzdejlPIf1KPHsY8cRQsVv35Fgee9ayAq5ZfCTJw6hnTaYtbc"
    "SRw70oiUEvCvw7JsRlSXMjSYYMeL+/1ymZepVgWalOzbeYR9u4+gaRLTNKkeU86KaxZQXFXCPf/1wKVhQz3vtWNF10WqIKln1+P1"
    "9SNzcwBIPfMCXqoXGS7KFHxdeubfdT0iWWE+9tmbefqxTVSPrmDB8tm4aevMffc8Dz0S4qFfPfkyjaIPPwm6odHd1U9TYxu6pjFp"
    "2hhmzZvMyDEVDPTHeHr1OtpaujBM3bdL7+hVcv4ZCqUgEMDpbCT13HrC778VgORTF0GSX2wt0EVaDV3XiA3G+fH37wcgvyCHJx5e"
    "j8yYhjMSObLsHIFSL5dIIQSO63LltYuYNX8y4XCIrs5ennvyRfbvOcrQQNyf2Xah81neFIDS9yAdNxOwqwtCP7l6DeH334p7ugVr"
    "6w6kFrkwkty2/TTX8J1+R1w5ON3YRlZ2hI3rdtLXO4h2zpAG5SlaTncQzYpgBgw8z34FkPi/1NbaRVdHH0L08cSaWhrrWwlHgkSz"
    "whnP9R1KwkoNlU4ihIYsyMPr7z+b8T8P1Su1CKnn1qHSadJbt+MOtiPDheevVpWHLCrE7WhDGAH/QXqHvPZgKIDjuP4Ey6h/34cl"
    "cNjmDw3GMU3tj+NITykCAZO6/Se5+z9/x5NrapkwaRQrrllAQWEOsaEErvsOhR+ajkrFkHlRCp98EGPaZJSTfOPCqleoV3eghdRz"
    "60lv2JzxUs9TqjQNzx4idMO1FPzmHpSXRjkWyHdmKqrKNOqkUxaJhD+AzLEdrLSF53kkE6mXhR5ngBRC4Ngu+QXZXHvzUiZNG0M4"
    "EuR0YxtCwPW3Xs4Nt11OMBTwu4HebhCTg2jVlZTs2obb2Exy3VOIYPaFSYTnIWSYoX//L1LPrUOICJzvg+i6yEAusV/+HK26iuJ1"
    "zyKCOspOvWNgep5i9oIpzF04Ddt2CEdDrPrwNXzpf3+E5VfPx7Ze3lZwVrVKQTptMX/JDBZfPuvMg51OWbiuy8RpYzl57DSHD5x8"
    "+zxXTUMlB9HHjKZ43VqU49D7hS8hzbwLV2ueh9DDpLds8/dFGMELd3Qw6Ln9Q5Q1H6X4mafovPYmVCKJ0ANvm5odjulLygqYNnM8"
    "zz31Iq7jcvUNS5g5bzJ7Xqpj9NgRtLd0c/zIqTOhoH6uOAeCJru2H+LQvuNomobj+LPpEokUY8ZVIaV4+zxWqaGsBFp5OUVPP4ZW"
    "VUn3NbehvDRCj1xkyKD84quLcVY8DxGM4nQ2MPDVb5L7/W9TtGY1XdffDJ7jtwy8TfdiOAzZsfUAe3ceYf7i6cyYPZEdW/bz+189"
    "SUFhLlNnjn+Z43k2/NA1+nqHWLd2G4MDsbM6WPg/7+0ewDR1AkET762+ACFAOQjToPCh+9HHjSa9YRPJZ59BBnPfXNz3Zs7VdZBG"
    "LvF77iX6pb8gcMVS8n/xM7o/9CFkKOft6U9RCsPQ6O7s45oblxAMBVi07DLSaYtN63bhuh7FZQWEI0HfCTpjI0WmW0EKbMvGsmwM"
    "00DTpf/SJJqmneH+XPdt8FqlxEsPkvffP8BcMAeA2N0/Ryn7/B2ctyu4M0zcwQ7iv/otKEX4g+8j+ytfxUv2vC0VBgrQNI3BgRhH"
    "Dp3i6huWEI6EePwP6+ho62buwqnceNvltLd1nxUCkcmmCiGwLYeyymKqR5Wzb/dRAgHzjKubTltMnDyKwcE4p040n41r3iLnxkv2"
    "EHn/h4h86iPgebgtbaSefh6pZ51fjvFtZpGECJP8/UNkf/0rCNMk59++SXrdBqw9uxDBt/4cVYZP3b/7KMePNqJpkngsSTAUIBFP"
    "cf+9T9Hd1U8waKK8eEa1+rPtfGcnZXHFNQtYed2iP7KFmiZZ88AL/O3XP87mTXtIJdNIIfzZNW9GpTpptMJycv/rO2cI7uTTz+LG"
    "Oi4NOs3zEIEQ9vEjWNt3Eli2GGEY5N39fTqWLH9rnB5xlqER54AZCgdwbAfbgkDAQCk4eawJmaFP1Vkx9nTAIbP/wgyY4tSJFhpO"
    "NqMb+pmpELbtUDWyDNdxWffMdnr7B6goKn3zjo+m4aV7yfvaP6OVlaLSFkLTSD+3DtAunQo3KVGkST23nsCyxah0GnP+HCIf/Rix"
    "e35yYUTD6+lUpV6WjEnEU+iG7g/ht108pSgoysOxHWKxBLpuqkyDl6MDlqcUUvri+9QjtbS1dJ0Bcpi6yy/IIRg0Od3QTlF5AVp5"
    "Zabr9mJvjt/WZoyaQvRzn8zUpxqotIW194CfarpUgFQKgYG1fYcvNRnuM/sbXyHx4GpIZ2z5xZ7vcPJX+F3aSvn/vv62y+nt7ufA"
    "3uMEAgbplEUoHODqG1ew48UDHD5wYjjDbEkg5XkKTZMMDcYZ6BsiOydKMGgSCAYwTQPTNPBcj0Q8RVZ2BCkzKlV7EycvNZQXJ/qX"
    "n0GEwyjHnwHqNrfgtrQhjEuo5tTzABPnRL1fUqLrKNdFHzWS8Kr34dkDfs7zzYijJlHCJwI8zyMSDhGNhjiw9xiFRbnc+fHrmTF7"
    "IqdONPPSlv3k5mfjut6wVkxJEHGAdNqmamQZo8ZVkkz6Ih0KB8jJz6KgMJcly2dTWl5EOm3huQpPE3BOwvOCbWM6jZZXQfjDH/CH"
    "F2Xyam5za6b1TbukioeFruN19eB1dp9j1hTRz30SoYXenMOjAF3iaiJThaEwAgbHDjfS2dbD8SONHD5Yz46tB/A8RXdnH7HBBEII"
    "PN/DievAkE+7CWVZtrjjzut8EQ4FMEwD3dDQdQ3d0Dl+pDHzZg9PAubwzb6QPJMvyV46RuSa29BKijM3IZNv6+5GYV9atUMqIzGJ"
    "JF5fP1pV5RkC35wzC3PmbKxdLyGC0YtwfjKpMlPDk/idzULgOi6Tp40hJy+K53qUlhcya94U+vsGGT22kvbWbpRSw4HFoI5QfZ7n"
    "G9l02sKxHfp6B2hNpEkmUqSSaQb6Y5SWF5KVHfbjScfFFkDoIptglH8BoVtvONumPazF4nGf4L7kasAEynNQyeTZi3A90HVCN11D"
    "etcmhMy+yJkECkI6jgDPGS78UoweV8mEyaNA+ABPnjoGhEAzdX79kzW+BvMUCtWvg+h2PYVSSmlSsnXjXjat34UR8O2iX7jtEY6E"
    "yMqOYJoGju3iCA8i5jk5sAtQq5aNll1MYMnCTD+/drZH8b3UipHRGoGVyxH/GLo4EIdDj4iJJTw8183EkQGee+pFDh84SSgS8rWi"
    "ruE6LlNmjENqEk8plVGt3bqUqsNz/YHxtmUTzY6g6RrBgJlRb77q9DyPgf4Ymq7hWDZpz4Psc4uYzv/ilZvCnDILraw006MhGA5I"
    "ZSTsJ2UuQUCF1P0xomcYFf+cjamT0UorcNvb/RLLC9ZSCrIDWJnqOMPQ6e7qp6O9F9d1iMWSZ0I9x3bo7x8iFAqgaTquk0JK1SGF"
    "os31/OE9sViS8soirrt5KZomM16R/4eGC4CGUyxpx0HlBf3vXYgalBKFhTFtypmU0bmRsCwsQFxqfYuZ6nQRDiHzcl/BEStkdhbG"
    "hLGg0hdOKWYkUuUFSTkOnushpcR1XTzPxTAMDEM/Ez1EomHSKYv+viEQwh9tpmiTQtDiuArH9UXigfuepu5AvV9b6XnYtkMikSad"
    "SjNqTOWZvRXJtIUqCF20m6ZPHP+qakqrrPCdBvfSagdXjoMszEcWFb3sfIfPUx83BsXFV6mrghCptAUCrLTF5VfO5wMfvY7CYp8A"
    "GGZ7PM8HWtMktu0J109Ct+hCiBbX9bAdJYWA8hGlHD/cgKZJv8QjO0JubhYjR1dQUlZAQ30LAkgkUnhF4QvvOVAK0NCrKnmZKJ4D"
    "pFZehlN/ChEMXBojzKQELPQxozMTmr0/kjytqvJNGXivKEwikWKY+xYCNjz3Eh1tPeiGfmYpjlL+8u9AQMdxlXRdDylEi25ooiWR"
    "9lKuq4JW2lKjxlSIRctmommS4tICsrIi6Iaf/ejpHiAUCjKQsknEkriFhehoFzjwViHQkQX5L8PxjPoKBjBmTMWuP4wQIcC9JFSr"
    "wsacP/ssQfAKIGVR4cXHkGi4hWHisS5/h2XQxPM84rEkuq5hWTau42bSjTqFJfkMDQwqx/GE66pUICBb9KJi2huaRY/teBVKSbo7"
    "+1hxzQJaTndwuqGNjvYeujv7aGvponpUOdm5Ufp7B0kMJbFGBwkEDD8vJy5MIkUo9HKJPCd3GFy5nMSa+y+dZhrPQxAguHL5y9Xq"
    "uUAOO2kXKpWugoCBlR8gcTLhVzNIwcw5k1i2Yi59vYMM9A3S0z1Ae2s3hqEjpGDL+p04/n3vLiqmXb97dVfs+gWFrbblVSA0FYsl"
    "xZNratm8YTeaFDi+6CI1SX/fELqhYQYNEkMJUlGNrPwQtMfAvBBPU7y6U5D5XvD6q5FfLfI5zDdDA75VhLmVRB89HnPhvJed5ytc"
    "2osLPWwPSqOkohrJWAqlFLm52RyrO0VsKEFJeRH5+dlMnlbIrHmTMSMhfv/zR/GUUrbtCZRou3t1V2Y6pFD1lu3OlTLgxQZj0ivJ"
    "J52yyM3LIpAJQYbvped6aFJip2zimqKoLAptA+df+iAy24wtiz8KHKUE10WvHkHwyhUkHlnt1+s47rsLpJcgfMetfseX47xqQlml"
    "0+ew3xfgDSsXyqPENIWV9KvkCopyOVJ3ipdePEBefjaalIQiQUzTYFHNZSQTKZTCs21PIlQ9GV2AlByxHQ+Q9PUOMn5SNdffWoNp"
    "Griu65/esDOi+Z1aA31D9CYTqNF5F8bECOFnzvoHXvfXsv76c/6MRe/dtY3YNjJaRPSzn3htaQS8vv4LV6sC/6EemctgOoVjO0hN"
    "kl+Yw9Lls5kwaSS6rqHpGqlkmq6OXl6s3UN3Vx9CSGzbRUqOnAFSQ9TZtodlu0IKwdOPbqKzvedMe4BtOSQTKWJDcaZdNp78ghwW"
    "LJ2JQmCPzL7gp1Dh4La1v8wunpujxPMIXL6U4FXX4qX6372hDZqOZ/cR/dTH0aqrzo54eTVT19J20X/GHZtH/4C/sSAQMNm6cS9P"
    "P7bpTMw+7MmGI0EG+mPEY0lcD2HZHhqi7gyQhiYPW7ZHKu1JISAeS7J/91HaW7sRQlBYnMecBVP54CdvZOV1CymvLOb6W2uIhEPE"
    "KsJ+ysm9sF0Nzon6N/DkIPffv4Uwg+9OF7GUKCuBXjqK7H/4ih8Gvdo5ZGyjU3+KC06Guwo0A3tUDh0t3djnxIsdbT30dPej69oZ"
    "VsfzVCaGFKRSrrRsD0OThyFTRRcIiPqEpTrTlltsO44aUVEsrrx+EalkmvGTqikszscwdPp6BujtHiCaFeYH370P3TC4cuEMVEkE"
    "Wgb9+eNvdCFKATr2gbrXVlWabyuNGdPI+frX6P/HryEjpX4vxjsJpBMn9we/9EOL4T7LV8mKYDs4R49fWDJcCLAcqMimP0cjpy9M"
    "eEJVhtXxOHb4FOFwEBBomvQzIhkeVtOkSlmu8DzVGQiIegD9LpDfeq4jfsPCwrp02i3W9ZDX2typ3fKBFeQX5NJY30Lt8ztorG+h"
    "vbWbdMqioqqEZCKN8lL0CZuiiYWI5j6Q+huHfZ6HEEHsA4dQ8QQiEj67aOSVHqzrkv1/vkp6yzaSzz2BjBS9M2AaBl68nawv/B3h"
    "99/26iAOAykEzqkGnFONvmY6X+JcAspFTC6iw0oyd95UsvOyOFJXz+x5U3jpxQO0nG6no72XRCyJ47hk50QyDdiuZ6U9TQhR95vn"
    "OuJ3gdQ31CCpxROSHem0e7nrKeXaDof2neTY4VOcOtnsTwKW0q8f0SSnG9sxTYNUIkVHzwBjLytFe/7Y+RMCgQBOcxPW3v0EFi84"
    "O5nqlU9s5lXw+3voXHYVVt0BZCT/7QXTMPDiHYSvvY28H3znde3i8H6N9OateOm+i6jdUTgzS+iPJTiwppZQNMjlV/ohzlU3LEbT"
    "JAP9MZqb2mk42UJObpRN63aSSrsqbbsIyQ6ADTXIc7qx5FbL9kimXBENG7z04n6SyRTBYOBMT/uwrjYM3wmSmkZncxfJyeOJBoPg"
    "nCcxICVKJUk9+YwP5Gupo8zMcVmQT9HTj9B59U3YRw4gI4VvPZjDM13j7YRW3kDB6l+ffbheyz5n5vkkH1+bydhcgH10FASDJCfl"
    "033yOH39g3godmw7yB9+9yzFpflUVpUyakwllVUlTJs7ma3rdvrNVEoTluUhhXzxzKlcXus7+IGA81LaclPxhKNJKVRsKHG2b907"
    "Uxvyssy9Yer0dQzQm6OhJhb6G6/PZ7T08ETHNU+gLPv1612kBNdDq6qkeP1agsuW48Xbz07oeCsOXQfl4cU7iHzwExQ+/ofXVvnn"
    "SqOUuK3tpNfXIvULqA6QAmwHNamQ3hyNvo4+wpEQlmXTcLIFTZO0t3azffM+Hvz10/zXd+7j6T+s43RDG0Kg4glHS1tuyjSdHQCX"
    "1+LJb/mRmli9rrcFxIFk0kEpPO2cqRbD4PmFzDbplEUykcK2HWzLpqWvH29plc+Lno9EDg/MPbKf9PPrz/Csrx0G+JKplRZT9NyT"
    "ZH/5f6NScX+O+fC4FXEREpiZU+clekEX5P3fH1Dw23vOkvWvl5LKgJb4/YO4A20XNqBQALh4S6to6evHTjt+yjBTBCcEBIMmkWiI"
    "UMTPf+7ZcZiTx09jmIaXSNiAOOBjhvgWfuUNNTVoGU96XSrtYtmuEuIsiOlMX55tO1RUlfCRz9zExz93GxUjinFdj5b6dpJzS1GR"
    "kE85nefFCGDov378x5zr66hZYRrkfu87FL/wFOaihXjJPrzkgP846vrZIYLn2NgzLyn9lUyaDq7rjy9LxQjfeDMlW2vJ+rsvnh0G"
    "8XqaJTOwUKVSxH7yS4S8wHGhtoeKhEjOKaWlvg3HcymrKOKjn72ZqTPH4TiuX2aTskjEkhiGRirll944LiqZdhCSdediJwGKizMc"
    "vKGtTaddYnFH+u1afvX5+IkjKS4tAOCW96+ksrqUo4fqmTpzPPlFOXQ2d9ERVbCgEhw7s/vijWIoFxHIIfX8M34jaibkeEMwM3N1"
    "AlfUULJlHUV/eJDQVVeDLvAS3XjJXl9aLcv/PNf1p1ylLVQqhpfoxUv2IKIhIqs+QPELT1P42GqM6VPOOjZvFLO6vlqN/+q32Cf2"
    "IwIX0NKuCf8eLaikParoaevDDBiMGT+CDc+9xKF9JxBCsHTFXD7zxVUsv2o+ruMXZOm6Rixuy1TaQzO0tS/DDqAuE9KNm5bd6cTd"
    "jwcMmZOXG/AS8ZSYs2AqC5bOZOvGveQV5LD8ynls3biXNQ88TyBgMnZ8FcfqTmFmhRgxvgL5bD3ifJtBNQlOGudIPZFPfThDVQre"
    "cK9Txm4iBcbkiUQ+cifh22/BmDwFmVPghwHD5kCTiFAQWVSAMXEioauvJuuvv0jud/+VyKc/jj5q5FkpPB+bm6kG9wYG6LnzE6iE"
    "hZAXljBQnsL+q7ns6euks6Wb3PxsKqtK2LRuF6FwkNvvvIqlK+dhpywmTBnF4ECMplOtBAOG19WdlANxp4Vc/e+PHYvZw9gNe62q"
    "pgb9iSfaEjcuKnomkXQ+YTvK85SSWdkRdm07SG/3AKXlRUhNIxFPkpUdYfL0saRTaYyAQfOxFrquG0H5zDLY0+qvSnqjpLDrIoI5"
    "pLatJ3b3L4j+5Wdek5R+1YdgOEMvBPrE8UQnjif6hc/6JmxgEBWL+YS7aSCzov4krFfaufMF8JxzRtcZ+Id/xmk+cWH9KVJAyoLL"
    "yukaG+X0kwcIhgMMDcaorCrls196P1lZYarGVFL7zDaef3orlVUljJs4EiEFluN5iaQjNKGeefyJtkRNDXptLc4Z1XquiBoaDybT"
    "rhgcsmQwaHD0cAPhSAgzYDByTAWaqTPQHyMcCVF//DQbX9hJOBIikUhzrKEV54NTL4w69lykmUP/176Bc/S4D+KFlHkM20TP829o"
    "5r0yJxutohytegRaWelZEDOq9kxy+EJAzDxkqbXPE7v7bmSw4IJ7PhTgfHAqxxpaSSXTaJq/3POFtdvIL8ghEDR56Ldrefyh9Qz0"
    "DZJOWfR296NJwVDMlsm0KwxNPnguZq/0MASgPlZTHexKx46WFAarRo/I9uLJtJy7YCpzFk6lvLKYlqYOfvOLx0ml0jiOi2kaZz9A"
    "wTU3Lab8OzuQWxogFDg/DlbTUMkhzFmzKd78nL/o83UyDeedwM4UjiHE68eD5yuJmoZ7uoWOeUtxu3surAVdE5BM4y0eSevfz2Xt"
    "Y1tQmR1jvkNpI6VvB+OxJKPGVjJ2QjUjR5fzyIMvMDQY81rakrK9O9lUFIhOuLe2McU5hagvexxratAfrR2wJlVHypRicVaW4QZN"
    "Q55uaqe7s4/6481sXLeLVDKNYeh+o8858mdZNpbrUrV8Itoz9Qj3PBMjSiECIZymY7jHm/whR+eul73YAP+VrzdRIYCmoeIJuq67"
    "DefYEUQg6/z3hfgpWJSpYd21jG0n6unpGkDTJIlEGttyMEyfNRsenBQKBykoyuXg3uO0NneikG57V0J6nvrpg5va1tbUoDc2nk3y"
    "vQzIjzdCLagpo3Oa02nnc4YhtawsQwih0dPdT0N9C1bawXVdbMvBcxUiM1dguK2gp62X7CllFORmI7c3nH985XmIQBRr71a8zgFC"
    "N157tgr93Sz5yEiiSlt033oHqY3rkOF8PyNz3upfQjqJ95m5HB8b5MC2I2imhmHorLx2AeMnjaKjrQfbstE0DalJBgdiHK1roL9v"
    "iHA4QEd3UvT1Wa5hGJ893BjrHsbqVYGsBbVqFdpDT8W6xo8ILVAwPi8n4DqOK8sqipgxeyLhcJBxk6qpHl3BqHGVVIwoYeny2dSs"
    "nEdDfSv9A0P0d/ZTccM0QnX9iNP9mazI+RIFWaS3rsNr6/GnOw6TBe9GC3rGJqqhGN233kHy2ScvnE/VBCQtvMsq6f3CDDa9sBsn"
    "06hz86oVTJo6hokzxhGNhti1/SAov3UDIBoND9cXO22dCS1tuWsf3dz5w1Wr0O6ue3nK/TXvjqHr/5lIuvT0pUUwaNDW0sW4iSOZ"
    "NX8KtuUwccoo9u8+hvIUk2dNoO7gSZqb2gmFggwOJti55xjpry7Eyw6C7Z6/VDkOMlzE0M/+H923fACvf+DMAPp3rHZHKd/b1XWc"
    "k6fovOJaks8+deEgCgG2i5cdJP3Vhezcc4xYLIFjO0yYPIpD+0/wr//wUx7+zVoKC3MZN3EkYydUMWf+FG5etYJodhjlufQNWDKR"
    "dDF0/T9f83l55Tfq6lB33YX8yX3x+rEV4auVp6rycgOu63iyo62b+uOn2bvrCCi4bO4kFi27jLp9J3gkE1cCSE3S096HWZVH8fQR"
    "yOdPIAz9/GOtjGRaB3eTenwt5mUz/Qy9EP4NfrM27/UAHE5ZSUny4Ufpvv0OnOPHkeG8C+9KlgJl2bjfvJwD4TSHdhzDCBhousb7"
    "7ryaeCxJW2sXx480sqjmMq64dhHTZo5n6sxxHDt8ikP7TqDrutvSFteSaXfrY1u6vnHXXci77/7jAphX9b2Li5F1dagpoyKnbdv7"
    "iK4LlZsblAMD/tMkpCQQNLnlAytpa+ni/nufQki/P2Teoun09QxgOQ6dTV0ULBtLTiCE3Nl0YXxkxma6rW3Ef/M7SKcJzJ3jFwgL"
    "kbmpbxGgw7POMwC6rW30/81XGfjaN1BJBxGMXHjluyYhlcT7+GyaLi9hyzO7MIIGVtpm1rwpVFaVMHPOJCZNHUNOXhbtLV28sHYr"
    "B/Ycpfb5Hezfc5xQyKSrJ6l6+9JS08SnjzQlTg5jc15A1tX5tvLhpxMnxlWGl9iOGpuTbbqGJqVu6CxZPosV1yzEdVx++4vHGRqK"
    "Y5omqVSaFdcsYOzEag7uPQ4COk53UX77TIKnk8gTnRcOphFAKEFqwzMkH3oSGQljTJroN8sI8fKB8+cbYgyPxx72jKX0yfPePoZ+"
    "+GP6Pvk5UpvWI4O5Pkt1oV1WmoRkErdmLD1fmMH6tTuwM9l9KSW247DxhV00NbRRNbKMWYumIZTiWF0DiXgqEwFYeAq3uTWhWbb3"
    "/GNbuv9x1Sq01atfPXX/mtHwlCk+8tPGRfelUu5nlafIzw+JwcGEGDmqnOmzJvC7Xz1JU0MrkUiIRCLJhEmjWLZyLo7lMNgfo7dn"
    "AMuy6ekdZMQdc9B3dSDbh8C8gFHTGUpMBLLw2jtIPPYwqUeeQiWT6OXlflPNuST5uSC98jXsAZ8DHkJgHzpM7Id30/eFvyHx0AOo"
    "hIsMZWceEHURIKZxJxQR/6caNmzeR1/PIEJAONMe19s9gGHqtLd2s2/XUXo6epk8bQw1K+dSXlHMnp1HwPNUa3tS9Q9aygzK2w+f"
    "incMY3JBQA5L5R+ejLdPrI5kO7a3OBjQ3EjElK3NXQSCJru3H0LX9UxRkOB9H7wKhOB/frSawYGYvwBakwz2DZFQLhW3z0Lb2IQY"
    "SINxga0GnocwTIQRxm1tJfnsUyR+dT/Wlm14Pb1g6MhoBBEInAXp1V5C4PX2Yu3eR+J3DzDwf/6ZwW/+C6n1z6AGLWQ4x0/bXUyv"
    "oxSQdvBKIqT/75W8ePQETSdbMUwd3dC45f0rqRpVxqmTLSilME0DIQSnG9rYue0g5ZXF7Nt1lPrjjaQd4bZ3JHQPvv9IbdevX08a"
    "zyN3hFi1CpmXLAu099r7ImF97NjROR6ekuFIEMtyUMofErFg6Uxu/dDVPPirJ9m1/SCu6zF2fBXxWILenkGU4zF9/kTmhgswvvg0"
    "Mun6YHoXua1cSrAsPCcOuAgZRasoQx9ZjT6iEllShMzO8rf1WBZqYAC3vQOn8TRuYxNueweKBAITYUTA0M+uJbyoYq2MhxrSsP/7"
    "WnYke9i39TDBSIB0yqbmyrmcOt5Mb88A0ewwLac7CARMPE9hBgxSyTSjx42gvbWLdMr2TjYMynjCOVGab8zoC7WlV69+/Tmlb8RO"
    "K4CfPdGWuOXywk8nks6G1va4V12ZJWKxpNA0ieu45ORlsfyqeRzYedifOJGXRV5BDotqLiMUCnLfzx9BC5kc2H6E4BUzmP5vK9G/"
    "8gwykw66mIF/w2yLNHLPxJru6Vbc0w2kcHj1ymYNgY6QJiIQRGiRjCp239ycHCHA9fA0cP5tJQdEnP3bDhOMBIkNxbn25mVEIiGO"
    "HW5ASMGV1y8iNpTw+dXCHHq6+jFNg4aTLZiGptq7kl4i6Ugp+fTPnmhLrFqF9kY6/g0Z47o6PzPyTG3i1ITqcMi2vGWGId2sqCk9"
    "pXAdlxtuXc7oSSPZvfUgFSOKGT95FHn5WRw/3EB7WzcFhbm0tXZhBk2aT7QRmlFOwaxq5PMn/Wke4k1yqhk1KHQdYQYQZhhhRP74"
    "pQcRunn24XkrNgoIAcpDuS7uv6zgSLlk2/O7CUSCWJbNxKmjWXL5bMZMqCYYMkmnLEpKC2hp7qSnq48FS2YSiYRobuogEgnQ059y"
    "O7uSuqvUdx/d3P2Lmhr0p55645a086L+Gxt9e9mbTKzXndDVqaRblZ1tuEIpmZ2bzax5kzmw6yjtbd10tHVz+EA9R+sa6OsZoKG+"
    "hZKyQi6bO4m6gycxgwbNx1vJWlBN3qQKxAv1CE2+dXGhUr66fi1n560kFTLOlbIc3LuWc2JiiC1rd6IFDGzLYbhcZuumvQSCJkuX"
    "z2bcpJGsfWwTB/cdx7YdRo2tJBwJ0XDiNB7CbTwd09O2t31Q7/7wvHmIp546v6aJ887hrFqFuPde3Knjo+uslPfxZMoN5OcFlW3Z"
    "Ys/OIxypq6ervZd4PImUkmDQRGoSwzBYfPlslqyci1TQeKoVoUmajrWQs2wMuWNKEBtOIXT53tpJNdzcZNm4X19G/axcNj35EugS"
    "O21TPbqcVMpicCCO53nUHThJKmUxadoYRo2ppKW5k4KiXIpLClj/7HaCQdOrbxwSiaQzaITkVWvXx3tWrULU1p6fS3jeQNbWZnjY"
    "J+K9k6uiRy3Hu8NxPTcvJyD9EeImuqGdUxXt4TgeN79/BfOWzeSFJ1+kob4Fx3F9sHXNB/OKceSMLELUnnprJfPtBlFlQPz7pTQs"
    "KKD2ie2gSRzHpWbFXMZPGsneHYczmxkkuq5z4mgTne09TJ81gbkLpxIKB3nu6a3omuB0e8IdGLQ0ifzgmtrOratWob0ag/OmgTzX"
    "Xq7dFK+bUBXR02n3ck1KOzvL0BxHvcIf8bhl1QrmLpvJC09s5rknX6S3ZwDlKapHVdDfO4iLovlYK7krxpEz6j0C5rkgfm0pDQsL"
    "qH18G54UeI5LVk6UiqoSwuEgx482veytwaBJU0M7XR29VI+p4A+/fRbHStM7YNudXQnDU/zLI1u67j5fu3jRQA7by5oa9Gc2JV4Y"
    "Xxm5LJl0JgeCmhMO6dJ3JAWppMWimstYft0i1j/5Is88uYVQyCQY9Msa5syfwszZkzha1wCaoOloRjLHliDW1/sOkJSXJoieh7Id"
    "3G/UcGpePrVPbMPLzJCLZoWxLJsDu49RWlGEn2zwu4zPfIQURLPCHD5QT3dnDylbOKdbYobr8uijW7r+IlO+ccENoRdV4dvYiLoL"
    "ZNuI0BOu4sZ4winNihquaWjSy6w6uP7Wy+nt6uMPv3sW0zTQNEkykSYv33eOYvEk6VSanq5+pC5pOtpM1uVjyJlSgVx/yg9wtUto"
    "3o70QwzleTjfWk79zGw2PrEdTxMk4ilmzZvMmAlVmUmOJol4ksnTxlB/vBlNP2c/h5QMDSaID8VxlXBPNQ7plu0dRKkb7jiddO5t"
    "vJC9FueRxnrD+PIueOzFniE9KG9OW25Xw+mYZjueJ6XEthzisSTtrd1nBhrEY0myciJcfeMSNq7bxUO/e5a+nkFmzZuMUgpXwMbH"
    "t3N8Uhjn2yvwpPLTX1JcGiDaLp5UeN+9khOTI2x83AfRcz1mz5/CLR9YyaJll3HHx65n9PgRtLV00dszSHllEdY5qx38kliBh/BO"
    "NQ1pacvt0oPy5sde7BnirrOx+zsika90fiaNjmy1LO/DyZQr83MDoJTobO9l7qJpdHf10dvdT/XoCmqunMuW9bs53dhGXkE2fT2D"
    "XHdLDcWlBRw73IAW0Gk80kxwViX5S8YiahsQKSczvPBdAlHL0G4hHfc/rmJPrsWmp17CCPkpO8PQySvIZu/OwyilmDpzHAuWzGD0"
    "2EocxyUUDtJQ34Jp6ucWO6j6piHiCduWurh+zfqufRfq3LxlQL6CLGiYODJ83LK8VWnLdYsKwqK3Z1B0tPUwd8FURo2tpKg0n83r"
    "d9HR3k00K8LQQIzZ86cwd+FUAPbtOuJvNpCCpmPN6BOLKLxqEnJLE3IwfX69l285iBKRtPEKQjjfu5p9gQTWQJqcgmw62roJBAxc"
    "16OtuYvuzj6OHKznwJ6j9PUMUDWqnGmz/O7uQ/uOY1kOQgikFKqxOeYODtmakHzo0Y3dT1+Mc/OWApmxl97s2RjrtyX2T6yODlhp"
    "9zrLdt2iwrDs6R7g2OEG2lq7OLTvBJZlEwoHGegbYt7i6eTmZqNlOnJ37ziMGTBQmUmVzSfbUJXZFN86A7mjFdETu7CsyVsAIskU"
    "XlUu1vevYleyhz0bD/Lhv7iFRDxFIGAgM8OKAwF/zJhh6CSTFg0nW9i78winTjaTk5NFZ0cv/X1DBAI6TS0xt3/A0pUQX350U9cv"
    "Zs/G2LqVNz24/S1pZ2prw6upQX9mY/zFCVXhQDrt1riusgvzQxr4jT/+ZjWPwf4Y85fMoLSskLWPbWbhspnYlsO+nYcJBExmzZtM"
    "T1c/jufR1tBBOtug+M7ZaAe7ka19F5bPfFMgJnGnlpH6j5W81HKa/duPYIRM9rxUR8WIEq65aRn7dh6mr3cIwzDwPJWpdZZnSkQ7"
    "23s4uPc4iXiKYECnpT1h9/SlDE+pbz+6uetfa2rQ3woQ3zIgXxGWPDdhRKQsmXLmAXZ21NAU/pC88RNHsmDJTAqKcnnqkVp0Q2fB"
    "khkk4kkO7D1GOm1z1fWLGDWukhNHGhGapKO5m0HPofSjc9Gb45nktPE2hRcZm5hM4taMIfatZWw+cJRj+08RioZwbIeCwlzKKooo"
    "Kslj8vSxnDzayOBADMMwzrQeDn81TR2paxi6oL0raXd2Jw3P46ePbu7+8sWGGW87kOdysg+vTTw+vioyMZFwZkhN2FlRQ3McD9f1"
    "WLZiDpvX76Kn299bMW/xdGKDcbZv2c8Nt9bQ0d7Dlg27MQwDO1NW0tvVT9fAEMUfnksgBXJ/62tvNn9TlBuQSuO+bzp9fzOH2s17"
    "aTrVRiAUwLJsdF0jlbI4sOcY7a3+MpWJU0ZzrK6BWCzh76XOhF/DSRpDE3T2pOy2jqThKu5/dHPXJ1atQstwqOqSBDLjAHHXXcje"
    "ZGKNOxS+LJ5wJmmacHKyA7KvL8axugZC4SBDA3F0w2Dhkpm0t3ZRUJQHwIZnX0IpRSqVZsGSGSQTKZKpNLGBBK0tnRS8fybhvCzE"
    "i02+S/9WxJrDMaLt4H5xAR13jGP9My/R3dmPZvi3aOz4Kgb6h87s4+hs76HldAfzFk1n/ORRHD5YT3wokdnP4atZQxd09aacto6E"
    "4XnqcbOs6wOrViHuvttfDvGWWoO3Q0PV1iLq6lDT5o98yB6MLYgnnHGaJp283KBMJi0GB4bQNL9Ad+bcSZRVFNFY38Lm9buJZoUz"
    "KUKPy6+cx7xF09i/+ygISCbTNB1vIeuqCWRPr0RsbkSk3Dfn0WoSUg5eQOL+0xU0zC9gwxPbiCVS/kx3XeP2D15NcWkBO7ceJBQO"
    "4LoeZtCko7Wb9tYuZi+YyqQpo4hGw7S1dvs+ganR3Zd2WtsTuuN6z5llxbesXt3l1tZe6AD4t5cQeMPU710gVq+us9xc4xbHU7Ut"
    "bXG9uzflBAI6mq77S58NjWg0zPbN+3mxdi+RqN9+rWmS5VfPJysr4u/jyrS+G6aO5bqsf+xFDlVp2D+6Dq88AsnU2e6si/FMyyPY"
    "P7qOQ1Ua6x99Ect1CQZNXNelsCiPl7bsx0pZzJ4/haHBhD+G2vWIREMc2Hec557YQjAU5GhmflzA1OnpSzktbXHd8VStm2vcsnp1"
    "nXXXmeaBt8E/e7scv1p8Gu/nx2LW9FGRP9iuWhaLOyMNQzrRsCFd128z6OnuY9umfZgBA8tyCAYDXHXDEk4ea2Li1DGcONrEwb3H"
    "/fVCjovremiGTtPRZuySMEV3zEI70Y9s6j1/J+iMU5PCnV1J6rsr2Nnbwa6N+9GDpr/jJJFCeYq+3kE62no4cayJ+Yum47oune29"
    "mKaBUh5SSKJZIba/uJ+mU21kZYXo7ks5za1x3XG9zaYmrn8kM0LlW2/jQLa3dT1pLX6x8933JtITKkN/UIJlsZgPZlbUlOmUTcvp"
    "DoKhAIGgSSgU4Mbbr2BL7W40TWPZyrk8/9SLdHX24noegYBJYXEug31D6KZBW1MH/a5F0UfmYiYVYn8rQtN9m6dexx4qfwigd/tU"
    "Bv7XfDbvPczR/acIRoKkUxaFJfnMmjuZ/MJccnKzyM6NIqWk4WQLi5bNpK9ngP7+IXTdX3fU2zPA0GCCaFaI7t6kD6LnbVGeum7N"
    "5u7Y2w3i8LP5th933YX81rfwblpUkCWkeErX5JKKsohTmB/UXVfhOC6z5k5mxbULeOTBF9i57SCf+atVlJQV8sN//zWFxXlUjChh"
    "/MRqpswYx6Z1u6h9/iWkFKSTFvkFOSxeOYuy2ja0721FOAqCxh+v4tUkpGyULnC/vJCOK8rZ/Nxuurv7CUWCWGmb4pJ8rru1hjHj"
    "RpBMpPx1xK6v3uOxJJomSSXT3PezRxkcjGX6G8E0JF29GXXqepuVp6577MWeoeFrf9v5i3cCyNpaXzJ/9ptkevqoyGrHY1Esbo/W"
    "NOFEI7pUStB8uh0BdLb3IqS/F2qgf4hwOEhhUR62ZTN30XTaW7s4WncK2/GJeTNgEo8naTzWQmDJSPKWj0PsaUP0xl/ePDTcn1iR"
    "jfOdKzk2JsDaP2xiKJYkEg2RSqWRUhAM+kF/Ip6kq6OPJ9fU0tLcScvpDuJDCeJxH9xINMyxww1ITTvjnba2J3TXUxtMTVy/ZnN3"
    "7J0C8R2TyFdK5g03lIX1fvsRTZNXlpWE7eLCkGHb7plAeubsSdSsnMvObQc5dOAEscEEn/rL2wlHQvz8v1fT1dFLOBI8k1GQUuJ5"
    "Hk7KZur8CVxWUUno+9vRak+CmVnvYKVwa8aQ/PJ89rS0sG9rHYuWz6Kvd5D9u48ydcY4Wps7SWRKVWzbYeW1C2luamfX9jqi2WFc"
    "x0XTtMx2Ih3X9TD04TgxYbiu95yTa9zyxBNtiXcSxHdMIl8mmRkHaOTk8gellZoZjzuTADsnK6Ap4aeFbNvmpS0H2LPzMJ7r8am/"
    "fB8lZQXc+9NH6GrvzYQoapiExrIcv3s6aNJ6qoPuZIL8O2cRKshGbG8BpfC+uICuT0xh40sHOLqvnkA4wPW31mBbDpOnj2XilFFs"
    "3bj3zAOi6Rr1J5pZXHMZoXCQltMd5ORlIaU8s1bK0CXtXclMsK+eSEdzb33qqabUXSC/VfvOTpp9R4E815v9xYleZ/r8xINOLDwx"
    "kXCme0rZ2VmGJqQkNhhHAYtqZnH1DYuprC7lgXuf4sTRRkLhIJ7yMgDaWGmbsooiKqtK/PrQoEl/7yBNp1oJXD6anCVjcK8Zy/Gp"
    "ETY+u5P21h4+9Vfvo7yymF/c/RBuZn5QTm4Ux3Zob+0+p51eUX+8mcuvnMfEKaNIxJL0dPWj6RqaJmjrSNgdXSnDVepBs7Rr1WOP"
    "9drvhGNzSQA5DCb4fQxHmxKrJ1aFqxJJZ47jKicny5RS9zeuu47LvEXT2P3SITa+sPPMdETbtkmnbSpGlDBn4VRuveNKPFdxcO9x"
    "fz+05jfKnDp8Gqs0TLNKsbN2P67y+U+/y9rfPNTZ1kNLcyeNp1qpWTGH7q5+hgbjvgrVNWKxJEUlPuu0c+shzICOFILm9rjT3ZMy"
    "lOKeRzZ1fSQzJkXWvkszn98VIM887iD8XszEoxOqIjnJpLM4bXtOTpYhNE2K7q5+jh4+hW5odHX2YaV9CSwtK2Th0ssIBkxC4SBj"
    "x1fx5JraM9va/a4ngdQ12pu76Grv9fsSNXlmuO3iy2dx/EgjsaEE0aywP+myvZeaFXOpP96cmb/nAz80GGfvrqOYAQMphWpqibm9"
    "fWldwfcf2dT1l3fdhaytvfjs/nsdyGG7SaaFb+3EkVGVTDkrkklX5WSbBIKmSCRStLV0UVSSz5hxI5g8fQzjJ42ksDiPndsOMnXG"
    "eIQQrHtmG7r+x5ej6zqarvmbaW2H0ooiVlyzkLETqpBScOTQKaQAwzTo7eonEg1RUVXil+9n9hrHYkmCQQOlUKeaYqp/0NIQ4q5H"
    "NnV9PZPZV7zL1UXvOpDDRHsmn7l+UnWkO5V2r48nHHKihjINXfij1NLMXTiNsvIifnPP45w62czyq+YzbdZ4drx4gGOHGzAzm9pf"
    "O8Hh72XcufUgCFh69XyU43Lk0CkM00BqGvF4krETqjl57HQm6Q2GoeE4nlffFBNDMVtKwV89sqnr/2Yy+x6XQInYJQFkJgXm1dSg"
    "r92Y2D6lOnwklXZvGYrbenbUcE1TSsfxOHqons72XpSCzo5eCgtzmTxtDGsf20RsKIH2BnyrEALHcdF1jZNHm8jNyWLJijkk4knq"
    "j58mkUgxf/F0bNumob4Fw9DRpMCyXPdk45AWT9iOLvnQmk3dvzx36tSlcFxylcDDN+jmJcUrlPL+EAxouSNHRN1o2NQc18NxXDRN"
    "MnHKaK65aSkDfUP88icPX9AG2OHwQUrJHR+7nonTx3D0wEls28HzPNbc/7wfzhgasYTlNpyOaam02y+EfN+jmztfuNRAvKQk8pWS"
    "+cym+MnJVZG1ruddMzDk5JsB6UTCuhTCH94upGDS1NE0N3VwYO8xAm+gVv+IcpV+S+DhQydxLJvcvGxON7bzzOObcRyXYECnbyDt"
    "NDbHddt2GyTimke2dG29FEG8JCXylZJ5a01Fpedaa3RNzCktCTvFBUHd8/zcZE5ulGkzx7Nr2yHURZipYclMp20E4Hqev1va0Ojo"
    "SjrtnQndcdVOqZm3rqltab5UQbykgQTfm129GndVTVHUdvmtlOKmgvyAU1Ea0QDh2C4Kf4ObutjEsgApRKbmVCCEUs1tcbenN617"
    "nnrM0PjQ6tqu2Bu1fv8ZyPPkZwFuXVL83wj1V9lZpldVGRWaFMJ7C/ZLDld/K6VUY3NMDQ5ZEiV+tGZz5xdfeQ5/BvJNnuddmdnd"
    "tywr/lvlqv+MRHRGjoi6AUPTHFdddB2WUqBrgrTtug2nY1o87iA08XePbOz8XoZue9djxPeks/M6tN5wrLll0sjo3rTlXTc4ZIdC"
    "Qc0JBzXpehdeVDdcIDUUt51TTTE9kXQHpRQfeGRT569qatAvtqHmz0Cer0e7MX548tjI046lrhiKWUVSE040YsgLNZO6Luju9csy"
    "LMs7pplc90htV+2l7NT8SQD5MuJgQ6Jt3Oic+5XnTo/FnPG27bnZWYYQQojXmww6bA+FQDW3JrzO7pTueN5aVzNueGxDx6n3Iojv"
    "JRv5mh4twK1Li/5TKf42GjWoroy6pi5f1W4O20PL8dzG5pgWi9kIwffWbOr6u1d+5p+BfGePYU7Ou3VZ0UddV/04FNDDlRURJydq"
    "6M4rxnDrmmAgZjvNLXE9mXYSmiY+v2Zj133nfs57/Ua8Vw+PzIaENRu77tMkS1Npt66xaUhv60w6Ugp1dvScUG2dSaexaUhPpd06"
    "TbJ0zcau+2pqGJ7n7b2Xb4TGn8BxDuHeMn1C4LeOLaviCXtGIumKrIjheR6qsTkme/tS0vP4nRlybntoQ2/De9Ue/imq1te0m7ct"
    "K/6s7Xj/HgrqOQDJlDNg6PKrD2/s/Nl73R7+/+UQmdltrKopmnnzksKDNy8pPLiqpmjmMIB/ag8wwP8H3uPm843NZhkAAAAASUVO"
    "RK5CYII="
)


# Kurum logosunun büyük, saydam (soluk) hali - normal (tam ekran
# olmayan) pencere boyutunda kullanılır; ana ekranın alt kısmında
# arka plan filigranı olarak durur.
ANA_LOGO_FILIGRAN_NORMAL_BASE64 = (
    "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"
    "G0k+eo6q6rw8Xo3c3GJ8HhlNlfC4JRas7sTtz8ZSZXqiEUYM9JMyHBz2hRS9lnNQbR6FeS5Wbeomv3wwP/nWWXSFYvy/215m5849"
    "ZGeoaLLglHEZDhFt2ei6E5jrho1lCWQZkimDeNJCYGPoTtdP70gCgCQkZFlClgVCEkiSQEn/f0UGVZFwuWRUxXZ+Tuqltkhn687N"
    "l6V0nRpIpix082jXkKIIQj1JDoVyuPv2q2g41MniFfuYObWWirIsLMtm38F26g93U1qUQXamh5KSTOobOvnnE+soKcqkKD9IdyhO"
    "U2uIukMh5s7fzdp5X0JWZGZd/DdGDivmwT9dTlNLmHA0yePPbOHeR9Zxy41TOO+MIWRlelFkwajZf2Lm1Aruu+Nytu9uYdYld/Ot"
    "L52CLAtKCgLk5gTYsvMIbk3j0vNGkJcXoKGhk65QnIYjXYwfXUZOtg9Vkdm87Qi79rVy+ikD0E3Bl771KBMGGqiK6GtOAdBUhy+1"
    "gUTSxLKOsaqitx/V4XRN00I3nIdZNywME0zTxjAtxxNZNqbp/Jx1zDmK9DkqinM+Po/DGad0C1lyZqg01flzRRFIQmLBqm5sIdHd"
    "k+Ccsydzy43TSOkWX//hM/hFM6OGBGlpT7J7X9vxltNpnhW0tMe57jODCPgUJOHhjl+cy9//mcesaQO4+QdzWbKuk4tOy6czZODS"
    "rPRTIohEUhw8HCU328uACo1wNIVLU0mmjk+VnF/bOPGydNQd9WX/HNPQm6aP7KNRt22DyyWIxkzWbevEMGD4gAxys1Vsy0Y3Qddt"
    "gkEX7lA3l1x3D60dcb7y2WkoisTaTYeYNKaClrYI7R1hfvq7+WiaxPVXTaK1NcIFZw0nkTTo7I4xfXI1tm3zuW/+G5dLZsvOZs6e"
    "PYjPXTuZX9y+mC07WxkxpABfOMHPvz+HYFDlQH0HNZW5SLLEE3O30doe5ppPjMXjUVm76RA5mRrNrV184wuzMQyLNRsbuGjOSA7U"
    "t/PQE+sJBDR+dfsCxo0q45JzR9DSFiY328+iZXsYVJvPgmUhzrvmXipLfVTm6Q73aoJbE06nE4K6xjh7DkbQNIUxgwN4PTIpw8Y2"
    "7b6DEDj3XFYVXBr4vXLfAVn20fAI28K0etOco2coSaDKAt0wcWkq2/b1YJoJBlcF8XqcERYrPUeVlaHywNPNlFcP5IufnsDKdfV8"
    "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"
)


# Aynı filigranın daha büyük hali - pencere tam ekran (maximize)
# durumundayken kullanılır.
ANA_LOGO_FILIGRAN_BUYUK_BASE64 = (
    "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"
    "RK5CYII="
)




class AnaPencere(tk.Tk):
    def __init__(self, db: VeriTabani):
        super().__init__()
        self.db = db
        self.title(APP_ADI)
        self.geometry("980x660")
        self.minsize(860, 560)
        self.configure(bg=RENK_ARKAPLAN)
        _baslik_cubugu_rengini_ayarla(self)

        self._filigran_label = None
        self._filigran_normal_img = None
        self._filigran_buyuk_img = None
        self._filigran_su_an_buyuk = False
        self._normal_pencere_geometrisi = "980x660"

        self._arayuzu_olustur()
        self._tam_ekran_baslat()
        self._periyodik_kontrol()

    def _tam_ekran_baslat(self):
        """Uygulama ilk açıldığında tam ekran (Windows'ta pencereyi
        büyütülmüş/"zoomed" halde) başlar; bu desteklenmiyorsa (farklı bir
        işletim sistemi/pencere yöneticisi) sessizce normal boyutta açılır,
        uygulama yine çalışır."""
        # Pencere daha ekrana çizilmeden (map edilmeden) "zoomed" yapılırsa,
        # bazı pencere yöneticileri tam ekrandan çıkıldığında dönülecek
        # normal boyutu (geometry("980x660") ile ayarladığımız) doğru
        # hatırlamıyor, çok daha küçük bir boyuta dönüyor - bu da filigran
        # logonun üst kısmının kırpılmasına yol açıyordu. update_idletasks()
        # ile pencerenin gerçek/normal boyutunu önce hesaplatıp ekrana
        # oturtuyoruz, öyle ki tam ekrandan çıkınca doğru boyuta dönsün.
        self.update_idletasks()
        try:
            self.state("zoomed")
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)
            except tk.TclError:
                pass
        self._filigran_durumu_guncelle()

    def _filigran_durumu_guncelle(self, event=None):
        """Ana ekrandaki büyük/saydam filigran logoyu, pencere tam ekran
        (maximize) mi yoksa normal boyutta mı olduğuna göre büyük ya da
        normal boyutlu görsel arasında değiştirir. Böylece tam ekrandan
        çıkıldığında (pencere küçültüldüğünde) filigran her zaman alana
        sığar ve üst kısmı kırpılmaz; tam ekranken de biraz daha büyük
        ve etkileyici görünür."""
        if self._filigran_label is None:
            return
        try:
            tam_ekran_mi = self.state() == "zoomed"
        except tk.TclError:
            return
        if tam_ekran_mi == self._filigran_su_an_buyuk:
            return
        self._filigran_su_an_buyuk = tam_ekran_mi
        if not tam_ekran_mi:
            # Tam ekrandan normal boyuta dönülüyor: bazı pencere
            # yöneticileri "zoomed" öncesi doğru boyutu hatırlamayıp çok
            # daha küçük bir pencereye dönebiliyor (bu da filigranın üst
            # kısmının kırpılmasına yol açar). Bilinen doğru normal
            # boyutu burada açıkça yeniden uyguluyoruz ki pencere her
            # zaman filigranın tamamının sığacağı kadar büyük olsun.
            self.geometry(self._normal_pencere_geometrisi)
        self._filigran_label.config(
            image=self._filigran_buyuk_img if tam_ekran_mi else self._filigran_normal_img
        )

    def _arayuzu_olustur(self):
        # ---- en üstte ince vurgu şeridi (marka rengi) ----
        tk.Frame(self, bg=RENK_VURGU, height=4).pack(side="top", fill="x")

        # ---- sol panel: grup butonları ----
        self.sol_panel = tk.Frame(self, bg=RENK_SOL_PANEL, width=230)
        self.sol_panel.pack(side="left", fill="y")
        self.sol_panel.pack_propagate(False)

        tk.Label(
            self.sol_panel, text="GRUPLAR", bg=RENK_SOL_PANEL, fg=RENK_METIN_ACIK,
            font=("Segoe UI", 11, "bold"),
        ).pack(pady=(26, 6))
        tk.Frame(self.sol_panel, bg=RENK_VURGU, height=2, width=36).pack(pady=(0, 18))

        self.grup_butonlari_cercevesi = tk.Frame(self.sol_panel, bg=RENK_SOL_PANEL)
        self.grup_butonlari_cercevesi.pack(fill="both", expand=True)
        self._grup_butonlarini_yenile()

        # ---- sol paneli orta alandan ayıran ince vurgu çizgisi ----
        tk.Frame(self, bg=RENK_VURGU, width=3).pack(side="left", fill="y")

        # ---- sağ panel: İzleyici (üst) / Yönetici (alt) ----
        sag_panel = tk.Frame(self, bg=RENK_SOL_PANEL, width=230)
        sag_panel.pack(side="right", fill="y")
        sag_panel.pack_propagate(False)

        self.sag_panel_baslik_label = tk.Label(
            sag_panel, text=VARSAYILAN_SAG_PANEL_BASLIGI, bg=RENK_SOL_PANEL, fg=RENK_METIN_ACIK,
            font=("Segoe UI", 11, "bold"), wraplength=200, justify="center",
        )
        self.sag_panel_baslik_label.pack(pady=(26, 6))
        tk.Frame(sag_panel, bg=RENK_VURGU, height=2, width=36).pack(pady=(0, 18))

        tk.Button(
            sag_panel, text="İzleyici", bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK,
            activebackground=RENK_VURGU, activeforeground="white",
            font=("Segoe UI", 11), relief="flat", bd=0, cursor="hand2",
            command=self._izleyici_giris,
        ).pack(fill="x", padx=18, pady=7, ipady=11)

        tk.Button(
            sag_panel, text="Yönetici", bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK,
            activebackground=RENK_VURGU, activeforeground="white",
            font=("Segoe UI", 11), relief="flat", bd=0, cursor="hand2",
            command=self._yonetici_giris,
        ).pack(fill="x", padx=18, pady=7, ipady=11)

        # ---- orta alanı sağ panelden ayıran ince vurgu çizgisi ----
        tk.Frame(self, bg=RENK_VURGU, width=3).pack(side="right", fill="y")

        # ---- orta alan ----
        orta = tk.Frame(self, bg=RENK_ARKAPLAN)
        orta.pack(side="top", fill="both", expand=True)

        # Büyük, soluk/saydam kurum logosu: orta alanın alt kısmında,
        # diğer içeriğin arkasında hafif bir filigran olarak durur. Pencere
        # tam ekranken daha büyük, normal boyuttayken (kırpılmayacak kadar)
        # daha küçük hali gösterilir - bkz. _filigran_durumu_guncelle.
        try:
            self._filigran_normal_img = tk.PhotoImage(data=ANA_LOGO_FILIGRAN_NORMAL_BASE64)
            self._filigran_buyuk_img = tk.PhotoImage(data=ANA_LOGO_FILIGRAN_BUYUK_BASE64)
            self._filigran_label = tk.Label(orta, image=self._filigran_normal_img, bg=RENK_ARKAPLAN)
            self._filigran_label.place(relx=0.5, rely=1.0, anchor="s", y=-6)
            self.bind("<Configure>", self._filigran_durumu_guncelle)
        except Exception:
            self._filigran_label = None  # filigran yüklenemezse sessizce geç, uygulama yine çalışsın

        ust_icerik = tk.Frame(orta, bg=RENK_ARKAPLAN)
        ust_icerik.pack(side="top", pady=(36, 0))

        try:
            self._logo_img = tk.PhotoImage(data=ANA_LOGO_PNG_BASE64)
            tk.Label(ust_icerik, image=self._logo_img, bg=RENK_ARKAPLAN).pack(pady=(0, 10))
        except Exception:
            pass  # logo yüklenemezse sessizce geç, uygulama yine çalışsın

        self.baslik_label = tk.Label(
            ust_icerik, text=VARSAYILAN_APP_BASLIGI, bg=RENK_ARKAPLAN, fg="#111827",
            font=("Segoe UI", 20, "bold"), justify="center",
        )
        self.baslik_label.pack(pady=(0, 10))

        ttk.Separator(ust_icerik, orient="horizontal").pack(fill="x", padx=140, pady=(0, 16))

        self.aciklama_label = tk.Label(
            ust_icerik, text=VARSAYILAN_ANA_ACIKLAMA,
            bg=RENK_ARKAPLAN, fg="#4b5563", font=("Segoe UI", 10), justify="left",
        )
        self.aciklama_label.pack(pady=(0, 22))

        saat_kapsul = tk.Frame(ust_icerik, bg="#e8ecf3")
        saat_kapsul.pack()
        self.saat_label = tk.Label(
            saat_kapsul, text="", bg="#e8ecf3", font=("Segoe UI", 12, "bold"), fg=RENK_VURGU,
        )
        self.saat_label.pack(padx=22, pady=9)

        self._saati_guncelle()
        self._metinleri_yenile()

    def _metinleri_yenile(self):
        """Yönetici, Genel Ayarlar'dan bu metinleri değiştirdiğinde (ya da
        uygulama ilk açıldığında) ana ekrandaki başlık/açıklama/sağ panel
        başlığını güncel değerleriyle gösterir."""
        baslik = self.db.ayar_getir("app_basligi", VARSAYILAN_APP_BASLIGI)
        # Pencere başlığı (işletim sistemi başlık çubuğu) tek satırdır; ekrandaki
        # etiket birden çok satırlı kalsın diye satır sonlarını burada boşlukla
        # değiştiriyoruz (aksi halde "...MÜDÜRLÜĞÜKURSİYER..." gibi birleşik görünür).
        self.title(baslik.replace("\n", " "))
        if hasattr(self, "baslik_label"):
            self.baslik_label.config(text=baslik)

        aciklama = self.db.ayar_getir("ana_aciklama", VARSAYILAN_ANA_ACIKLAMA)
        if hasattr(self, "aciklama_label"):
            self.aciklama_label.config(text=aciklama)

        sag_panel_basligi = self.db.ayar_getir("sag_panel_basligi", VARSAYILAN_SAG_PANEL_BASLIGI)
        if hasattr(self, "sag_panel_baslik_label"):
            self.sag_panel_baslik_label.config(text=sag_panel_basligi)

    # Geriye dönük uyumluluk: eski adla çağrılırsa da çalışsın.
    def _baslik_yenile(self):
        self._metinleri_yenile()

    def _grup_butonlarini_yenile(self):
        """Sol paneldeki grup butonlarını veritabanındaki güncel gruplarla
        yeniden oluşturur; yönetici Derslik Yönetimi'nden yeni bir grup
        ekleyince ya da bir grubun adını değiştirince, ana ekranı yeniden
        başlatmaya gerek kalmadan buradan çağrılır."""
        for w in self.grup_butonlari_cercevesi.winfo_children():
            w.destroy()
        for grup in self.db.gruplari_getir():
            tk.Button(
                self.grup_butonlari_cercevesi, text=grup["ad"], bg=RENK_SOL_BUTON, fg=RENK_METIN_ACIK,
                activebackground=RENK_VURGU, activeforeground="white",
                font=("Segoe UI", 11), relief="flat", bd=0, cursor="hand2",
                command=lambda g=grup: self._grup_giris(g),
            ).pack(fill="x", padx=18, pady=7, ipady=11)

    def _saati_guncelle(self):
        self.saat_label.config(
            text=datetime.now().strftime("%d.%m.%Y  %H:%M:%S") + f"  ({GUN_ADLARI[bugun_gun_index()]})"
        )
        self.after(1000, self._saati_guncelle)

    def _periyodik_kontrol(self):
        try:
            self.db.otomatik_devamsizlik_kontrolu()
        except Exception:
            pass
        self.after(30000, self._periyodik_kontrol)

    # ---- yardımcı: yönetici şifresi doğrulama ----
    def _admin_sifresini_dogrula(self, prompt_metni="Yönetici şifresini girin:"):
        admin_hash = self.db.ayar_getir("admin_sifre_hash")
        sifre = simpledialog.askstring(APP_ADI, prompt_metni, show="*", parent=self)
        if sifre is None:
            return False
        if sifre_dogrula(sifre, admin_hash):
            return True
        messagebox.showerror(APP_ADI, "Şifre hatalı.", parent=self)
        return False

    # ---- işlemler ----
    def _grup_giris(self, grup_row):
        sifre = simpledialog.askstring(
            APP_ADI, f"{grup_row['ad']} şifresini girin:", show="*", parent=self
        )
        if sifre is None:
            return
        admin_hash = self.db.ayar_getir("admin_sifre_hash")
        # Grubun kendi şifresi ya da yönetici şifresi ile girilebilir; böylece
        # yönetici tüm grup ekranlarına kendi şifresiyle de erişebilir.
        if sifre_dogrula(sifre, grup_row["sifre_hash"]) or sifre_dogrula(sifre, admin_hash):
            GrupEkrani(self, self.db, grup_row)
        else:
            messagebox.showerror(APP_ADI, "Şifre hatalı.", parent=self)

    def _yonetici_giris(self):
        if not self._admin_sifresini_dogrula("Yönetici şifresini girin:"):
            return
        YoneticiAnaPaneli(
            self, self.db, baslik_degisti_geri_cagirma=self._metinleri_yenile,
            gruplar_degisti_geri_cagirma=self._grup_butonlarini_yenile,
        )

    def _izleyici_giris(self):
        izleyici_hash = self.db.ayar_getir("izleyici_sifre_hash")
        if not izleyici_hash:
            messagebox.showerror(
                APP_ADI,
                "İzleyici şifresi henüz belirlenmemiş. Yönetici, Yönetici "
                "panelinden -> Şifreleri Yönet ile bu şifreyi belirleyebilir.",
                parent=self,
            )
            return
        sifre = simpledialog.askstring(
            APP_ADI, "İzleyici şifresini girin:", show="*", parent=self
        )
        if sifre is None:
            return
        admin_hash = self.db.ayar_getir("admin_sifre_hash")
        # İzleyici kendi şifresiyle ya da yönetici şifresiyle girebilir.
        if sifre_dogrula(sifre, izleyici_hash) or sifre_dogrula(sifre, admin_hash):
            YoneticiPaneli(self, self.db, salt_okunur=True)
        else:
            messagebox.showerror(APP_ADI, "Şifre hatalı.", parent=self)


# --------------------------------------------------------------------------
# Giriş noktası
# --------------------------------------------------------------------------

def _beklenmeyen_hata_goster(exc_type, exc_value, exc_tb):
    """Bir buton/olay islenirken beklenmeyen bir hata olursa, sessizce
    hicbir sey olmamis gibi davranmak yerine kullaniciya bir uyari
    penceresi gosterir (ozellikle konsolsuz .exe'de hatalar aksi halde
    tamamen görünmez olurdu)."""
    import traceback
    hata_metni = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    try:
        sys.stderr.write(hata_metni)
    except Exception:
        pass
    try:
        messagebox.showerror(
            APP_ADI,
            "Beklenmeyen bir hata oluştu, işlem tamamlanamadı:\n\n"
            f"{exc_type.__name__}: {exc_value}\n\n"
            "Lütfen bu penceredeki bilgileri kontrol edip tekrar deneyin. "
            "Sorun devam ederse bu hatayı geliştiriciye bildirin.",
        )
    except Exception:
        pass


def main():
    kok = tk.Tk()
    kok.report_callback_exception = _beklenmeyen_hata_goster
    kok.withdraw()

    db_yolu = veri_dosyasi_yolunu_al()
    if not db_yolu:
        db_yolu = ortak_klasoru_sec_ve_kaydet(kok)

    db = VeriTabani(db_yolu)
    try:
        db.kurulumu_baslat()
    except Exception as e:
        messagebox.showerror(
            APP_ADI,
            f"Veritabanı başlatılamadı:\n{db_yolu}\n\nHata: {e}\n\n"
            "Lütfen ayarlar.ini dosyasındaki veri klasörü yolunun doğru ve "
            "erişilebilir olduğundan emin olun.",
        )
        sys.exit(1)

    def uygulamayi_baslat():
        kok.destroy()
        pencere = AnaPencere(db)
        pencere.report_callback_exception = _beklenmeyen_hata_goster
        pencere.mainloop()

    if not db.kurulum_tamam_mi():
        KurulumSihirbazi(kok, db, uygulamayi_baslat)
        kok.deiconify()
        kok.withdraw()
        kok.mainloop()
    else:
        uygulamayi_baslat()


if __name__ == "__main__":
    main()

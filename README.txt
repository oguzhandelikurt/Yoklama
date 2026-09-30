YOKLAMA UYGULAMASI - KURULUM VE KULLANIM KILAVUZU
====================================================

Bu uygulama, kurumunuzdaki gruplar/derslikler için dijital yoklama
almanızı sağlayan, internet gerektirmeyen bir Windows masaüstü
uygulamasıdır. Veriler internete değil, kurumunuzun kendi ortak ağ
klasörüne kaydedilir.

Varsayılan olarak gelen gruplar:
  1) İngilizce
  2) Batı
  3) Balkan
  4) Doğu Dilleri
  5) Çağdaş Türk Lehçeleri

Derslikler (fiziksel sınıflar) gruplardan BAĞIMSIZ, kurumun TAMAMINA ait
ORTAK bir havuzdur: uygulama ilk kurulduğunda "Derslik 1"den "Derslik 40"a
kadar 40 adet derslik otomatik oluşturulur. Herhangi bir grubun kursu, bu
40 dersliğin herhangi birine atanabilir (ör. İngilizce grubunun bir kursu
Derslik 5'te, Doğu Dilleri grubunun başka bir kursu Derslik 12'de olabilir).
(Yönetici, Derslik Yönetimi ekranından yeni grup ekleyebilir/var olan
grupların adını değiştirebilir VE bu ortak derslik havuzuna yeni derslik
ekleyebilir, yeniden adlandırabilir veya kullanılmıyorsa silebilir; derslik
sayısında bir üst sınır yoktur.)


İÇİNDEKİ DOSYALAR
------------------
- yoklama_app.py   -> Uygulamanın kaynak kodu (Python)
- build_exe.bat    -> EXE dosyasını oluşturan derleme betiği (Windows)
- README.txt       -> Bu kılavuz


ADIM 1: EXE DOSYASINI OLUŞTURMA (SADECE 1 KERE)
--------------------------------------------------
Kurum bilgisayarlarınızın interneti kısıtlı olduğu ve Python kurulu
olmadığı için, uygulamayı tek bir ".exe" dosyası haline getirip öyle
dağıtmanız gerekiyor. Bu işlemi İNTERNETİ OLAN VE PYTHON KURULABİLEN
herhangi bir bilgisayarda (ev bilgisayarınız olabilir) SADECE BİR KERE
yapmanız yeterli.

1. Eğer o bilgisayarda Python yoksa, https://www.python.org/downloads/
   adresinden Python'u indirip kurun. Kurulum ekranında mutlaka
   "Add python.exe to PATH" kutusunu işaretleyin.
2. yoklama_app.py, build_exe.bat ve logo.ico dosyalarını aynı klasöre
   koyun (logo.ico, EXE'nin simgesi/ikonu olur; bu dosya olmadan da
   derleme çalışır, sadece EXE varsayılan simgeyle oluşur).
3. build_exe.bat dosyasına ÇİFT TIKLAYIN. İşlem birkaç dakika sürebilir
   (PyInstaller, Excel dosyası oluşturmak için gereken "openpyxl" ve
   yoklama panelindeki takvim/tarih seçici için gereken "tkcalendar"
   paketleri otomatik olarak indirilip kurulur).
4. İşlem bitince, aynı klasör içinde oluşan "dist" klasörünün içinde
   "Yoklama.exe" dosyasını bulacaksınız.

Bu "Yoklama.exe" dosyası tek başına çalışır; kopyalandığı bilgisayarda
Python veya internet GEREKMEZ.


ADIM 2: ORTAK AĞ KLASÖRÜ HAZIRLAMA
--------------------------------------
Tüm bilgisayarların aynı yoklama verisini görmesi için, kurumunuzun iç
ağındaki (internet değil, yerel ağ / sunucu) paylaşımlı bir klasöre
ihtiyacınız var. Örnek: \\SUNUCU\Yoklama gibi bir ağ klasörü, ya da
tüm bilgisayarların eriştiği paylaşımlı bir sürücü (Z:\Yoklama gibi).

- Bu klasörde tüm ilgili bilgisayarların OKUMA ve YAZMA izni olmalı.
- Bulut senkronizasyon klasörleri (OneDrive, Google Drive vb.) SQLite
  veritabanı dosyaları için ÖNERİLMEZ; senkronizasyon gecikmeleri veri
  çakışmalarına yol açabilir. Mümkünse gerçek bir yerel ağ paylaşımı
  (dosya sunucusu) kullanın.

NOT: Uygulama, veriyi bu klasördeki "yoklama.db" adlı tek bir dosyada
tutar. Çok yoğun aynı anda yazma olmadığı sürece (ki bu uygulamada
yoklama girişleri seyrek ve kısa sürelidir) bu yöntem sorunsuz çalışır.
Eğer "veritabanı kilitli" gibi hatalarla sıklıkla karşılaşırsanız, ağ
paylaşımınızın dosya kilitlemeyi (file locking) düzgün destekleyip
desteklemediğini BT ekibinizle kontrol edin.


ADIM 3: YOKLAMA.EXE DOSYASINI DAĞITMA VE İLK ÇALIŞTIRMA
------------------------------------------------------------
1. Oluşturduğunuz "Yoklama.exe" dosyasını, yoklama alınacak her
   bilgisayara kopyalayın (Masaüstü veya C:\Yoklama gibi normal,
   yazılabilir bir klasöre koymanız yeterli - exe'nin kendisi ağ
   klasöründe olmak ZORUNDA değil, sadece VERİ ağ klasöründe olacak).
2. İlk çalıştırdığınızda uygulama sizden ORTAK AĞ KLASÖRÜNÜ seçmenizi
   isteyecek (Adım 2'de hazırladığınız klasör). Bu seçim bir defaya
   mahsustur ve "ayarlar.ini" dosyasına kaydedilir (exe ile aynı
   klasörde oluşur).
3. Ardından bir KURULUM SİHİRBAZI açılır: burada 5 grup için ayrı
   ayrı giriş şifreleri, bir YÖNETİCİ şifresi ve isteğe bağlı bir
   İZLEYİCİ şifresi belirlersiniz (izleyici şifresini boş bırakırsanız
   sonra Yönetici panelinden -> Şifreleri Yönet ile ekleyebilirsiniz).
   Bu ekran sadece İLK çalıştırmada, verinin ilk oluşturulduğu anda bir
   kez çıkar (çünkü şifreler ortak veritabanında saklanır); diğer
   bilgisayarlarda aynı ağ klasörünü seçtiğiniz sürece tekrar sorulmaz.

Windows, tanınmayan bir uygulama çalıştırırken "Windows tarafından
korunan PC" gibi bir uyarı gösterebilir. Bu durumda "Daha fazla bilgi"
-> "Yine de çalıştır" seçeneğini kullanabilirsiniz. Kurumunuzda BT
politikası exe çalıştırmayı engelliyorsa BT ekibinizden izin almanız
gerekebilir.

NOT: Pencerelerin başlık çubuğu (simge/kapat düğmesinin olduğu en üst
şerit) uygulamanın kendi lacivert rengiyle boyanır. Bu özellik yalnızca
Windows 11'de (derleme 22000 ve üzeri) çalışır; Windows 10'da veya daha
eski sürümlerde desteklenmediği için başlık çubuğu Windows'un varsayılan
rengiyle görünür - bu bir hata değildir, uygulama yine sorunsuz çalışır.


KULLANIM
--------

EKRAN DÜZENİ:
- SOL taraf: grup butonları (varsayılan olarak İngilizce, Batı, Balkan,
  Doğu Dilleri, Çağdaş Türk Lehçeleri; yönetici yeni grup eklediyse
  listede o da görünür). Öğretmenler buradan kendi grubuna girer.
- SAĞ taraf: "İzleyici" (üstte) ve "Yönetici" (altta) butonları.
  İzleyici, sadece yoklama kayıtlarını görüntülemek isteyenler
  (örn. üst yönetim) içindir; Yönetici ise tüm yönetim işlemlerinin
  yapıldığı ana giriş noktasıdır.
- Ana ekranın üstündeki başlık metni ("Yoklama Uygulaması" yazısı)
  Yönetici panelinden -> Genel Ayarlar ile değiştirilebilir.

ÖĞRETMENLER İÇİN:
- Sol taraftaki gruplardan kendi grubunuzu seçin (örn. İngilizce).
- Grubunuzun şifresini girin (yönetici, kendi şifresiyle de herhangi
  bir grup ekranına girebilir).
- Açılan ekranda, grubunuzun O AN aktif olan kurslarının listesi görünür;
  her satırda KURS ADI ve DERSLİK birlikte gösterilir (ör. "İngilizce 101
  (Derslik 5)" - derslikler kurumun ortak havuzundan geldiği için
  numaraları gruptan gruba farklı olabilir). İçinde ders vereceğiniz
  kursa/dersliğe tıklayarak girin.
- Derslik ekranında üstte Pazartesi-Cuma gün sekmeleri bulunur;
  varsayılan olarak BUGÜNÜN periyotları (saatleri) listelenir. Diğer
  günlere tıklayarak o günün programını bilgi amaçlı görebilirsiniz
  (yoklama SADECE bugünün periyotları için alınabilir; kurs, kendisine
  tanımlanmış tarih aralığı içindeyse listelenir).
- Bir dersliğe atanan kurs HER GÜN aynı dersliktedir ve her gün 6
  PERİYODA (6 ayrı saate) sahiptir; yani bir kurs için günde 6 KEZ
  ayrı ayrı yoklama alınır. Pazartesi-Perşembe günlerinin periyot
  saatleri aynıdır, Cuma'nın saatleri ayrı tanımlanmış olabilir.
- İlgili periyodun saati geldiğinde ve başlangıçtan itibaren yönetici
  tarafından o kurs için belirlenmiş süre içindeyse, o periyoda ait
  "Yoklama Al" butonu aktif olur. Bu butona basıp kursiyerlerin
  karşısındaki kutucukları işaretleyerek (işaretli=Var, işaretsiz=Yok)
  "Yoklamayı Kaydet" ile onaylayın. Aynı gün içinde sırayla 1. Saat,
  2. Saat ... 6. Saat için ayrı ayrı "Yoklama Al" işlemi yaparsınız.
- Bir kursiyeri "Yok" (devamsız) işaretlediğinizde, satırın sağında bir
  mazeret listesi belirir (yöneticinin önceden tanımladığı "Raporlu",
  "Revir", "Sevk" gibi seçenekler); uygunsa oradan seçim yapabilirsiniz
  (mazeret seçmek zorunlu değildir).
- Süre dolmadığı sürece o periyodun yoklamasını istediğiniz kadar
  tekrar açıp düzeltebilirsiniz ("Yoklamayı Düzenle" butonu); Excel'e
  aktarıldığında en SON kaydettiğiniz zaman "Giriş Zamanı" olarak
  görünür. Süre dolduktan sonra kayıt kilitlenir, değişiklik için
  yöneticiye başvurulması gerekir.
- Eğer süre içinde bir periyot için hiç yoklama girilmezse, sistem o
  periyodu otomatik olarak "Öğretmen girmedi" şeklinde işaretler ve o
  periyottaki TÜM kursiyerleri devamsız (Yok) sayar; diğer periyotlar
  bundan etkilenmez, düzeltme için yöneticiye başvurulmalıdır.

YÖNETİCİ PANELİ (sağdaki "Yönetici" butonu):
- Yönetici şifresini bir kere girdikten sonra açılan panelde aşağıdaki
  bölümlerin TÜMÜNE tekrar şifre girmeden erişebilirsiniz:

  1) YOKLAMA (Görüntüle / Excel'e Aktar): Seçilen tarih için TÜM
     kursların TÜM periyotlarının (o gün planlanmış her 6 saatin ayrı
     bir satır olduğu) durumunu (yoklama alındı / öğretmen girmedi /
     bekliyor), son güncelleme zamanını, toplam kişi, HAZIR (var) ve
     devamsız (yok) sayılarını gösterir; "Saat" sütununda "1. Saat
     08:00-08:50" gibi periyot bilgisi de yer alır. "◀ Önceki Gün" /
     "Sonraki Gün ▶" butonları veya tarih kutusuna elle yazarak GEÇMİŞ
     HAFTA/AY kayıtlarını da görüntüleyebilirsiniz. Tarih kutusuna
     tıkladığınızda bir takvim açılır ve tarihi oradan da
     seçebilirsiniz (bu özellik "tkcalendar" paketiyle çalışır; EXE,
     build_exe.bat ile derlendiği sürece otomatik olarak dahil edilir).
     Tek bir gün yerine BELİRLİ BİR TARİH ARALIĞINDA girilen TÜM
     kursların yoklamasını görmek için, üstteki "Tarih aralığı" satırında
     başlangıç ve bitiş tarihlerini seçip "Aralığı Göster" butonuna
     basın; tablo o aralıktaki her günün her periyodunu "Tarih" sütunuyla
     birlikte listeler (en fazla 400 günlük bir aralık seçilebilir).
     "Excel'e Aktar (.xlsx)" butonu HER ZAMAN o an ekranda görünen
     kayıtları (tek gün ya da seçili tarih aralığı, hangisi gösteriliyorsa)
     2 sayfalı gerçek bir Excel dosyası olarak kaydeder:
       * "Özet" sayfası: her periyodun tarihini, toplam/hazır/devamsız
         sayılarını ve devamsız kursiyerlerin kodu, adı soyadı VE
         MAZERETİNİ ("101 - Ahmet Yılmaz (Raporlu)" şeklinde) tek
         sütunda, her kursiyer kendi satırında listeler.
       * "Detay (Kursiyer)" sayfası: her kursiyerin tarihini, kodu, adı
         soyadı, Var/Yok durumu ve mazereti ayrı sütunlarda, satır
         satır listelenir.
     Giriş Zamanı sütunu metin olarak biçimlendirilmiştir; Excel'de
     "#####" şeklinde görünmez. Tüm Türkçe karakterler doğru görünür.

  2) KURS EKLE: Grup, DERSLİK (kurumun TAMAMINA ait ORTAK derslik
     havuzunun TÜM listesi gösterilir - ör. 40 derslik varsa "Derslik 1"den
     "Derslik 40"a kadar hepsi listelenir, Grup ne olursa olsun; seçtiğiniz
     derslik artık o kursa özel/tanımlı olur), kurs adı girilir. Ardından
     HAFTALIK DERS PROGRAMI belirlenir: bir kurs, atandığı derslikte HER
     GÜN vardır ve her gün 6 PERİYODU (6 ayrı saati) vardır - yani o kurs
     için öğretmen günde 6 kez ayrı yoklama alır. "Pazartesi - Salı -
     Çarşamba - Perşembe" başlığı altında bu 4 günde ORTAK olacak 6
     periyodun başlangıç/bitiş saatleri (HH:SS, ör. 08:00-08:50) girilir;
     "Cuma" başlığı altında ise Cuma günü için AYRI 6 periyodun saatleri
     girilir (varsayılan olarak öğle arası 4. ve 5. periyot arasına
     gelecek şekilde saatler önceden doldurulmuş gelir, dilerseniz
     değiştirebilirsiniz). Bu program HER HAFTA aynı gün/aynı saatte
     tekrar eder (haftadan haftaya tarih bazlı istisna girilmez). Son
     olarak İSTEĞE BAĞLI başlangıç/bitiş TARİHİ (GG.AA.YYYY; boş
     bırakılırsa kurs sınırsız sürer - bir kursun 2-3 ay sürecekse buraya
     o tarih aralığını girin) ve o kurs için YOKLAMA GİRME SÜRESİ
     (dakika, varsayılan 5) girilir.

  3) KURSİYER EKLE: Grup ve kurs seçilir, "Kaç kursiyer
     ekleyeceksiniz?" kutusuna sayı girilip "Oluştur"a basılır; o kadar
     Kod / Ad Soyad satırı açılır. AD SOYAD ARTIK ZORUNLU DEĞİLDİR -
     sadece kod girmeniz yeterlidir. "Hepsini Kaydet" ile kaydedilir;
     kod eksik veya tekrar eden satırlar atlanıp özet mesajında
     bildirilir. Bir kursun kursiyer listesi, o kursun TÜM periyotları
     için ortaktır (aynı kursiyerler günde 6 periyoda da girer).

  4) KURS / KURSİYER DÜZENLE: Yanlış girilmiş bir kursu (ad, DERSLİK -
     ortak 40'lık havuzun tam listesinden değiştirilebilir -, haftalık
     program, tarih aralığı, yoklama süresi - Kurs Ekle'dekiyle aynı
     haftalık program editörüyle) veya kursiyeri (kod, ad soyad)
     düzeltebilir ya da tamamen silebilirsiniz. Bir kursu silmek, ona
     kayıtlı tüm kursiyerleri ve yoklama geçmişini de siler - onay
     istenir.

  5) DERSLİK YÖNETİMİ: Burada iki ayrı/bağımsız şey yapılır:
     a) Grup yönetimi: "Yeni Grup Ekle" ile listeye yeni bir grup
        eklenebilir (yeni grubun giriş şifresini daha sonra "Şifreleri
        Yönet" ekranından belirlemeniz gerekir), "Grubu Yeniden
        Adlandır" ile var olan bir grubun adı değiştirilebilir. Bir grup
        silindiğinde ortak derslik havuzu ETKİLENMEZ.
     b) Derslik yönetimi: Burada listelenen derslikler GRUPTAN BAĞIMSIZ,
        kurumun TAMAMINA ait ORTAK bir havuzdur (varsayılan: "Derslik 1"
        .. "Derslik 40"). Buradan yeni derslik ekleyebilir, seçili
        dersliği yeniden adlandırabilir veya (hiçbir kurs tarafından
        kullanılmıyorsa) silebilirsiniz. Derslik sayısında bir üst sınır
        yoktur.

  6) MAZERET YÖNETİMİ: Devamsızlık mazeret listesini (Raporlu, Revir,
     Sevk gibi) burada oluşturur, düzenler veya silersiniz. Öğretmenler
     yoklama alırken bir kursiyeri "Yok" işaretlediğinde bu listeden
     seçim yapar. İstediğiniz kadar mazeret tanımlayabilir, sonradan
     ekleyip çıkarabilirsiniz.

  7) ŞİFRELERİ YÖNET: Grup şifrelerini, yönetici şifresini ve izleyici
     şifresini istediğiniz zaman değiştirebilirsiniz.

  8) GENEL AYARLAR: Ana ekranın üstünde görünen uygulama başlığını
     değiştirebilirsiniz.

İZLEYİCİ (sağdaki "İzleyici" butonu):
- Kendi ayrı bir şifresi vardır (Yönetici -> Şifreleri Yönet'ten
  belirlenir); İzleyici ekranına YÖNETİCİ ŞİFRESİYLE de girilebilir.
  İzleyici, Yoklama panelini (görüntüleme + Excel'e aktarma) açabilir
  ama şifreleri değiştiremez ve diğer yönetim bölümlerine (Kurs Ekle,
  Kursiyer Ekle, Düzenle, Derslik/Mazeret Yönetimi) erişemez.


BİLİNMESİ GEREKENLER / SINIRLAMALAR
------------------------------------
- ÖNEMLİ (bir önceki sürümden güncelleyenler için): Bu sürümle birlikte
  kurs modeli "günde tek saat" yerine "her gün 6 periyotluk haftalık
  program"a geçmiştir. Uygulama ilk açıldığında bunu otomatik algılar
  ve eski yöntemle girilmiş TÜM kursları, kursiyerleri ve yoklama
  geçmişini KALICI OLARAK SİLER; GRUP tanımlarınız (ve genel ayarlar/
  mazeret listeniz) korunur. Silme sonrası kursları yeni Kurs Ekle
  ekranından yeniden tanımlamanız gerekir.
- ÖNEMLİ (derslik modeli değişikliği): Derslikler artık her grubun kendi
  sınıfları DEĞİL, kurumun TAMAMINA ait ORTAK bir havuzdur (varsayılan:
  "Derslik 1".."Derslik 40"). Eski (gruba özel) dersliklerle kurulmuş bir
  veritabanı açıldığında, uygulama bunu da otomatik algılar; eski
  gruba-özel derslik kayıtlarını (ve onlara bağlı kurs/kursiyer/yoklama
  verilerini) KALICI OLARAK SİLİP yerine ortak "Derslik 1".."Derslik 40"
  havuzunu oluşturur. Bu geçişten sonra da kursları Kurs Ekle ekranından
  yeniden tanımlamanız gerekir.
- Kursiyer kodları elle girilir, sistem otomatik kod üretmez.
- Kursiyer ad soyadı artık opsiyoneldir; sadece kod girilerek de
  kursiyer eklenebilir (ekranlarda ad soyad boşsa "-" gösterilir).
- Veritabanı ortak ağ klasöründe SQLite dosyası olarak tutulur. Çok
  sayıda bilgisayarın TAM AYNI ANDA yazma yapması (aynı saniyede
  birden fazla "Yoklamayı Kaydet" tıklaması) nadiren geçici bir
  "database is locked" hatasına yol açabilirse de, uygulama bunu
  otomatik olarak birkaç kez tekrar deneyerek aşmaya çalışır.
- Bir kursu veya kursiyeri silmek GERİ ALINAMAZ; "Kurs/Kursiyer
  Düzenle" ekranında bu yüzden her silme işleminden önce onay
  istenir.

# Ürün Görüntüleyici (ExcelViewer)

Excel (`.xlsx`) tabanlı ürün/stok listelerini görüntülemek, aramak, filtrelemek, sıralamak, düzenlemek ve iki dosyayı birleştirmek için geliştirilmiş bir Windows masaüstü (WinForms) uygulaması. Uygulama bir Excel dosyasını açar, içindeki ürün satırlarını bir tabloda gösterir ve yapılan ekleme/düzenleme/silme işlemlerini doğrudan aynı dosyaya yazar.

## 1) Proje Tanımı

Uygulama bir `.xlsx` dosyasındaki ürün kayıtlarını okuyup tablo (DataGridView) olarak gösterir. Beş mantıksal alan üzerinde çalışır: **Ürün Kodu, Stok Adeti, Birim, Fiyat, Para Birimi**. Temel işlevler:

- Excel dosyası açma ve görüntüleme
- Ürün kodu / birim üzerinden arama
- Stok ve fiyat aralığı ile birim/para birimi seçimine göre filtreleme
- İstenen kolona göre artan/azalan sıralama
- Ürün ekleme, düzenleme ve silme (değişiklikler anında dosyaya yazılır)
- Görünen listeyi yeni bir Excel dosyasına dışa aktarma
- İki Excel dosyasını Ürün Kodu üzerinden birleştirme
- Her yazma işlemini bir metin log dosyasına kaydetme

Ürün Kodu dışındaki alanların dosyada mevcut olması zorunlu değildir; eksik kolonlar tabloda gizlenir ve o alanlara veri yazılması engellenir.

## 2) Gereksinimler

- **.NET SDK:** .NET 10 (`net10.0-windows` hedef çerçevesi — bkz. `ExcelViewer/ExcelViewer.csproj`)
- **İşletim sistemi:** Windows (uygulama WinForms — `UseWindowsForms` — kullanır; Windows dışı platformlarda çalışmaz)
- **NuGet paketleri:**
  - [ClosedXML](https://www.nuget.org/packages/ClosedXML) `0.104.2` — `.xlsx` dosyalarının okunması/yazılması
  - [Guna.UI2.WinForms](https://www.nuget.org/packages/Guna.UI2.WinForms) `2.0.4.7` — modern görünümlü arayüz kontrolleri

Projede `Nullable` ve `ImplicitUsings` etkindir ve `TreatWarningsAsErrors` açıktır (uyarılar derlemeyi durdurur).

## 3) Kurulum ve Çalıştırma

Proje kökünde `ExcelViewer.sln` çözüm dosyası bulunur.

### dotnet CLI ile

```bash
# Bağımlılıkları geri yükle
dotnet restore

# Derle
dotnet build

# Çalıştır (Windows üzerinde)
dotnet run --project ExcelViewer/ExcelViewer.csproj
```

### Rider / Visual Studio ile

1. `ExcelViewer.sln` dosyasını açın.
2. IDE NuGet paketlerini otomatik geri yükler (gerekirse `dotnet restore` çalıştırın).
3. `ExcelViewer` projesini başlangıç projesi olarak seçip çalıştırın (Run/Debug).

> Not: Hedef çerçeve `net10.0-windows` olduğu için derleme ve çalıştırma Windows üzerinde yapılmalıdır.

## 4) Yayınlama (Tek Dosya .exe Oluşturma)

Uygulamayı başka bir bilgisayara dağıtmak için tek bir `.exe` dosyası olarak yayınlayabilirsiniz. Aşağıdaki komut, `.NET`'in **hedef makinede kurulu olmasını gerektirmeyen** bağımsız (self-contained) tek dosyalık bir çıktı üretir:

```bash
dotnet publish ExcelViewer/ExcelViewer.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

Bu komut test edilip başarıyla çalıştığı doğrulanmıştır (yukarıdaki parametreler .NET 10 için geçerlidir).

- **Çıktı konumu:** `ExcelViewer/bin/Release/net10.0-windows/win-x64/publish/`
- **Üretilen dosya:** `ExcelViewer.exe`
- **.NET kurulumu gerekli mi?** Hayır. `--self-contained true` kullanıldığı için gerekli .NET çalışma zamanı `.exe` içine gömülür; hedef bilgisayarda ayrıca .NET kurulu olması gerekmez. (Çalışma zamanını gömmeyen, boyutu daha küçük ama hedefte .NET 10 gerektiren bir çıktı için `--self-contained false` kullanılabilir.)

> `-r win-x64` çalışma zamanı tanımlayıcısı 64-bit Windows içindir; 32-bit hedef için `win-x86` kullanılabilir.

## 5) Excel Dosya Formatı

Uygulama, açtığı dosyanın **ilk çalışma sayfasını** okur. Sayfanın en fazla ilk **10 dolu satırı** içinde başlık satırını otomatik tespit eder; böylece başlıktan önce logo, rapor başlığı, tarih veya özet satırları bulunabilir. Bu satırlar içinden **en çok mantıksal kolona eşlenen** satır başlık kabul edilir (eşitlik durumunda daha çok dolu hücreye sahip olan seçilir).

### Beklenen kolonlar ve kabul edilen alternatif başlıklar

Başlık eşleştirme büyük/küçük harf duyarsızdır. Önce **tam eşleşme** (alias listesi) denenir; bulunamayan kolonlar için başlığın bir anahtar kelimeyi **içermesi** (substring) yeterlidir.

| Mantıksal Alan | Kabul edilen başlıklar (tam eşleşme alias'ları) |
|---|---|
| **Ürün Kodu** (zorunlu) | Ürün Kodu, UrunKodu, Ürün Kod, Kod, Product Code, ProductCode, Code, Stock Code, StockCode, Model |
| **Stok Adeti** | Stok Adeti, Stok Adedi, StokAdeti, Stok, Adet, Miktar, Stock, Quantity, Qty, Amount, Count, Total |
| **Birim** | Birim, Ölçü, Ölçü Birimi, Unit, Measure, UOM |
| **Fiyat** | Fiyat, Ücret, Tutar, Price, Cost, Amount Price |
| **Para Birimi** | Para Birimi, ParaBirimi, Döviz, Kur, Currency, Curr |

Tam eşleşme bulunamazsa "içerir" (substring) yedeği devreye girer; örneğin `2026 Unit Price (euro)` başlığı `price` içerdiği için **Fiyat** kolonuna aday olur. Aynı hücre hem Birim hem Fiyat anahtar kelimesini içeriyorsa (örn. `Unit Price`), altındaki verinin sayısal mı metin mi olduğuna bakılır: veri metinse **Birim**, sayısalsa (veya karar verilemezse) **Fiyat** atanır.

### Zorunlu alan

**Ürün Kodu** çapa (anahtar) alandır; arama, çakışma kontrolü ve satır bulma bu alana dayanır. Dosyada tanınabilir bir Ürün Kodu kolonu bulunamazsa dosya okunamaz ve hata mesajı gösterilir.

### Kısmi kolonlu dosyalar

Beş alanın tamamının bulunması gerekmez. Bulunamayan alanlar:

- Tabloda **gizlenir** (yalnızca gerçekten var olan kolonlar gösterilir).
- İlk açılışta bir bilgi mesajıyla ("... sütunları şunlarla eşleşmediği için eklenmemiştir: ...") kullanıcıya bildirilir; hangi Excel başlıklarının eşleşmediği de listelenir.
- Bu alanlara veri yazılması engellenir: kullanıcı olmayan bir alana değer girmişse kayıt kısmen değil, tamamen iptal edilir ve uyarı verilir.

Sayısal alanlar (Stok Adeti, Fiyat) hem gerçek Excel sayısı hem de metin olarak yazılmış sayı biçiminde okunabilir (kültürden bağımsız ayrıştırma). Ürün Kodu ve Birim'in ikisi de boş olan satırlar okunurken atlanır.

## 6) Özellikler

### Excel açma ve görüntüleme
**Dosya ▾ → Excel Aç** ile bir `.xlsx` seçilir. Dosya arka planda okunur (arayüz donmaz) ve ürünler tabloya yüklenir. Dosya Excel'de açık olsa bile okuma yapılabilir (`FileShare.ReadWrite`). Durum çubuğu yüklenen ürün sayısını gösterir.

### Arama
Üstteki arama kutusuna yazılan metin **Ürün Kodu** veya **Birim** alanında geçen ürünleri süzer. Arama büyük/küçük harf duyarsızdır ve yazarken ~200 ms gecikmeli (debounce) çalışır; her tuşta değil, yazma durunca uygulanır.

### Filtreleme
**Filtre** butonuyla açılan panelde:

- Stok Adeti alt/üst sınırı
- Fiyat alt/üst sınırı
- Görünecek birim(ler) seçimi
- Görünecek para birim(ler)i seçimi

Boş bırakılan sınır "sınır yok", boş bırakılan seçim kümesi "hepsi görünsün" anlamına gelir. Filtre aktifken buton üzerinde bir işaret (●) belirir. Filtre durumu panel kapansa da korunur.

### Sıralama
**Sırala** butonuyla bir kolon (Ürün Kodu, Stok Adeti, Birim, Fiyat, Para Birimi) ve yön (artan/azalan) seçilir. Metin alanları büyük/küçük harf duyarsız sıralanır. Sıralama aktifken buton üzerinde işaret (●) görünür ve durum korunur.

> Arama, filtre ve sıralama sırayla uygulanır: **önce arama, sonra filtre, en son sıralama.** Hepsi bellek üzerinde çalışır; dosya yalnızca bir kez okunur.

### Ürün ekleme
**Ürün Ekle** formu Ürün Kodu, Stok Adeti + Birim, Fiyat + Para Birimi alanlarını toplar. Birim ve Para Birimi combo'ları hem sabit çekirdek değerlerle (Birim: `m`, `adet`; Para Birimi: `₺ $ € £ ¥`) hem de dosyadaki mevcut değerlerle beslenir ve kullanıcı yeni değer de yazabilir.

- Ekleme modunda girilen kod mevcut bir ürüne aitse **Birim ve Para Birimi otomatik doldurulup kilitlenir.**
- Kaydederken aynı Ürün Kodu dosyada zaten varsa **çakışma çözümü** sorulur:
  - **Evet →** Stokları güncelle (mevcut stok yeni değerle değiştirilir)
  - **Hayır →** Olan stokun üzerine ekle (mevcut stok + yeni stok)
  - **İptal →** İşlemi iptal et
- Stok Adeti girişinde birim `adet` ise yalnızca tam sayı; diğer birimlerde ondalıklı değer kabul edilir. Negatif stok ve negatif fiyat reddedilir.

### Ürün düzenleme ve silme
Tablo satırına **çift tıklamak** veya sağ tık menüsünden **Düzenle** seçmek düzenleme formunu seçili satırın verisiyle dolu açar; Kaydet mevcut satırı günceller (yeni satır eklemez). Düzenleme modunda, ekleme modundaki "kod mevcut bir ürünle eşleşirse Birim/Para Birimi otomatik dolup kilitlenir" davranışı **uygulanmaz** — Ürün Kodu dâhil tüm alanlar serbestçe değiştirilebilir. Sağ tık menüsünden **Sil**, onay sorduktan sonra satırı dosyadan kaldırır. Her iki işlem sonrası dosya yeniden okunup tablo yenilenir.

### Stok/Fiyat alanlarında +/- aritmetik ifade
Stok Adeti ve Fiyat alanlarına düz sayı yerine `+`/`-` içeren basit bir ifade yazılabilir (örn. `120-16+5`). İfade **yalnızca Kaydet'e basıldığında** soldan sağa hesaplanır (`120-16` → `104`); kullanıcı formda gezinirken alan içeriği değişmez. Baştaki `-` negatif işaret sayılır. Çarpma/bölme/parantez desteklenmez; yarım kalmış (`120-`) veya bozuk (`120--16`) ifadeler geçersizdir ve uyarı verir.

### Dışa aktarma
**Dosya ▾ → Dışa Aktar**, tabloda o an görünen (arama + filtre + sıralama uygulanmış) listeyi seçilen konuma **yeni** bir `.xlsx` dosyası olarak yazar. Orijinal dosyaya dokunulmaz; çıktı sabit Türkçe başlıklarla (Ürün Kodu, Stok Adeti, Birim, Fiyat, Para Birimi) oluşturulur, başlık satırı kalın ve dondurulmuş olur.

### Birleştirme
**Dosya ▾ → Birleştir**, açık dosya (A) ile seçilen ikinci bir Excel dosyasını (B) **Ürün Kodu** üzerinden birleştirir:

- A'da eksik olan bir kolon B'de varsa, o kolon B'den doldurulur.
- B'de olup A'da bulunmayan ürünler yeni satır olarak eklenir.
- Bir kolon hem A hem B'de varsa (**çakışma**) kullanıcıya kolon bazında karar sorulur: **A'yı kullan / B'yi kullan / Topla (yalnızca Stok Adeti için)**. Karar tüm satırlara aynı biçimde uygulanır. Çakışma panelinde her kolon için **varsayılan seçim "B'yi kullan"dır** (kullanıcı istediği gibi değiştirebilir).
- Eşleştirme anahtarı Ürün Kodu'dur (trim + büyük/küçük harf duyarsız).

Sonuç **yeni** bir dosyaya yazılır; A ve B'ye dokunulmaz. Kaydetme sonrası ekranda birleşim sonucu gösterilir ve kaç ürünün eksik bilgisinin doldurulduğu, kaç yeni ürün eklendiği özetlenir.

Kullanıcı çakışma diyaloğunda **İptal** seçerse (veya pencereyi kapatırsa) birleştirme durdurulur: **hiçbir dosya oluşturulmaz**, durum çubuğunda "Birleştirme iptal edildi." yazar ve ayrıca aynı bilgiyi veren bir uyarı penceresi (MessageBox) gösterilir.

### Değişiklik logu
Ekleme, güncelleme, düzenleme ve silme işlemleri zaman damgalı olarak bir metin dosyasına yazılır:

- **Konum:** Excel dosyasının bulunduğu klasörün altındaki `log` alt klasörü
- **Dosya:** `log/urun_log.txt`
- **Biçim:** `[yyyy-MM-dd HH:mm:ss] EKLEME | Kod=... | Stok=... | Birim=... | Fiyat=... | ParaBirimi=...` (işlem türüne göre EKLEME / GÜNCELLEME / DÜZENLEME / SİLME satırları)

Log klasörü yoksa otomatik oluşturulur ve her işlem dosyanın sonuna eklenir.

### Arayüz ve tablo davranışı
- **Dinamik üst bar:** Pencere yeniden boyutlandırıldığında arama kutusu pencereyle birlikte genişler/daralır; **Filtre** ve **Sırala** butonları sağ kenara sabit kalır (arama kutusu büyürken üzerlerine binmez). Pencerenin bir alt sınırı (`MinimumSize`) tanımlıdır; böylece kontroller üst üste binecek kadar küçültülemez.
- **Sabit satır yüksekliği:** Tablo satırlarının yüksekliği sabittir; kullanıcı satırları fareyle sürükleyerek büyütemez.
- **Kolon genişliği hafızası:** Kullanıcının elle değiştirdiği kolon genişlikleri, **aynı dosya** üzerinde arama/filtre/sıralama/ekleme/düzenleme/silme yapılsa bile korunur. Genişlikler yalnızca **yeni bir dosya açıldığında** varsayılan oranlara döner.

## 7) Bilinen Sınırlamalar

- **Yalnızca ilk çalışma sayfası** okunur; çok sayfalı dosyalarda diğer sayfalar yok sayılır.
- Yalnızca `.xlsx` desteklenir (`.xls`, `.csv` vb. yoktur).
- Başlık satırı, dosyanın ilk **10 dolu satırı** içinde aranır; başlık daha aşağıdaysa bulunamaz.
- **Ürün Kodu kolonu zorunludur;** tanınamazsa dosya hiç açılmaz.
- Dosya başka bir programda (ör. Excel) açıkken **yazma** işlemleri yapılamaz; bu durumda "dosya kilitli" uyarısı gösterilir ve girilen bilgiler korunur (okuma ise mümkündür).
- **+/- ifadeleri** yalnızca toplama/çıkarma içindir; çarpma, bölme ve parantez desteklenmez.
- **Dosya menüsündeki "Kaydet" öğesi pasif bir yer tutucudur** — uygulama her ekleme/düzenleme/silme işlemini zaten anlık olarak dosyaya yazdığı için ayrı bir kaydetme adımı yoktur.
- Birleştirmede çakışan kolon kararı **satır bazında değil, kolon bazındadır** (bir kolon için verilen karar o kolonun tüm satırlarına uygulanır); metin kolonlarında "Topla" seçeneği yoktur.
- Dışa aktarma ve birleştirme çıktıları sabit Türkçe başlıklarla üretilir; kaynak dosyanın orijinal başlık adları korunmaz.

## 8) Proje Yapısı

```
ExcelViewer/
├── ExcelViewer.sln
└── ExcelViewer/
    ├── ExcelViewer.csproj        # Hedef çerçeve ve NuGet paketleri
    ├── Program.cs                # Uygulama giriş noktası (MainForm'u başlatır)
    ├── MainForm.cs               # Ana pencere: menü, tablo, arama/filtre/sıralama, ekle/düzenle/sil akışları
    ├── Models/                   # Veri modelleri ve enum'lar (davranış içermez)
    │   ├── Urun.cs               # Ürün satırı modeli (5 alan)
    │   ├── FiltreKriteri.cs      # Filtre durumu (stok/fiyat aralığı, seçili birim/para birimi)
    │   ├── SiralamaKriteri.cs    # Sıralama alanı + yönü (enum'lar dâhil)
    │   ├── CakismaCozumu.cs      # Aynı kod çakışmasında çözüm yolu (enum)
    │   └── BirlestirmeModelleri.cs # Birleştirme kararları ve sonucu
    ├── Services/                 # İş mantığı (arayüzden bağımsız)
    │   ├── ColumnResolver.cs     # Başlık tespiti ve kolon eşleştirme (alias + substring fallback)
    │   ├── ExcelReaderService.cs # .xlsx okuma → bellek içi ürün listesi
    │   ├── ExcelWriterService.cs # Ekle/güncelle/sil işlemlerini dosyaya yazma
    │   ├── ExcelExportService.cs # Görünen listeyi yeni .xlsx'e dışa aktarma
    │   ├── ProductMergeService.cs # İki listeyi Ürün Kodu üzerinden birleştirme
    │   ├── UrunFilterService.cs  # Arama + filtre + sıralama (bellek üzerinde)
    │   └── LogService.cs         # İşlemleri log/urun_log.txt'ye yazma
    └── Forms/                    # Yardımcı diyaloglar ve paneller
        ├── AddProductForm.cs     # Ürün ekleme/düzenleme formu (+/- ifade hesaplama)
        ├── FiltrePanel.cs        # Filtre kriterleri paneli
        ├── SiralamaPanel.cs      # Sıralama seçim paneli
        └── BirlestirmePanel.cs   # Çakışan kolon kararları paneli
```

**Katman özeti:**

- **Models/** — Yalnızca veri taşıyan sınıflar ve enum'lar. Davranış barındırmaz.
- **Services/** — Excel okuma/yazma, kolon çözümleme, filtreleme, birleştirme ve loglama gibi tüm iş mantığı. Arayüzden bağımsızdır ve I/O işlemleri arka plan thread'inde çalışır.
- **Forms/** — Ana pencere dışındaki diyaloglar ve açılır paneller (ürün ekle/düzenle, filtre, sıralama, birleştirme kararları).

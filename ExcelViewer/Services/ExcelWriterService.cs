using System.Globalization;
using ClosedXML.Excel;
using ExcelViewer.Models;

namespace ExcelViewer.Services;

/// <summary>
/// Ürünleri mevcut .xlsx dosyasına yazar: yeni satır ekler veya
/// aynı Ürün Kodu varsa kullanıcının seçtiği çözüm yoluna göre günceller.
/// Her işlem LogService aracılığıyla zaman damgalı olarak loglanır.
/// </summary>
public sealed class ExcelWriterService
{
    private readonly LogService _logService;

    public ExcelWriterService(LogService logService)
    {
        ArgumentNullException.ThrowIfNull(logService);
        _logService = logService;
    }

    /// <summary>
    /// Ürünü Excel'e yazar. Aynı Ürün Kodu yoksa yeni satır olarak ekler.
    /// Varsa <paramref name="cozum"/> değerine göre davranır. Çakışma olduğunda
    /// ve çözüm henüz Iptal ise, çağıran tarafa çakışmayı bildirmek için
    /// <see cref="YazmaSonucu.CakismaVar"/> döner (dosyaya dokunmadan).
    /// </summary>
    /// <remarks>Bloklayan I/O arka plan thread'inde çalışır.</remarks>
    public Task<YazmaSonucu> UrunEkleAsync(
        string filePath,
        Urun urun,
        CakismaCozumu cozum,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(urun);
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("Dosya yolu boş olamaz.", nameof(filePath));

        return Task.Run(() => UrunEkle(filePath, urun, cozum, cancellationToken), cancellationToken);
    }

    private YazmaSonucu UrunEkle(
        string filePath, Urun urun, CakismaCozumu cozum, CancellationToken cancellationToken)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel dosyası bulunamadı.", filePath);

        cancellationToken.ThrowIfCancellationRequested();

        // Not: Dosya Excel'de açıksa Save() bir IOException fırlatır; bu, çağıran
        // tarafta yakalanıp kullanıcıya "dosyayı kapatın" uyarısı olarak gösterilir.
        using var workbook = new XLWorkbook(filePath);
        IXLWorksheet worksheet = workbook.Worksheets.First();

        IXLRange? usedRange = worksheet.RangeUsed();
        if (usedRange == null)
            throw new InvalidOperationException("Excel sayfası boş; başlık satırı bulunamadı.");

        // Okuma tarafıyla aynı tespit: ilk 10 satır içinde en zengin başlık satırı
        // seçilir (tam eşleşme + substring fallback + çoklu anahtar kelime çakışması).
        List<IXLRangeRow> rows = usedRange.RowsUsed().ToList();
        (ColumnMap columns, int headerIndex) = ResolveColumns(rows);

        // Eksik kolona yazma engeli: kullanıcı bir alana değer girdiyse ama o kolon
        // dosyada yoksa, kısmi kayıt yerine işlemi tamamen iptal et ve uyar.
        EksikKolonaYazmayiEngelle(urun, columns);

        // Aynı Ürün Koduna sahip mevcut satırı ara (başlık ve üstündeki satırları atla).
        IXLRangeRow? mevcutSatir = FindRowByCode(rows, headerIndex, columns.UrunKodu, urun.UrunKodu);

        if (mevcutSatir != null)
        {
            // Çakışma var. Kullanıcı henüz karar vermediyse çağıran tarafa bildir.
            if (cozum == CakismaCozumu.Iptal)
                return YazmaSonucu.CakismaVar;

            decimal eskiStok = ReadDecimal(mevcutSatir, columns.StokAdeti);
            decimal yeniStok = cozum == CakismaCozumu.UzerineEkle
                ? eskiStok + urun.StokAdeti
                : urun.StokAdeti; // StoklariGuncelle: yeni değerle değiştir.

            WriteDecimal(mevcutSatir, columns.StokAdeti, yeniStok);
            workbook.Save();

            string islem = cozum == CakismaCozumu.UzerineEkle ? "üzerine eklendi" : "güncellendi";
            _logService.Append(filePath,
                $"GÜNCELLEME | Kod={urun.UrunKodu} | " +
                $"Stok {eskiStok.ToString(CultureInfo.InvariantCulture)} -> " +
                $"{yeniStok.ToString(CultureInfo.InvariantCulture)} ({islem})");

            return YazmaSonucu.Guncellendi;
        }

        // Çakışma yok: yeni satır ekle.
        int newRowNumber = LastUsedRowNumber(worksheet) + 1;
        IXLRow newRow = worksheet.Row(newRowNumber);

        WriteString(newRow, columns.UrunKodu, urun.UrunKodu);
        WriteDecimal(newRow, columns.StokAdeti, urun.StokAdeti);
        WriteString(newRow, columns.Birim, urun.Birim);
        WriteDecimal(newRow, columns.Fiyat, urun.Fiyat);
        WriteString(newRow, columns.ParaBirimi, urun.ParaBirimi);

        workbook.Save();

        _logService.Append(filePath,
            $"EKLEME | Kod={urun.UrunKodu} | " +
            $"Stok={urun.StokAdeti.ToString(CultureInfo.InvariantCulture)} | " +
            $"Birim={urun.Birim} | Fiyat={urun.Fiyat.ToString(CultureInfo.InvariantCulture)} | " +
            $"ParaBirimi={urun.ParaBirimi}");

        return YazmaSonucu.Eklendi;
    }

    /// <summary>
    /// <paramref name="orijinalKod"/> ile bulunan mevcut satırı, <paramref name="urun"/>
    /// içindeki değerlerle (kod dâhil) günceller. Yeni satır eklemez.
    /// </summary>
    /// <remarks>Bloklayan I/O arka plan thread'inde çalışır.</remarks>
    public Task<YazmaSonucu> UrunGuncelleAsync(
        string filePath,
        string orijinalKod,
        Urun urun,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(urun);
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("Dosya yolu boş olamaz.", nameof(filePath));
        if (string.IsNullOrWhiteSpace(orijinalKod))
            throw new ArgumentException("Orijinal ürün kodu boş olamaz.", nameof(orijinalKod));

        return Task.Run(
            () => UrunGuncelle(filePath, orijinalKod, urun, cancellationToken), cancellationToken);
    }

    private YazmaSonucu UrunGuncelle(
        string filePath, string orijinalKod, Urun urun, CancellationToken cancellationToken)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel dosyası bulunamadı.", filePath);

        cancellationToken.ThrowIfCancellationRequested();

        using var workbook = new XLWorkbook(filePath);
        IXLWorksheet worksheet = workbook.Worksheets.First();

        IXLRange? usedRange = worksheet.RangeUsed();
        if (usedRange == null)
            throw new InvalidOperationException("Excel sayfası boş; başlık satırı bulunamadı.");

        List<IXLRangeRow> rows = usedRange.RowsUsed().ToList();
        (ColumnMap columns, int headerIndex) = ResolveColumns(rows);

        // Eksik kolona yazma engeli: değeri olan bir alanın kolonu yoksa iptal et.
        EksikKolonaYazmayiEngelle(urun, columns);

        IXLRangeRow? satir = FindRowByCode(rows, headerIndex, columns.UrunKodu, orijinalKod);
        if (satir == null)
            throw new InvalidOperationException(
                $"Güncellenecek ürün bulunamadı: {orijinalKod}. Dosya değişmiş olabilir.");

        WriteString(satir, columns.UrunKodu, urun.UrunKodu);
        WriteDecimal(satir, columns.StokAdeti, urun.StokAdeti);
        WriteString(satir, columns.Birim, urun.Birim);
        WriteDecimal(satir, columns.Fiyat, urun.Fiyat);
        WriteString(satir, columns.ParaBirimi, urun.ParaBirimi);

        workbook.Save();

        string kodBilgisi = orijinalKod.Equals(urun.UrunKodu, StringComparison.OrdinalIgnoreCase)
            ? $"Kod={urun.UrunKodu}"
            : $"Kod={orijinalKod} -> {urun.UrunKodu}";

        _logService.Append(filePath,
            $"DÜZENLEME | {kodBilgisi} | " +
            $"Stok={urun.StokAdeti.ToString(CultureInfo.InvariantCulture)} | " +
            $"Birim={urun.Birim} | Fiyat={urun.Fiyat.ToString(CultureInfo.InvariantCulture)} | " +
            $"ParaBirimi={urun.ParaBirimi}");

        return YazmaSonucu.Guncellendi;
    }

    /// <summary>
    /// <paramref name="kod"/> ile bulunan satırı dosyadan siler. Silinecek satır
    /// bulunursa true, bulunamazsa false döner.
    /// </summary>
    /// <remarks>Bloklayan I/O arka plan thread'inde çalışır.</remarks>
    public Task<bool> UrunSilAsync(
        string filePath, string kod, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("Dosya yolu boş olamaz.", nameof(filePath));
        if (string.IsNullOrWhiteSpace(kod))
            throw new ArgumentException("Ürün kodu boş olamaz.", nameof(kod));

        return Task.Run(() => UrunSil(filePath, kod, cancellationToken), cancellationToken);
    }

    private bool UrunSil(string filePath, string kod, CancellationToken cancellationToken)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel dosyası bulunamadı.", filePath);

        cancellationToken.ThrowIfCancellationRequested();

        using var workbook = new XLWorkbook(filePath);
        IXLWorksheet worksheet = workbook.Worksheets.First();

        IXLRange? usedRange = worksheet.RangeUsed();
        if (usedRange == null)
            return false;

        List<IXLRangeRow> rows = usedRange.RowsUsed().ToList();
        (ColumnMap columns, int headerIndex) = ResolveColumns(rows);

        IXLRangeRow? satir = FindRowByCode(rows, headerIndex, columns.UrunKodu, kod);
        if (satir == null)
            return false;

        // Log için satır bilgisini silmeden önce oku.
        decimal stok = ReadDecimal(satir, columns.StokAdeti);
        string birim = columns.Birim == 0 ? string.Empty : satir.Cell(columns.Birim).GetString().Trim();
        decimal fiyat = ReadDecimal(satir, columns.Fiyat);
        string paraBirimi =
            columns.ParaBirimi == 0 ? string.Empty : satir.Cell(columns.ParaBirimi).GetString().Trim();

        // Tüm satırı sil; alttaki satırlar yukarı kayar.
        satir.WorksheetRow().Delete();
        workbook.Save();

        _logService.Append(filePath,
            $"SİLME | Kod={kod} | " +
            $"Stok={stok.ToString(CultureInfo.InvariantCulture)} | " +
            $"Birim={birim} | Fiyat={fiyat.ToString(CultureInfo.InvariantCulture)} | " +
            $"ParaBirimi={paraBirimi}");

        return true;
    }

    /// <summary>
    /// Okuma tarafıyla (ExcelReaderService/ColumnResolver) aynı mantıkla başlık
    /// satırını ve kolon haritasını çözer. Başlık bulunamazsa ya da zorunlu
    /// Ürün Kodu kolonu yoksa işlem tamamen engellenir.
    /// </summary>
    private static (ColumnMap Columns, int HeaderIndex) ResolveColumns(List<IXLRangeRow> rows)
    {
        if (!ColumnResolver.TryResolveHeader(rows, out int headerIndex, out ColumnMap columns))
            throw new InvalidOperationException(
                "Excel dosyasında tanınan bir başlık satırı bulunamadı.");

        // Okuma tarafındaki aynı güvenlik kontrolü: Ürün Kodu zorunlu çapa kolondur.
        if (columns.UrunKodu == 0)
            throw new InvalidOperationException("Ürün Kodu başlıklı kolon bulunamadı.");

        return (columns, headerIndex);
    }

    /// <summary>
    /// Kullanıcı bir alana değer girdiği hâlde o alanın kolonu dosyada yoksa
    /// (konum 0), kısmi/sessiz kayıt yerine işlemi tamamen iptal eder ve hangi
    /// sütun(lar)ın eksik olduğunu bildirir. Ürün Kodu ayrıca <see cref="ResolveColumns"/>
    /// içinde zorunlu tutulur.
    /// </summary>
    private static void EksikKolonaYazmayiEngelle(Urun urun, ColumnMap columns)
    {
        var eksik = new List<string>();

        if (urun.StokAdeti != 0m && columns.StokAdeti == 0) eksik.Add("Stok Adeti");
        if (!string.IsNullOrWhiteSpace(urun.Birim) && columns.Birim == 0) eksik.Add("Birim");
        if (urun.Fiyat != 0m && columns.Fiyat == 0) eksik.Add("Fiyat");
        if (!string.IsNullOrWhiteSpace(urun.ParaBirimi) && columns.ParaBirimi == 0) eksik.Add("Para Birimi");

        if (eksik.Count == 0)
            return;

        string mesaj = eksik.Count == 1
            ? $"Bu Excel dosyasında '{eksik[0]}' sütunu bulunamadığı için bu bilgi kaydedilemiyor. " +
              "Kayıt iptal edildi."
            : $"Bu Excel dosyasında şu sütunlar bulunamadığı için girdiğiniz bilgiler kaydedilemiyor: " +
              $"{string.Join(", ", eksik)}. Kayıt iptal edildi.";

        throw new InvalidOperationException(mesaj);
    }

    private static IXLRangeRow? FindRowByCode(
        List<IXLRangeRow> rows, int headerIndex, int kodColumn, string kod)
    {
        string aranan = kod.Trim();

        // Başlık ve üstündeki (logo/özet) satırları atla; yalnızca veri satırlarına bak.
        foreach (IXLRangeRow row in rows.Skip(headerIndex + 1))
        {
            string mevcut = row.Cell(kodColumn).GetString().Trim();
            if (mevcut.Equals(aranan, StringComparison.OrdinalIgnoreCase))
                return row;
        }

        return null;
    }

    private static int LastUsedRowNumber(IXLWorksheet worksheet)
    {
        IXLRow? last = worksheet.LastRowUsed();
        return last?.RowNumber() ?? 1;
    }

    private static decimal ReadDecimal(IXLRangeRow row, int column)
    {
        if (column == 0)
            return 0m;

        IXLCell cell = row.Cell(column);
        if (cell.TryGetValue(out double numeric))
            return (decimal)numeric;

        return decimal.TryParse(cell.GetString().Trim(),
            NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsed) ? parsed : 0m;
    }

    private static void WriteString(IXLRow row, int column, string value)
    {
        if (column == 0)
            return;

        row.Cell(column).Value = value;
    }

    private static void WriteString(IXLRangeRow row, int column, string value)
    {
        if (column == 0)
            return;

        row.Cell(column).Value = value;
    }

    private static void WriteDecimal(IXLRow row, int column, decimal value)
    {
        if (column == 0)
            return;

        row.Cell(column).Value = value;
    }

    private static void WriteDecimal(IXLRangeRow row, int column, decimal value)
    {
        if (column == 0)
            return;

        row.Cell(column).Value = value;
    }
}

/// <summary>
/// Excel'e yazma işleminin sonucu.
/// </summary>
public enum YazmaSonucu
{
    /// <summary>Yeni satır olarak eklendi.</summary>
    Eklendi = 0,

    /// <summary>Mevcut satır güncellendi.</summary>
    Guncellendi = 1,

    /// <summary>Aynı kod bulundu; kullanıcı kararı bekleniyor (dosyaya dokunulmadı).</summary>
    CakismaVar = 2,
}

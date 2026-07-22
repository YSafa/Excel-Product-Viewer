using System.Globalization;
using ClosedXML.Excel;
using ExcelViewer.Models;

namespace ExcelViewer.Services;

/// <summary>
/// Bir .xlsx dosyasındaki ürün satırlarını belleğe okur.
/// Dosya bir kez okunur; arama daha sonra bellek üzerinde yapılır.
/// </summary>
public sealed class ExcelReaderService
{
    // Başlık satırı, dosyanın en fazla ilk bu kadar dolu satırı içinde aranır.
    // Böylece başlıktan önce logo, rapor başlığı, tarih gibi satırlar tolere edilir.
    private const int HeaderScanRows = 10;

    /// <summary>
    /// İlk çalışma sayfasındaki tüm ürün satırlarını okur.
    /// Bloklayan I/O arka plan thread'inde çalışır, böylece UI donmaz.
    /// </summary>
    public Task<ExcelReadResult> ReadAsync(string filePath, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("Dosya yolu boş olamaz.", nameof(filePath));

        return Task.Run(() => Read(filePath, cancellationToken), cancellationToken);
    }

    private static ExcelReadResult Read(string filePath, CancellationToken cancellationToken)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel dosyası bulunamadı.", filePath);

        // FileShare.ReadWrite, dosya Excel'de açık olsa bile okumamıza izin verir.
        using var stream = new FileStream(
            filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var workbook = new XLWorkbook(stream);

        IXLWorksheet worksheet = workbook.Worksheets.First();
        IXLRange? usedRange = worksheet.RangeUsed();
        if (usedRange == null)
            throw new InvalidDataException("Excel sayfası boş; okunacak veri bulunamadı.");

        // Başlık satırını tespit et: ilk 10 dolu satır içinde, EN ÇOK mantıksal
        // kolona eşlenen satır başlık kabul edilir. Böylece başlıktan önce logo,
        // rapor başlığı, tarih veya "Model/Total" gibi tek bir alias'a çakışan
        // özet satırları olsa da doğru (en zengin) başlık satırı seçilir.
        List<IXLRangeRow> rows = usedRange.RowsUsed().ToList();
        int headerIndex = -1;
        int enIyiSkor = 0;
        int enIyiGenislik = 0;
        ColumnMap columns = default;
        int scanLimit = Math.Min(HeaderScanRows, rows.Count);
        for (int i = 0; i < scanLimit; i++)
        {
            // Altındaki veri satırlarını da geç: aynı hücre hem Birim hem Fiyat
            // anahtar kelimesini içerdiğinde (örn. "Unit Price") kararı verinin
            // sayısal/metin olmasına göre vermek için kullanılır.
            ColumnMap aday = ColumnResolver.Resolve(rows[i], rows.Skip(i + 1));
            int skor = ColumnResolver.EslesenKolonSayisi(aday);
            if (skor == 0)
                continue;

            // Önce en çok kolona eşlenen; eşitlikte daha geniş (daha çok dolu
            // hücreli) satır — gerçek başlık, üstteki dar özet satırını yener.
            int genislik = ColumnResolver.DoluBaslikSayisi(rows[i]);
            if (skor > enIyiSkor || (skor == enIyiSkor && genislik > enIyiGenislik))
            {
                enIyiSkor = skor;
                enIyiGenislik = genislik;
                headerIndex = i;
                columns = aday;
            }
        }

        if (headerIndex < 0)
            throw new InvalidDataException(
                "Excel dosyasında tanınan bir başlık satırı bulunamadı. " +
                "En az 'Ürün Kodu' sütununu içeren bir başlık satırı gerekli.");

        IXLRangeRow headerRow = rows[headerIndex];

        // Ürün Kodu, arama/çakışma kontrolü gibi işlevlerin dayandığı çapa alandır;
        // bulunamazsa dosya okunamadı sayılır (sessizce boş veri gösterilmez).
        if (columns.UrunKodu == 0)
            throw new InvalidDataException(
                "'Ürün Kodu' sütunu bulunamadı. Bu sütun zorunlu olduğundan dosya okunamadı.");

        IReadOnlyList<string> unmatchedHeaders = ColumnResolver.UnmatchedHeaders(headerRow, columns);

        var products = new List<Urun>();
        foreach (IXLRangeRow row in rows.Skip(headerIndex + 1))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var urun = new Urun
            {
                UrunKodu = GetString(row, columns.UrunKodu),
                StokAdeti = GetDecimal(row, columns.StokAdeti),
                Birim = GetString(row, columns.Birim),
                Fiyat = GetDecimal(row, columns.Fiyat),
                ParaBirimi = GetString(row, columns.ParaBirimi),
            };

            // Tamamen boş satırları atla (Ürün Kodu ve birim yoksa).
            if (urun.UrunKodu.Length == 0 && urun.Birim.Length == 0)
                continue;

            products.Add(urun);
        }

        return new ExcelReadResult(products, columns, unmatchedHeaders);
    }

    private static string GetString(IXLRangeRow row, int column)
    {
        if (column == 0)
            return string.Empty;

        return row.Cell(column).GetString().Trim();
    }

    private static decimal GetDecimal(IXLRangeRow row, int column)
    {
        if (column == 0)
            return 0m;

        IXLCell cell = row.Cell(column);
        if (cell.TryGetValue(out double numeric))
            return (decimal)numeric;

        string raw = cell.GetString().Trim();
        return decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsed)
            ? parsed
            : 0m;
    }
}

/// <summary>
/// Bir Excel dosyasının okuma sonucu: okunan ürünler, hangi mantıksal kolonların
/// gerçekten bulunduğu (<see cref="ColumnMap"/>; 0 = bulunamadı) ve bizim
/// şemamıza uymayan (eşleşmeyen) Excel başlıkları.
/// </summary>
public sealed record ExcelReadResult(
    List<Urun> Products,
    ColumnMap Columns,
    IReadOnlyList<string> UnmatchedHeaders);

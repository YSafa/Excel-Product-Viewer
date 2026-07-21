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
    /// <summary>
    /// İlk çalışma sayfasındaki tüm ürün satırlarını okur.
    /// Bloklayan I/O arka plan thread'inde çalışır, böylece UI donmaz.
    /// </summary>
    public Task<List<Urun>> ReadAsync(string filePath, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("Dosya yolu boş olamaz.", nameof(filePath));

        return Task.Run(() => Read(filePath, cancellationToken), cancellationToken);
    }

    private static List<Urun> Read(string filePath, CancellationToken cancellationToken)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel dosyası bulunamadı.", filePath);

        var result = new List<Urun>();

        // FileShare.ReadWrite, dosya Excel'de açık olsa bile okumamıza izin verir.
        using var stream = new FileStream(
            filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var workbook = new XLWorkbook(stream);

        IXLWorksheet worksheet = workbook.Worksheets.First();
        IXLRange? usedRange = worksheet.RangeUsed();
        if (usedRange == null)
            return result; // Boş sayfa.

        // Kolon konumlarını başlık satırından çöz (Türkçe + İngilizce alternatifler).
        IXLRangeRow headerRow = usedRange.FirstRow();
        ColumnMap columns = ColumnResolver.Resolve(headerRow);

        foreach (IXLRangeRow row in usedRange.RowsUsed().Skip(1))
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

            // Tamamen boş satırları atla (kod ve birim yoksa).
            if (urun.UrunKodu.Length == 0 && urun.Birim.Length == 0)
                continue;

            result.Add(urun);
        }

        return result;
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

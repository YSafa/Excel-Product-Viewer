using ClosedXML.Excel;
using ExcelViewer.Models;

namespace ExcelViewer.Services;

/// <summary>
/// Bellekteki ürün listesini (grid'de o an görünen, filtre/arama/sıralama
/// uygulanmış hâliyle) YENİ bir .xlsx dosyasına yazar. Orijinal açık dosyaya
/// dokunmaz; her zaman sıfırdan bir workbook oluşturur ve standart Türkçe
/// başlıkları kullanır.
/// </summary>
public sealed class ExcelExportService
{
    private static readonly string[] Basliklar =
        { "Ürün Kodu", "Stok Adeti", "Birim", "Fiyat", "Para Birimi" };

    /// <summary>
    /// Ürünleri <paramref name="filePath"/> konumuna yeni bir Excel dosyası
    /// olarak yazar. Bloklayan I/O arka plan thread'inde çalışır, böylece UI donmaz.
    /// </summary>
    public Task ExportAsync(
        string filePath, IReadOnlyList<Urun> urunler, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(urunler);
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("Dosya yolu boş olamaz.", nameof(filePath));

        return Task.Run(() => Export(filePath, urunler, cancellationToken), cancellationToken);
    }

    private static void Export(string filePath, IReadOnlyList<Urun> urunler, CancellationToken cancellationToken)
    {
        using var workbook = new XLWorkbook();
        IXLWorksheet worksheet = workbook.AddWorksheet("Ürünler");

        for (int c = 0; c < Basliklar.Length; c++)
            worksheet.Cell(1, c + 1).Value = Basliklar[c];

        int row = 2;
        foreach (Urun urun in urunler)
        {
            cancellationToken.ThrowIfCancellationRequested();

            worksheet.Cell(row, 1).Value = urun.UrunKodu;
            worksheet.Cell(row, 2).Value = urun.StokAdeti;
            worksheet.Cell(row, 3).Value = urun.Birim;
            worksheet.Cell(row, 4).Value = urun.Fiyat;
            worksheet.Cell(row, 5).Value = urun.ParaBirimi;
            row++;
        }

        worksheet.Row(1).Style.Font.Bold = true;
        worksheet.SheetView.FreezeRows(1);
        worksheet.Columns().AdjustToContents();

        workbook.SaveAs(filePath);
    }
}

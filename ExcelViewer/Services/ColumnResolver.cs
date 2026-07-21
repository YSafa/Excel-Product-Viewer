using ClosedXML.Excel;

namespace ExcelViewer.Services;

/// <summary>
/// Excel başlık satırındaki kolonları, Türkçe ve İngilizce alternatif adları
/// dikkate alarak çözer. Hem okuma hem yazma servisleri bu sınıfı kullanır,
/// böylece başlık eşleştirme mantığı tek yerde tutulur.
/// </summary>
public static class ColumnResolver
{
    // Her mantıksal kolon için kabul edilen başlık adları (büyük/küçük harf duyarsız).
    private static readonly string[] UrunKoduAliases =
        { "Ürün Kodu", "UrunKodu", "Ürün Kod", "Kod", "Product Code", "ProductCode", "Code", "Stock Code", "StockCode" };

    private static readonly string[] StokAdetiAliases =
        { "Stok Adeti", "Stok Adedi", "StokAdeti", "Stok", "Adet", "Miktar", "Stock", "Quantity", "Qty", "Amount", "Count" };

    private static readonly string[] BirimAliases =
        { "Birim", "Ölçü", "Ölçü Birimi", "Unit", "Measure", "UOM" };

    private static readonly string[] FiyatAliases =
        { "Fiyat", "Ücret", "Tutar", "Price", "Cost", "Amount Price" };

    private static readonly string[] ParaBirimiAliases =
        { "Para Birimi", "ParaBirimi", "Döviz", "Kur", "Currency", "Curr" };

    /// <summary>
    /// Başlık satırından her mantıksal kolonun konumunu döndürür.
    /// Bulunamayan kolon 0 değerini alır.
    /// </summary>
    public static ColumnMap Resolve(IXLRangeRow headerRow)
    {
        int urunKodu = 0, stokAdeti = 0, birim = 0, fiyat = 0, paraBirimi = 0;

        foreach (IXLCell cell in headerRow.Cells())
        {
            string header = cell.GetString().Trim();
            if (header.Length == 0)
                continue;

            int index = cell.Address.ColumnNumber;

            if (Matches(header, UrunKoduAliases)) urunKodu = index;
            else if (Matches(header, StokAdetiAliases)) stokAdeti = index;
            else if (Matches(header, BirimAliases)) birim = index;
            else if (Matches(header, FiyatAliases)) fiyat = index;
            else if (Matches(header, ParaBirimiAliases)) paraBirimi = index;
        }

        return new ColumnMap(urunKodu, stokAdeti, birim, fiyat, paraBirimi);
    }

    private static bool Matches(string header, string[] aliases)
    {
        foreach (string alias in aliases)
        {
            if (header.Equals(alias, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }
}

/// <summary>
/// Excel'deki her mantıksal kolonun 1 tabanlı konumu. 0 = kolon bulunamadı.
/// </summary>
public readonly record struct ColumnMap(
    int UrunKodu, int StokAdeti, int Birim, int Fiyat, int ParaBirimi);

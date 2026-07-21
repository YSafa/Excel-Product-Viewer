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
        { "Ürün Kodu", "UrunKodu", "Ürün Kod", "Kod", "Product Code", "ProductCode", "Code", "Stock Code", "StockCode", "Model" };

    private static readonly string[] StokAdetiAliases =
        { "Stok Adeti", "Stok Adedi", "StokAdeti", "Stok", "Adet", "Miktar", "Stock", "Quantity", "Qty", "Amount", "Count", "Total" };

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

    /// <summary>
    /// Verilen başlık metni, bilinen kolon adlarından herhangi birine birebir
    /// eşleşiyorsa true döner (büyük/küçük harf duyarsız).
    /// </summary>
    public static bool IsKnownHeader(string header)
    {
        return Matches(header, UrunKoduAliases)
            || Matches(header, StokAdetiAliases)
            || Matches(header, BirimAliases)
            || Matches(header, FiyatAliases)
            || Matches(header, ParaBirimiAliases);
    }

    /// <summary>
    /// Satır, bilinen bir kolon adına birebir eşleşen en az bir hücre içeriyor mu?
    /// Başlık satırını tespit etmek için kullanılır (logo/tarih gibi satırları eler).
    /// </summary>
    public static bool RowHasKnownHeader(IXLRangeRow row)
    {
        foreach (IXLCell cell in row.Cells())
        {
            string header = cell.GetString().Trim();
            if (header.Length > 0 && IsKnownHeader(header))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Başlık satırındaki, bizim 5 kolonumuzdan hiçbirine eşleşmeyen sütun
    /// başlıkları (örn. RAF, TOTAL). Kullanıcı uyarısında gösterilir.
    /// </summary>
    public static List<string> UnmatchedHeaders(IXLRangeRow headerRow)
    {
        var unmatched = new List<string>();

        foreach (IXLCell cell in headerRow.Cells())
        {
            string header = cell.GetString().Trim();
            if (header.Length > 0 && !IsKnownHeader(header))
                unmatched.Add(header);
        }

        return unmatched;
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

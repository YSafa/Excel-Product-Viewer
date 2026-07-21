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

            // İlk (soldaki) eşleşme kazanır: bir mantıksal kolon zaten
            // bulunduysa, sonradan gelen ve geniş alias'lara çakışan bir başlık
            // (örn. gerçek "Stok Adeti" varken "Total") onu ezmez; o başlık
            // eşleşmeyen sütun olarak kalır. Her kolonun bir kez geçtiği normal
            // dosyalarda davranış değişmez.
            if (urunKodu == 0 && Matches(header, UrunKoduAliases)) urunKodu = index;
            else if (stokAdeti == 0 && Matches(header, StokAdetiAliases)) stokAdeti = index;
            else if (birim == 0 && Matches(header, BirimAliases)) birim = index;
            else if (fiyat == 0 && Matches(header, FiyatAliases)) fiyat = index;
            else if (paraBirimi == 0 && Matches(header, ParaBirimiAliases)) paraBirimi = index;
        }

        return new ColumnMap(urunKodu, stokAdeti, birim, fiyat, paraBirimi);
    }

    /// <summary>
    /// Bir satırın kaç mantıksal kolona eşlendiğini döndürür (0-5). Başlık
    /// satırını tespit ederken kullanılır: en çok kolona eşlenen satır başlıktır.
    /// Böylece başlıktan önceki özet/başlık satırlarında tek bir alias'a uyan
    /// hücre (örn. "Model", "Total") bulunsa bile gerçek başlık satırı seçilir.
    /// </summary>
    public static int EslesenKolonSayisi(ColumnMap map)
    {
        int sayi = 0;
        if (map.UrunKodu != 0) sayi++;
        if (map.StokAdeti != 0) sayi++;
        if (map.Birim != 0) sayi++;
        if (map.Fiyat != 0) sayi++;
        if (map.ParaBirimi != 0) sayi++;
        return sayi;
    }

    /// <summary>
    /// Satırdaki dolu (boş olmayan) hücre sayısı. Başlık tespitinde eşitlik
    /// bozucu olarak kullanılır: aynı sayıda kolona eşlenen iki satırdan daha
    /// geniş olanı (gerçek başlık) tercih edilir; böylece üstteki dar özet
    /// satırları elenir.
    /// </summary>
    public static int DoluBaslikSayisi(IXLRangeRow row)
    {
        int sayi = 0;
        foreach (IXLCell cell in row.Cells())
        {
            if (cell.GetString().Trim().Length > 0)
                sayi++;
        }

        return sayi;
    }

    /// <summary>
    /// Başlık satırındaki, bizim 5 kolonumuzdan hiçbirine ATANMAMIŞ sütun
    /// başlıkları (örn. RAF, TOTAL, haftalık sayı sütunları). Kullanıcı
    /// uyarısında gösterilir.
    ///
    /// Not: Alias'a uyup uymadığına değil, gerçekte hangi kolon konumlarının
    /// <paramref name="columns"/> içinde kullanıldığına bakılır. Böylece geniş
    /// alias'lara (örn. "Total") çakışan ama şemamıza atanmayan bir başlık da
    /// doğru şekilde "eşleşmeyen" olarak listelenir.
    /// </summary>
    public static List<string> UnmatchedHeaders(IXLRangeRow headerRow, ColumnMap columns)
    {
        var atananKonumlar = new HashSet<int>
        {
            columns.UrunKodu, columns.StokAdeti, columns.Birim, columns.Fiyat, columns.ParaBirimi,
        };

        var unmatched = new List<string>();
        foreach (IXLCell cell in headerRow.Cells())
        {
            string header = cell.GetString().Trim();
            if (header.Length == 0)
                continue;

            if (!atananKonumlar.Contains(cell.Address.ColumnNumber))
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

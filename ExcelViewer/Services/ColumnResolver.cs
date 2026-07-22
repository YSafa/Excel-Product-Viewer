using System.Globalization;
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

    // Tam eşleşme (yukarıdaki alias'lar) hiçbir kolonda bulunamadığında kullanılan
    // "içerir" (substring) anahtar kelimeleri. Başlık, bu kelimelerden birini
    // (büyük/küçük harf duyarsız) içeriyorsa ilgili kolon kabul edilir. Örn.
    // "2026 Unit Price (euro)" başlığı "price" içerdiği için Fiyat'a aday olur.
    // Anahtar kelimeler bilinçli olarak alias'lardan daha dar ve ayırt edicidir;
    // böylece rastgele başlıklara yanlış eşleşme riski azalır.
    private static readonly string[] UrunKoduKeywords =
        { "ürün kod", "urun kod", "product code", "item code", "stock code", "kod", "code", "model" };

    private static readonly string[] StokAdetiKeywords =
        { "stok", "stock", "adet", "miktar", "quantity", "qty" };

    private static readonly string[] BirimKeywords =
        { "birim", "ölçü", "olcu", "unit", "measure", "uom" };

    private static readonly string[] FiyatKeywords =
        { "fiyat", "ücret", "ucret", "tutar", "price", "cost" };

    private static readonly string[] ParaBirimiKeywords =
        { "para birim", "döviz", "doviz", "currency", "kur" };

    /// <summary>
    /// Başlık satırındaki mantıksal kolonların hangi mantıksal alana ait olduğunu
    /// belirtir. Fallback eşleştirmede iç kullanım içindir.
    /// </summary>
    private enum LogicalColumn { UrunKodu, StokAdeti, Birim, Fiyat, ParaBirimi }

    /// <summary>
    /// Başlık satırından her mantıksal kolonun konumunu döndürür.
    /// Bulunamayan kolon 0 değerini alır.
    ///
    /// İki geçişli çalışır: önce tüm kolonlar için TAM eşleşme (alias) denenir;
    /// tam eşleşmeyle bulunamayan kolonlar için ikinci geçişte "içerir" (substring)
    /// fallback'i uygulanır. Bir kolon tam eşleşmeyle bulunduysa fallback o kolon
    /// için hiç çalışmaz, dolayısıyla eski (tam eşleşen) dosyalarda davranış aynıdır.
    /// </summary>
    /// <param name="headerRow">Başlık satırı.</param>
    /// <param name="dataRows">
    /// Başlığın altındaki veri satırları (opsiyonel). Yalnızca aynı başlık hücresi
    /// hem Birim hem Fiyat anahtar kelimesini içerdiğinde (örn. "Unit Price"),
    /// altındaki verinin sayısal mı metin mi olduğuna bakarak karar vermek için
    /// kullanılır. null ise böyle bir çakışmada Fiyat varsayılır.
    /// </param>
    public static ColumnMap Resolve(IXLRangeRow headerRow, IEnumerable<IXLRangeRow>? dataRows = null)
    {
        int urunKodu = 0, stokAdeti = 0, birim = 0, fiyat = 0, paraBirimi = 0;

        // Başlık hücrelerini (boş olmayanları) konumlarıyla topla.
        var cells = new List<(int Index, string Text)>();
        foreach (IXLCell cell in headerRow.Cells())
        {
            string header = cell.GetString().Trim();
            if (header.Length > 0)
                cells.Add((cell.Address.ColumnNumber, header));
        }

        // Atanmış başlık konumları: aynı hücre iki kez kullanılmasın.
        var assigned = new HashSet<int>();

        // 1. GEÇİŞ — TAM EŞLEŞME. İlk (soldaki) eşleşme kazanır: bir mantıksal
        // kolon zaten bulunduysa, sonradan gelen ve geniş alias'lara çakışan bir
        // başlık (örn. gerçek "Stok Adeti" varken "Total") onu ezmez; o başlık
        // eşleşmeyen sütun olarak kalır. Her kolonun bir kez geçtiği normal
        // dosyalarda davranış değişmez.
        foreach ((int index, string header) in cells)
        {
            if (urunKodu == 0 && Matches(header, UrunKoduAliases)) { urunKodu = index; assigned.Add(index); }
            else if (stokAdeti == 0 && Matches(header, StokAdetiAliases)) { stokAdeti = index; assigned.Add(index); }
            else if (birim == 0 && Matches(header, BirimAliases)) { birim = index; assigned.Add(index); }
            else if (fiyat == 0 && Matches(header, FiyatAliases)) { fiyat = index; assigned.Add(index); }
            else if (paraBirimi == 0 && Matches(header, ParaBirimiAliases)) { paraBirimi = index; assigned.Add(index); }
        }

        // 2. GEÇİŞ — "İÇERİR" FALLBACK. Yalnızca tam eşleşmeyle BULUNAMAYAN
        // kolonlar için çalışır. Hücreler soldan sağa taranır; bir kolona ilk
        // uyan (henüz atanmamış) hücre kazanır — mevcut "ilk eşleşme kazanır"
        // mantığıyla tutarlı. Bir hücre birden fazla kolonun anahtar kelimesini
        // içeriyorsa çakışma çözümü uygulanır.
        foreach ((int index, string header) in cells)
        {
            if (assigned.Contains(index))
                continue;

            var matched = new List<(LogicalColumn Col, int KeywordLength)>();
            if (urunKodu == 0) AddIfMatched(matched, LogicalColumn.UrunKodu, header, UrunKoduKeywords);
            if (stokAdeti == 0) AddIfMatched(matched, LogicalColumn.StokAdeti, header, StokAdetiKeywords);
            if (birim == 0) AddIfMatched(matched, LogicalColumn.Birim, header, BirimKeywords);
            if (fiyat == 0) AddIfMatched(matched, LogicalColumn.Fiyat, header, FiyatKeywords);
            if (paraBirimi == 0) AddIfMatched(matched, LogicalColumn.ParaBirimi, header, ParaBirimiKeywords);

            if (matched.Count == 0)
                continue;

            LogicalColumn winner = ResolveConflict(matched, dataRows, index);

            switch (winner)
            {
                case LogicalColumn.UrunKodu: urunKodu = index; break;
                case LogicalColumn.StokAdeti: stokAdeti = index; break;
                case LogicalColumn.Birim: birim = index; break;
                case LogicalColumn.Fiyat: fiyat = index; break;
                case LogicalColumn.ParaBirimi: paraBirimi = index; break;
            }

            assigned.Add(index);
        }

        return new ColumnMap(urunKodu, stokAdeti, birim, fiyat, paraBirimi);
    }

    /// <summary>
    /// Aynı başlık hücresine birden fazla kolonun anahtar kelimesi uyduğunda
    /// kazananı seçer.
    ///
    /// Özel durum — Birim ↔ Fiyat: "Unit Price" gibi bir başlık hem "unit" hem
    /// "price" içerir. Bu durumda kolonun altındaki veri hücrelerine bakılır:
    /// veri sayısalsa Fiyat, metinse Birim atanır. Karar verilemezse (veri yok
    /// ya da karışık) Fiyat tercih edilir — sayısal kolonun yanlış atanması
    /// metinsel kolondan daha görünür bir hatadır.
    ///
    /// Diğer tüm çoklu eşleşmelerde daha uzun (daha ayırt edici) anahtar kelimeyi
    /// içeren kolon kazanır; eşitlikte kolon sırası (enum) belirleyicidir.
    /// </summary>
    private static LogicalColumn ResolveConflict(
        List<(LogicalColumn Col, int KeywordLength)> matched,
        IEnumerable<IXLRangeRow>? dataRows,
        int columnNumber)
    {
        if (matched.Count == 1)
            return matched[0].Col;

        bool birimVar = matched.Exists(m => m.Col == LogicalColumn.Birim);
        bool fiyatVar = matched.Exists(m => m.Col == LogicalColumn.Fiyat);
        if (matched.Count == 2 && birimVar && fiyatVar)
        {
            return SutunVerisiMetinsel(dataRows, columnNumber)
                ? LogicalColumn.Birim
                : LogicalColumn.Fiyat;
        }

        LogicalColumn winner = matched[0].Col;
        int bestLength = matched[0].KeywordLength;
        foreach ((LogicalColumn col, int length) in matched)
        {
            if (length > bestLength || (length == bestLength && (int)col < (int)winner))
            {
                winner = col;
                bestLength = length;
            }
        }

        return winner;
    }

    /// <summary>
    /// Verilen kolonun altındaki ilk birkaç dolu veri hücresine bakarak verinin
    /// baskın olarak metinsel mi olduğunu döndürür. Sayısal (decimal olarak parse
    /// edilebilen) değerler metin sayılmaz. Örnekleme küçük tutulur; tüm dosyayı
    /// taramak gerekmez.
    /// </summary>
    private static bool SutunVerisiMetinsel(IEnumerable<IXLRangeRow>? dataRows, int columnNumber)
    {
        if (dataRows == null)
            return false; // Veri yoksa Fiyat varsayılır (metinsel değil).

        const int MaxOrnek = 5;
        int metin = 0, sayisal = 0, incelenen = 0;

        foreach (IXLRangeRow row in dataRows)
        {
            IXLCell cell = row.Cell(columnNumber);
            string raw = cell.GetString().Trim();
            if (raw.Length == 0)
                continue;

            if (cell.TryGetValue(out double _) ||
                decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            {
                sayisal++;
            }
            else
            {
                metin++;
            }

            if (++incelenen >= MaxOrnek)
                break;
        }

        // Metin baskınsa Birim; sayısal, eşit ya da boşsa Fiyat (metinsel değil).
        return metin > sayisal;
    }

    /// <summary>
    /// Başlık, verilen anahtar kelimelerden birini içeriyorsa listeye en uzun
    /// eşleşen kelimenin uzunluğuyla ekler. Uzunluk, çoklu eşleşme çözümünde
    /// "en ayırt edici kelime kazanır" için kullanılır.
    /// </summary>
    private static void AddIfMatched(
        List<(LogicalColumn Col, int KeywordLength)> matched,
        LogicalColumn col, string header, string[] keywords)
    {
        int bestLength = 0;
        foreach (string keyword in keywords)
        {
            if (keyword.Length > bestLength &&
                header.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                bestLength = keyword.Length;
            }
        }

        if (bestLength > 0)
            matched.Add((col, bestLength));
    }

    // Başlık satırı, dosyanın en fazla ilk bu kadar dolu satırı içinde aranır.
    // Böylece başlıktan önce logo, rapor başlığı, tarih gibi satırlar tolere edilir.
    // Hem okuma hem yazma tarafı aynı sınırı kullansın diye burada tutulur.
    public const int HeaderScanRows = 10;

    /// <summary>
    /// Verilen satırlar içinde başlık satırını tespit eder ve kolon haritasını
    /// döndürür. Okuma ve yazma tarafı bu ortak mantığı paylaşır, böylece ikisi
    /// de aynı satırı başlık kabul edip aynı kolonlara yazar/okur.
    ///
    /// İlk <see cref="HeaderScanRows"/> satır içinde EN ÇOK mantıksal kolona
    /// eşlenen satır başlık kabul edilir; eşitlikte daha geniş (daha çok dolu
    /// hücreli) satır — gerçek başlık, üstteki dar özet satırını yener. Her aday
    /// satır için tam eşleşme + substring fallback + tek hücrede çoklu anahtar
    /// kelime çakışması (<see cref="Resolve"/>) uygulanır.
    /// </summary>
    /// <returns>Başlık bulunduysa true; hiçbir satır bir kolona eşlenmezse false.</returns>
    public static bool TryResolveHeader(
        IReadOnlyList<IXLRangeRow> rows, out int headerIndex, out ColumnMap columns)
    {
        headerIndex = -1;
        columns = default;
        int enIyiSkor = 0;
        int enIyiGenislik = 0;
        int scanLimit = Math.Min(HeaderScanRows, rows.Count);

        for (int i = 0; i < scanLimit; i++)
        {
            // Altındaki satırları geç: aynı hücre hem Birim hem Fiyat anahtar
            // kelimesini içerdiğinde (örn. "Unit Price") kararı verinin
            // sayısal/metin olmasına göre vermek için kullanılır.
            ColumnMap aday = Resolve(rows[i], rows.Skip(i + 1));
            int skor = EslesenKolonSayisi(aday);
            if (skor == 0)
                continue;

            int genislik = DoluBaslikSayisi(rows[i]);
            if (skor > enIyiSkor || (skor == enIyiSkor && genislik > enIyiGenislik))
            {
                enIyiSkor = skor;
                enIyiGenislik = genislik;
                headerIndex = i;
                columns = aday;
            }
        }

        return headerIndex >= 0;
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

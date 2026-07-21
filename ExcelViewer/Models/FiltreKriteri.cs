namespace ExcelViewer.Models;

/// <summary>
/// Ürün listesine uygulanacak filtre kriterleri.
/// Null aralık sınırları "sınır yok" anlamına gelir; boş seçim kümesi
/// "hepsi görünsün" anlamına gelir.
/// </summary>
public sealed class FiltreKriteri
{
    public decimal? StokMin { get; set; }
    public decimal? StokMax { get; set; }

    public decimal? FiyatMin { get; set; }
    public decimal? FiyatMax { get; set; }

    /// <summary>Görünecek birimler. Boşsa tüm birimler görünür.</summary>
    public HashSet<string> SeciliBirimler { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Görünecek para birimleri. Boşsa tümü görünür.</summary>
    public HashSet<string> SeciliParaBirimleri { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Hiçbir kriter uygulanmıyorsa true.</summary>
    public bool BosMu =>
        StokMin is null && StokMax is null &&
        FiyatMin is null && FiyatMax is null &&
        SeciliBirimler.Count == 0 && SeciliParaBirimleri.Count == 0;
}

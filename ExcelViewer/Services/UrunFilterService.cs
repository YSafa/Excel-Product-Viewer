using ExcelViewer.Models;

namespace ExcelViewer.Services;

/// <summary>
/// Bellek üzerindeki ürün listesine metin araması, filtre kriterleri ve
/// sıralama uygular. Sıra: önce metin araması, sonra filtre, en son sıralama.
/// </summary>
public static class UrunFilterService
{
    /// <summary>
    /// Ürünleri arama terimi + filtre + sıralamaya göre işler.
    /// Herhangi biri null/boş olabilir; o adım atlanır.
    /// </summary>
    public static List<Urun> Uygula(
        IReadOnlyList<Urun> source,
        string? searchTerm,
        FiltreKriteri? filtre,
        SiralamaKriteri? siralama)
    {
        ArgumentNullException.ThrowIfNull(source);

        IEnumerable<Urun> query = source;

        // 1) Metin araması (Ürün Kodu veya Birim).
        string term = (searchTerm ?? string.Empty).Trim();
        if (term.Length > 0)
        {
            string upperTerm = term.ToUpperInvariant();
            query = query.Where(u =>
                Contains(u.UrunKodu, upperTerm) || Contains(u.Birim, upperTerm));
        }

        // 2) Filtre kriterleri.
        if (filtre is { BosMu: false })
        {
            query = query.Where(u => FiltreyeUyuyor(u, filtre));
        }

        // 3) Sıralama.
        if (siralama is { Alan: not SiralamaAlani.Yok })
        {
            query = Sirala(query, siralama);
        }

        return query.ToList();
    }

    private static bool FiltreyeUyuyor(Urun urun, FiltreKriteri f)
    {
        if (f.StokMin is { } smin && urun.StokAdeti < smin)
            return false;
        if (f.StokMax is { } smax && urun.StokAdeti > smax)
            return false;

        if (f.FiyatMin is { } fmin && urun.Fiyat < fmin)
            return false;
        if (f.FiyatMax is { } fmax && urun.Fiyat > fmax)
            return false;

        if (f.SeciliBirimler.Count > 0 && !f.SeciliBirimler.Contains(urun.Birim))
            return false;

        if (f.SeciliParaBirimleri.Count > 0 && !f.SeciliParaBirimleri.Contains(urun.ParaBirimi))
            return false;

        return true;
    }

    private static IEnumerable<Urun> Sirala(IEnumerable<Urun> query, SiralamaKriteri s)
    {
        bool artan = s.Yon == SiralamaYonu.Artan;

        return s.Alan switch
        {
            SiralamaAlani.UrunKodu => artan
                ? query.OrderBy(u => u.UrunKodu, StringComparer.OrdinalIgnoreCase)
                : query.OrderByDescending(u => u.UrunKodu, StringComparer.OrdinalIgnoreCase),

            SiralamaAlani.StokAdeti => artan
                ? query.OrderBy(u => u.StokAdeti)
                : query.OrderByDescending(u => u.StokAdeti),

            SiralamaAlani.Birim => artan
                ? query.OrderBy(u => u.Birim, StringComparer.OrdinalIgnoreCase)
                : query.OrderByDescending(u => u.Birim, StringComparer.OrdinalIgnoreCase),

            SiralamaAlani.Fiyat => artan
                ? query.OrderBy(u => u.Fiyat)
                : query.OrderByDescending(u => u.Fiyat),

            SiralamaAlani.ParaBirimi => artan
                ? query.OrderBy(u => u.ParaBirimi, StringComparer.OrdinalIgnoreCase)
                : query.OrderByDescending(u => u.ParaBirimi, StringComparer.OrdinalIgnoreCase),

            _ => query,
        };
    }

    private static bool Contains(string? value, string upperTerm)
    {
        if (string.IsNullOrEmpty(value))
            return false;

        return value.ToUpperInvariant().Contains(upperTerm, StringComparison.Ordinal);
    }
}

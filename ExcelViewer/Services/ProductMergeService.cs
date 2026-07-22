using ExcelViewer.Models;

namespace ExcelViewer.Services;

/// <summary>
/// İki ürün listesini (açık dosya A + seçilen dosya B) bellekte birleştirir.
/// Karar kolon bazındadır (satır bazında değil): çakışan her kolon için tek bir
/// <see cref="KolonKarari"/> tüm satırlara aynı şekilde uygulanır.
///
/// Kurallar:
/// <list type="bullet">
/// <item>A'daki bir ürünün kodu B'de eşleşiyorsa, A'da hiç bulunmayan kolonlar
/// B'den doldurulur.</item>
/// <item>Kolon hem A hem B'de varsa (çakışma) kullanıcı kararı uygulanır:
/// A değeri / B değeri / (yalnızca Stok Adeti) toplam.</item>
/// <item>B'de olup A'da bulunmayan kodlar yeni satır olarak sona eklenir.</item>
/// </list>
/// Eşleştirme anahtarı Ürün Kodu: trim + <see cref="StringComparer.OrdinalIgnoreCase"/>.
/// </summary>
public static class ProductMergeService
{
    public static BirlestirmeSonucu Birlestir(
        IReadOnlyList<Urun> a, ColumnMap aCols,
        IReadOnlyList<Urun> b, ColumnMap bCols,
        BirlestirmeKararlari kararlar)
    {
        ArgumentNullException.ThrowIfNull(a);
        ArgumentNullException.ThrowIfNull(b);
        ArgumentNullException.ThrowIfNull(kararlar);

        // B'yi koda göre indeksle; ilk görülen satır kazanır (çakışma kontrolüyle tutarlı).
        var bMap = new Dictionary<string, Urun>(StringComparer.OrdinalIgnoreCase);
        foreach (Urun u in b)
        {
            string key = u.UrunKodu.Trim();
            if (key.Length > 0 && !bMap.ContainsKey(key))
                bMap[key] = u;
        }

        var sonuc = new List<Urun>(a.Count);
        var eslesenBKodlari = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int dolduruldu = 0;

        foreach (Urun au in a)
        {
            string key = au.UrunKodu.Trim();
            Urun? bu = null;
            if (key.Length > 0 && bMap.TryGetValue(key, out Urun? bulunan))
            {
                bu = bulunan;
                eslesenBKodlari.Add(key);
            }

            bool bilgiDolduruldu = false;
            var merged = new Urun
            {
                UrunKodu = au.UrunKodu,
                StokAdeti = SayiBirlestir(aCols.StokAdeti != 0, au.StokAdeti,
                    bu, bCols.StokAdeti != 0, bu?.StokAdeti ?? 0m, kararlar.StokAdeti, ref bilgiDolduruldu),
                Birim = MetinBirlestir(aCols.Birim != 0, au.Birim,
                    bu, bCols.Birim != 0, bu?.Birim ?? string.Empty, kararlar.Birim, ref bilgiDolduruldu),
                Fiyat = SayiBirlestir(aCols.Fiyat != 0, au.Fiyat,
                    bu, bCols.Fiyat != 0, bu?.Fiyat ?? 0m, kararlar.Fiyat, ref bilgiDolduruldu),
                ParaBirimi = MetinBirlestir(aCols.ParaBirimi != 0, au.ParaBirimi,
                    bu, bCols.ParaBirimi != 0, bu?.ParaBirimi ?? string.Empty, kararlar.ParaBirimi, ref bilgiDolduruldu),
            };

            if (bilgiDolduruldu)
                dolduruldu++;
            sonuc.Add(merged);
        }

        // B'de olup A'da hiç bulunmayan kodlar: yeni satır olarak ekle.
        int eklendi = 0;
        foreach (Urun bu in b)
        {
            string key = bu.UrunKodu.Trim();
            if (key.Length == 0 || eslesenBKodlari.Contains(key))
                continue;
            eslesenBKodlari.Add(key); // B içindeki tekrarlar tek satır olarak eklensin.

            sonuc.Add(new Urun
            {
                UrunKodu = bu.UrunKodu,
                StokAdeti = bCols.StokAdeti != 0 ? bu.StokAdeti : 0m,
                Birim = bCols.Birim != 0 ? bu.Birim : string.Empty,
                Fiyat = bCols.Fiyat != 0 ? bu.Fiyat : 0m,
                ParaBirimi = bCols.ParaBirimi != 0 ? bu.ParaBirimi : string.Empty,
            });
            eklendi++;
        }

        return new BirlestirmeSonucu(sonuc, dolduruldu, eklendi);
    }

    /// <summary>
    /// Sayısal (decimal) bir kolon için birleşik değeri hesaplar.
    /// <paramref name="dolduruldu"/>, yalnızca A'da olmayan bir kolon B'den
    /// gerçek (sıfırdan farklı) bir değerle doldurulduğunda true yapılır.
    /// </summary>
    private static decimal SayiBirlestir(
        bool aVar, decimal aDeger, Urun? bu, bool bVar, decimal bDeger,
        KolonKarari? karar, ref bool dolduruldu)
    {
        if (aVar && bu != null && bVar) // Çakışma: kullanıcı kararı.
        {
            return karar switch
            {
                KolonKarari.ByiKullan => bDeger,
                KolonKarari.Topla => aDeger + bDeger,
                _ => aDeger,
            };
        }

        if (aVar)
            return aDeger;

        if (bu != null && bVar) // A'da yok, B'de var: doldur.
        {
            if (bDeger != 0m)
                dolduruldu = true;
            return bDeger;
        }

        return 0m;
    }

    /// <summary>
    /// Metinsel bir kolon için birleşik değeri hesaplar. Toplama metin için
    /// geçerli olmadığından çakışmada yalnızca A veya B değeri seçilir.
    /// </summary>
    private static string MetinBirlestir(
        bool aVar, string aDeger, Urun? bu, bool bVar, string bDeger,
        KolonKarari? karar, ref bool dolduruldu)
    {
        if (aVar && bu != null && bVar) // Çakışma: kullanıcı kararı.
        {
            return karar == KolonKarari.ByiKullan ? bDeger : aDeger;
        }

        if (aVar)
            return aDeger;

        if (bu != null && bVar) // A'da yok, B'de var: doldur.
        {
            if (!string.IsNullOrWhiteSpace(bDeger))
                dolduruldu = true;
            return bDeger;
        }

        return string.Empty;
    }
}

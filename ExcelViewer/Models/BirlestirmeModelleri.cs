namespace ExcelViewer.Models;

/// <summary>
/// Bir kolon çakışmasında (aynı kolon hem açık dosya A hem seçilen dosya B'de
/// varsa) hangi değerin kullanılacağı. <see cref="Topla"/> yalnızca Stok Adeti
/// için geçerlidir.
/// </summary>
public enum KolonKarari
{
    AyiKullan = 0,
    ByiKullan = 1,
    Topla = 2,
}

/// <summary>
/// A ve B'de ortak bulunan (çakışan) veri kolonları için kullanıcı kararları.
/// Yalnızca çakışan kolonlar doldurulur; çakışmayan kolonlar null kalır ve
/// birleştirmede karar uygulanmaz.
/// </summary>
public sealed class BirlestirmeKararlari
{
    public KolonKarari? StokAdeti { get; set; }
    public KolonKarari? Birim { get; set; }
    public KolonKarari? Fiyat { get; set; }
    public KolonKarari? ParaBirimi { get; set; }
}

/// <summary>
/// Birleştirme sonucu: üretilen ürün listesi ve özet sayaçlar.
/// </summary>
public sealed record BirlestirmeSonucu(
    List<Urun> Urunler,
    int DoldurulanUrunSayisi,
    int EklenenUrunSayisi);

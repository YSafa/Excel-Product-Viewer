namespace ExcelViewer.Models;

/// <summary>
/// Aynı Ürün Kodu zaten mevcut olduğunda kullanıcının seçtiği çözüm yolu.
/// </summary>
public enum CakismaCozumu
{
    /// <summary>İşlemi iptal et, hiçbir değişiklik yapma.</summary>
    Iptal = 0,

    /// <summary>Mevcut satırın stok adetini yeni değerle değiştir.</summary>
    StoklariGuncelle = 1,

    /// <summary>Mevcut stok adetinin üzerine yeni adeti ekle.</summary>
    UzerineEkle = 2,
}

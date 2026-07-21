using System.ComponentModel;

namespace ExcelViewer.Models;

/// <summary>
/// Excel dosyasındaki bir ürün satırını temsil eder.
/// [DisplayName] öznitelikleri DataGridView'de kolon başlığı olarak görünür.
/// StokAdeti decimal'dir: "adet" ürünleri tam sayı (40), "metre" gibi ölçüler
/// küsüratlı (12.5) tutulabilsin diye. Giriş formunda birime göre tam sayı
/// kısıtı uygulanır.
/// </summary>
public sealed class Urun
{
    [DisplayName("Ürün Kodu")]
    public string UrunKodu { get; set; } = string.Empty;

    [DisplayName("Stok Adeti")]
    public decimal StokAdeti { get; set; }

    [DisplayName("Birim")]
    public string Birim { get; set; } = string.Empty;

    [DisplayName("Fiyat")]
    public decimal Fiyat { get; set; }

    [DisplayName("Para Birimi")]
    public string ParaBirimi { get; set; } = string.Empty;
}

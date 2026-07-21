namespace ExcelViewer.Models;

/// <summary>Sıralanacak kolon.</summary>
public enum SiralamaAlani
{
    Yok = 0,
    UrunKodu = 1,
    StokAdeti = 2,
    Birim = 3,
    Fiyat = 4,
    ParaBirimi = 5,
}

/// <summary>Sıralama yönü.</summary>
public enum SiralamaYonu
{
    Artan = 0,
    Azalan = 1,
}

/// <summary>Bir alan + yön çiftinden oluşan sıralama kriteri.</summary>
public sealed class SiralamaKriteri
{
    public SiralamaAlani Alan { get; set; } = SiralamaAlani.Yok;
    public SiralamaYonu Yon { get; set; } = SiralamaYonu.Artan;
}

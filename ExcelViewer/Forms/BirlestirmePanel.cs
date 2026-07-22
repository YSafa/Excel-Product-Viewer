using ExcelViewer.Models;
using Guna.UI2.WinForms;

namespace ExcelViewer.Forms;

/// <summary>
/// Birleştirmede A (açık dosya) ve B (seçilen dosya) arasında çakışan kolonlar
/// için, kolon başına tek bir seçim aldıran diyalog. Her çakışan kolon için
/// "A'yı kullan" / "B'yi kullan"; Stok Adeti için ek olarak "Topla (A + B)".
/// Varsayılan "A'yı kullan". "Uygula" ile <see cref="Sonuc"/> döner; "İptal"
/// veya kapatma birleştirmeyi durdurur.
/// </summary>
public sealed class BirlestirmePanel : Form
{
    // Combo seçenek sırası KolonKarari değerleriyle hizalı: 0=A, 1=B, 2=Topla.
    private const string SecenekA = "A'yı kullan (açık dosya)";
    private const string SecenekB = "B'yi kullan (seçilen dosya)";
    private const string SecenekTopla = "Topla (A + B)";

    private readonly IReadOnlyList<string> _cakisanKolonlar;
    private readonly Dictionary<string, Guna2ComboBox> _combolar = new();

    /// <summary>Uygula'ya basıldığında oluşan kolon kararları.</summary>
    public BirlestirmeKararlari? Sonuc { get; private set; }

    public BirlestirmePanel(IReadOnlyList<string> cakisanKolonlar)
    {
        ArgumentNullException.ThrowIfNull(cakisanKolonlar);
        _cakisanKolonlar = cakisanKolonlar;
        BuildUi();
    }

    private void BuildUi()
    {
        Text = "Kolon Çakışması";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9F);

        const int kenar = 20;
        const int alanGenislik = 300;
        int y = 16;

        var baslik = new Label
        {
            Text = "Bu kolonlar iki dosyada da var. Her biri için hangi değerin\n" +
                   "kullanılacağını seçin:",
            Location = new Point(kenar, y),
            Size = new Size(alanGenislik, 36),
            ForeColor = Color.FromArgb(60, 60, 70),
        };
        Controls.Add(baslik);
        y += 46;

        foreach (string kolon in _cakisanKolonlar)
        {
            Controls.Add(new Label
            {
                Text = kolon,
                Location = new Point(kenar, y),
                Size = new Size(alanGenislik, 18),
                Font = new Font("Segoe UI Semibold", 9F),
                ForeColor = Color.FromArgb(60, 60, 70),
            });
            y += 22;

            var combo = new Guna2ComboBox
            {
                Location = new Point(kenar, y),
                Size = new Size(alanGenislik, 36),
                BorderRadius = 6,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.5F),
            };
            combo.Items.Add(SecenekA);
            combo.Items.Add(SecenekB);
            if (kolon == "Stok Adeti")
                combo.Items.Add(SecenekTopla);
            combo.SelectedIndex = 0; // Varsayılan: A'yı kullan.

            _combolar[kolon] = combo;
            Controls.Add(combo);
            y += 48;
        }

        y += 4;
        var btnUygula = new Guna2Button
        {
            Text = "Uygula",
            Size = new Size(145, 38),
            Location = new Point(kenar, y),
            BorderRadius = 6,
            FillColor = Color.FromArgb(94, 92, 230),
            Font = new Font("Segoe UI Semibold", 9F),
        };
        btnUygula.Click += BtnUygula_Click;

        var btnIptal = new Guna2Button
        {
            Text = "İptal",
            Size = new Size(145, 38),
            Location = new Point(kenar + 155, y),
            BorderRadius = 6,
            FillColor = Color.FromArgb(150, 150, 160),
            Font = new Font("Segoe UI Semibold", 9F),
        };
        btnIptal.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };

        Controls.Add(btnUygula);
        Controls.Add(btnIptal);

        ClientSize = new Size(alanGenislik + 2 * kenar, y + 38 + 16);
        AcceptButton = btnUygula;
        CancelButton = btnIptal;
    }

    private void BtnUygula_Click(object? sender, EventArgs e)
    {
        var kararlar = new BirlestirmeKararlari();

        foreach ((string kolon, Guna2ComboBox combo) in _combolar)
        {
            KolonKarari karar = (KolonKarari)Math.Max(combo.SelectedIndex, 0);
            switch (kolon)
            {
                case "Stok Adeti": kararlar.StokAdeti = karar; break;
                case "Birim": kararlar.Birim = karar; break;
                case "Fiyat": kararlar.Fiyat = karar; break;
                case "Para Birimi": kararlar.ParaBirimi = karar; break;
            }
        }

        Sonuc = kararlar;
        DialogResult = DialogResult.OK;
        Close();
    }
}

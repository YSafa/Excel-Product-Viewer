using ExcelViewer.Models;
using Guna.UI2.WinForms;

namespace ExcelViewer.Forms;

/// <summary>
/// Sıralama dropdown paneli: hangi alana göre ve hangi yönde (artan/azalan).
/// "Uygula" ile kapanır; seçim <see cref="Sonuc"/> üzerinden döner.
/// </summary>
public sealed class SiralamaPanel : Form
{
    private static readonly (string Etiket, SiralamaAlani Alan)[] AlanSecenekleri =
    {
        ("Sıralama yok", SiralamaAlani.Yok),
        ("Ürün Kodu", SiralamaAlani.UrunKodu),
        ("Stok Adeti", SiralamaAlani.StokAdeti),
        ("Birim", SiralamaAlani.Birim),
        ("Fiyat", SiralamaAlani.Fiyat),
        ("Para Birimi", SiralamaAlani.ParaBirimi),
    };

    private Guna2ComboBox _cmbAlan = null!;
    private Guna2ComboBox _cmbYon = null!;

    private readonly SiralamaKriteri _mevcut;

    /// <summary>Uygula'ya basıldığında oluşan sıralama kriteri.</summary>
    public SiralamaKriteri? Sonuc { get; private set; }

    public SiralamaPanel(SiralamaKriteri mevcut)
    {
        ArgumentNullException.ThrowIfNull(mevcut);
        _mevcut = mevcut;
        BuildUi();
    }

    private void BuildUi()
    {
        Text = "Sıralama";
        // Yükseklik iki combo + iki başlık + buton + boşlukları tam saracak şekilde.
        // Genişlik önceki 290'a (10px büyütülmüştü) 5px daha eklenerek 295'e çıkarıldı.
        Size = new Size(295, 250);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.Manual;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9F);

        const int x = 16;
        const int alanGenislik = 263; // Panel 5px büyüdü; iç alanlar da aynı oranda genişledi.
        int y = 14;

        // --- Alan ------------------------------------------------------------
        AddLabel("Alan", x, y);
        y += 22;
        _cmbAlan = new Guna2ComboBox
        {
            Location = new Point(x, y),
            Size = new Size(alanGenislik, 36),
            BorderRadius = 6,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 9.5F),
        };
        foreach ((string etiket, _) in AlanSecenekleri)
            _cmbAlan.Items.Add(etiket);
        _cmbAlan.SelectedIndex = IndexOfAlan(_mevcut.Alan);
        Controls.Add(_cmbAlan);
        y += 48;

        // --- Yön -------------------------------------------------------------
        AddLabel("Yön", x, y);
        y += 22;
        _cmbYon = new Guna2ComboBox
        {
            Location = new Point(x, y),
            Size = new Size(alanGenislik, 36),
            BorderRadius = 6,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 9.5F),
        };
        _cmbYon.Items.Add("Artan");
        _cmbYon.Items.Add("Azalan");
        _cmbYon.SelectedIndex = _mevcut.Yon == SiralamaYonu.Azalan ? 1 : 0;
        Controls.Add(_cmbYon);
        y += 52;

        // --- Uygula ----------------------------------------------------------
        var btnUygula = new Guna2Button
        {
            Text = "Uygula",
            Size = new Size(alanGenislik, 38),
            Location = new Point(x, y),
            BorderRadius = 6,
            FillColor = Color.FromArgb(94, 92, 230),
            Font = new Font("Segoe UI Semibold", 9F),
        };
        btnUygula.Click += BtnUygula_Click;
        Controls.Add(btnUygula);
    }

    private void AddLabel(string text, int x, int y)
    {
        Controls.Add(new Label
        {
            Text = text,
            Location = new Point(x, y),
            Size = new Size(200, 18),
            Font = new Font("Segoe UI Semibold", 9F),
            ForeColor = Color.FromArgb(60, 60, 70),
        });
    }

    private static int IndexOfAlan(SiralamaAlani alan)
    {
        for (int i = 0; i < AlanSecenekleri.Length; i++)
        {
            if (AlanSecenekleri[i].Alan == alan)
                return i;
        }

        return 0;
    }

    private void BtnUygula_Click(object? sender, EventArgs e)
    {
        int alanIndex = _cmbAlan.SelectedIndex >= 0 ? _cmbAlan.SelectedIndex : 0;

        Sonuc = new SiralamaKriteri
        {
            Alan = AlanSecenekleri[alanIndex].Alan,
            Yon = _cmbYon.SelectedIndex == 1 ? SiralamaYonu.Azalan : SiralamaYonu.Artan,
        };

        DialogResult = DialogResult.OK;
        Close();
    }
}

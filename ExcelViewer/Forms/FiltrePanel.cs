using System.Globalization;
using ExcelViewer.Models;
using Guna.UI2.WinForms;

namespace ExcelViewer.Forms;

/// <summary>
/// Filtre dropdown paneli. 4 bölüm: stok aralığı, birim seçimi (checkbox),
/// fiyat aralığı, para birimi seçimi (checkbox). "Uygula" ile kapanır ve
/// seçilen kriter <see cref="Sonuc"/> üzerinden döner.
/// Bir Form olarak, çağıran butonun hemen altında açılır.
/// </summary>
public sealed class FiltrePanel : Form
{
    private Guna2TextBox _txtStokMin = null!;
    private Guna2TextBox _txtStokMax = null!;
    private Guna2TextBox _txtFiyatMin = null!;
    private Guna2TextBox _txtFiyatMax = null!;
    private CheckedListBox _clbBirimler = null!;
    private CheckedListBox _clbParaBirimleri = null!;

    private readonly FiltreKriteri _mevcut;
    private readonly List<string> _birimSecenekleri;
    private readonly List<string> _paraBirimiSecenekleri;

    /// <summary>Uygula'ya basıldığında oluşan filtre.</summary>
    public FiltreKriteri? Sonuc { get; private set; }

    public FiltrePanel(
        IReadOnlyList<Urun> urunler, FiltreKriteri mevcutFiltre)
    {
        ArgumentNullException.ThrowIfNull(urunler);
        ArgumentNullException.ThrowIfNull(mevcutFiltre);

        _mevcut = mevcutFiltre;

        _birimSecenekleri = urunler
            .Select(u => u.Birim.Trim())
            .Where(b => b.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(b => b, StringComparer.OrdinalIgnoreCase)
            .ToList();

        _paraBirimiSecenekleri = urunler
            .Select(u => u.ParaBirimi.Trim())
            .Where(p => p.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
            .ToList();

        BuildUi();
    }

    private void BuildUi()
    {
        Text = "Filtre";
        Size = new Size(320, 452);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.Manual;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9F);

        const int x = 16;
        int y = 14;

        // --- 1) Stok aralığı -------------------------------------------------
        AddBaslik("Stok Aralığı", x, ref y);
        _txtStokMin = AralikKutusu(x, y, "Min");
        _txtStokMax = AralikKutusu(x + 140, y, "Max");
        if (_mevcut.StokMin is { } smin) _txtStokMin.Text = smin.ToString(CultureInfo.InvariantCulture);
        if (_mevcut.StokMax is { } smax) _txtStokMax.Text = smax.ToString(CultureInfo.InvariantCulture);
        y += 44;

        // --- 2) Birimler -----------------------------------------------------
        AddBaslik("Birimler", x, ref y);
        _clbBirimler = CheckListe(x, y, _birimSecenekleri, _mevcut.SeciliBirimler);
        y += 90;

        // --- 3) Fiyat aralığı ------------------------------------------------
        AddBaslik("Fiyat Aralığı", x, ref y);
        _txtFiyatMin = AralikKutusu(x, y, "Min");
        _txtFiyatMax = AralikKutusu(x + 140, y, "Max");
        if (_mevcut.FiyatMin is { } fmin) _txtFiyatMin.Text = fmin.ToString(CultureInfo.InvariantCulture);
        if (_mevcut.FiyatMax is { } fmax) _txtFiyatMax.Text = fmax.ToString(CultureInfo.InvariantCulture);
        y += 44;

        // --- 4) Para birimleri ----------------------------------------------
        AddBaslik("Para Birimleri", x, ref y);
        _clbParaBirimleri = CheckListe(x, y, _paraBirimiSecenekleri, _mevcut.SeciliParaBirimleri);
        y += 90;

        // --- Butonlar --------------------------------------------------------
        var btnTemizle = new Guna2Button
        {
            Text = "Temizle",
            Size = new Size(130, 34),
            Location = new Point(x, y),
            BorderRadius = 6,
            FillColor = Color.FromArgb(150, 150, 160),
            Font = new Font("Segoe UI Semibold", 9F),
        };
        btnTemizle.Click += (_, _) => { Sonuc = new FiltreKriteri(); DialogResult = DialogResult.OK; Close(); };

        var btnUygula = new Guna2Button
        {
            Text = "Uygula",
            Size = new Size(130, 34),
            Location = new Point(x + 140, y),
            BorderRadius = 6,
            FillColor = Color.FromArgb(94, 92, 230),
            Font = new Font("Segoe UI Semibold", 9F),
        };
        btnUygula.Click += BtnUygula_Click;

        Controls.Add(btnTemizle);
        Controls.Add(btnUygula);
    }

    private void AddBaslik(string text, int x, ref int y)
    {
        var label = new Label
        {
            Text = text,
            Location = new Point(x, y),
            Size = new Size(240, 18),
            Font = new Font("Segoe UI Semibold", 9F),
            ForeColor = Color.FromArgb(60, 60, 70),
        };
        Controls.Add(label);
        y += 22;
    }

    private Guna2TextBox AralikKutusu(int x, int y, string placeholder)
    {
        var box = new Guna2TextBox
        {
            Location = new Point(x, y),
            Size = new Size(124, 32),
            BorderRadius = 6,
            PlaceholderText = placeholder,
            Font = new Font("Segoe UI", 9.5F),
        };
        box.KeyPress += OndalikKeyPress;
        Controls.Add(box);
        return box;
    }

    private CheckedListBox CheckListe(
        int x, int y, List<string> secenekler, HashSet<string> seciliOlanlar)
    {
        var clb = new CheckedListBox
        {
            Location = new Point(x, y),
            Size = new Size(264, 80),
            BorderStyle = BorderStyle.FixedSingle,
            CheckOnClick = true,
            Font = new Font("Segoe UI", 9.5F),
            IntegralHeight = false,
        };

        foreach (string secenek in secenekler)
        {
            bool secili = seciliOlanlar.Contains(secenek);
            clb.Items.Add(secenek, secili);
        }

        Controls.Add(clb);
        return clb;
    }

    private void BtnUygula_Click(object? sender, EventArgs e)
    {
        var kriter = new FiltreKriteri
        {
            StokMin = ParseNullable(_txtStokMin.Text),
            StokMax = ParseNullable(_txtStokMax.Text),
            FiyatMin = ParseNullable(_txtFiyatMin.Text),
            FiyatMax = ParseNullable(_txtFiyatMax.Text),
        };

        foreach (string? secili in _clbBirimler.CheckedItems)
        {
            if (secili != null)
                kriter.SeciliBirimler.Add(secili);
        }

        foreach (string? secili in _clbParaBirimleri.CheckedItems)
        {
            if (secili != null)
                kriter.SeciliParaBirimleri.Add(secili);
        }

        Sonuc = kriter;
        DialogResult = DialogResult.OK;
        Close();
    }

    private static decimal? ParseNullable(string raw)
    {
        raw = raw.Trim().Replace(',', '.');
        if (raw.Length == 0)
            return null;

        return decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal v)
            ? v
            : null;
    }

    private void OndalikKeyPress(object? sender, KeyPressEventArgs e)
    {
        if (char.IsControl(e.KeyChar) || char.IsDigit(e.KeyChar))
            return;

        if (e.KeyChar is '.' or ',')
        {
            if (sender is Guna2TextBox box && (box.Text.Contains('.') || box.Text.Contains(',')))
                e.Handled = true;
            return;
        }

        e.Handled = true;
    }
}

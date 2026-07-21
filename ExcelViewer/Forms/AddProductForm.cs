using System.Globalization;
using ExcelViewer.Models;
using Guna.UI2.WinForms;

namespace ExcelViewer.Forms;

/// <summary>
/// Ürün bilgilerini toplayan diyalog. İki modda çalışır:
/// <list type="bullet">
/// <item>Ekleme: boş form; girilen kod mevcut bir ürüne aitse Birim ve Para
/// Birimi otomatik doldurulup kilitlenir.</item>
/// <item>Düzenleme: alanlar seçilen satırın verisiyle dolu gelir; Kaydet
/// mevcut satırı günceller, yeni satır eklemez. Otomatik kilit uygulanmaz.</item>
/// </list>
/// Düzen: 1. satır Ürün Kodu, 2. satır Stok Adeti + Birim combo,
/// 3. satır Fiyat + Para Birimi combo.
/// Combo'lar Excel'deki değerler + sabit çekirdek listeyle beslenir ve
/// kullanıcı yeni değer de yazabilir.
/// </summary>
public sealed class AddProductForm : Form
{
    // Combo'lara her zaman eklenecek sabit çekirdek değerler.
    private static readonly string[] SabitBirimler = { "m", "adet" };
    private static readonly string[] SabitParaBirimleri = { "₺", "$", "€", "£", "¥" };

    // Kod -> mevcut ürün eşlemesi (çakışma kontrolü ve otomatik doldurma için).
    private readonly Dictionary<string, Urun> _mevcutUrunler;

    private Guna2TextBox _txtUrunKodu = null!;
    private Guna2TextBox _txtStokAdeti = null!;
    private Guna2ComboBox _cmbBirim = null!;
    private Guna2TextBox _txtFiyat = null!;
    private Guna2ComboBox _cmbParaBirimi = null!;
    private Guna2Button _btnKaydet = null!;
    private Guna2Button _btnIptal = null!;
    private Guna2HtmlLabel _lblBilgi = null!;

    // Kod mevcut bir ürünle eşleştiğinde true; combo'lar kilitlenir.
    private bool _mevcutUrunKilidi;

    // Kaydet'e basıldığında çağrılan işlev (ekleme veya güncelleme). Form kapanmadan
    // önce çalışır; true dönerse form kapanır, false dönerse (ör. dosya açık) form
    // açık kalır ve kullanıcının girdiği bilgiler korunur.
    private readonly Func<Urun, Task<bool>> _kaydet;

    /// <summary>Yeni ürün ekleme modu.</summary>
    public AddProductForm(IReadOnlyList<Urun> mevcutUrunler, Func<Urun, Task<bool>> kaydet)
    {
        ArgumentNullException.ThrowIfNull(mevcutUrunler);
        ArgumentNullException.ThrowIfNull(kaydet);

        _kaydet = kaydet;

        // Aynı kod birden fazla kez geçerse ilk görülen satırı esas al.
        _mevcutUrunler = new Dictionary<string, Urun>(StringComparer.OrdinalIgnoreCase);
        foreach (Urun urun in mevcutUrunler)
        {
            if (urun.UrunKodu.Length > 0 && !_mevcutUrunler.ContainsKey(urun.UrunKodu))
                _mevcutUrunler[urun.UrunKodu] = urun;
        }

        BuildUi(mevcutUrunler);
    }

    /// <summary>
    /// Mevcut bir ürünü düzenleme modu. Alanlar <paramref name="duzenlenecek"/>
    /// verisiyle dolu gelir; Kaydet'e basılınca <paramref name="guncelle"/> çağrılır.
    /// </summary>
    public AddProductForm(
        IReadOnlyList<Urun> mevcutUrunler, Urun duzenlenecek, Func<Urun, Task<bool>> guncelle)
        : this(mevcutUrunler, guncelle)
    {
        ArgumentNullException.ThrowIfNull(duzenlenecek);

        // Düzenlemede kodu değiştirmek meşru; otomatik doldur/kilit devre dışı.
        _txtUrunKodu.Leave -= TxtUrunKodu_Leave;

        Text = "Ürünü Düzenle";
        _lblBilgi.Text = "Mevcut ürün düzenleniyor.";
        _lblBilgi.ForeColor = Color.FromArgb(60, 60, 70);

        _txtUrunKodu.Text = duzenlenecek.UrunKodu;
        _txtStokAdeti.Text = SayiBicim(duzenlenecek.StokAdeti);
        _cmbBirim.Text = duzenlenecek.Birim;
        _txtFiyat.Text = SayiBicim(duzenlenecek.Fiyat);
        _cmbParaBirimi.Text = duzenlenecek.ParaBirimi;
    }

    /// <summary>Decimal'i düzenleme kutusuna basmak için kültürden bağımsız biçimler.</summary>
    private static string SayiBicim(decimal deger) =>
        deger.ToString(CultureInfo.InvariantCulture);

    private void BuildUi(IReadOnlyList<Urun> mevcutUrunler)
    {
        Text = "Yeni Ürün Ekle";
        Size = new Size(440, 340);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(245, 246, 250);
        Font = new Font("Segoe UI", 9.5F);

        const int labelX = 24;
        const int rowStartY = 24;
        const int rowHeight = 60;
        int y = rowStartY;

        // Combo'lar %33 daraltıldı (132 -> 88); soldaki sayısal alan buna göre genişler.
        const int comboGenislik = 88;
        const int araBosluk = 16;
        const int comboX = labelX + 304;              // Sağ blok başlangıcı.
        const int sagKenar = comboX + comboGenislik;  // Tüm alanların ortak sağ kenarı.
        const int solAlanGenislik = comboX - labelX - araBosluk;

        // Kilitlenen (Enabled = false) combo'larda metin dikeyde kesik görünüyordu:
        // Guna2ComboBox metni tüm yüksekliğe ortalayıp yalnızca sığdığında çiziyor.
        // ItemHeight'i artırınca kontrol yükselir (Yükseklik = ItemHeight + 6),
        // metne dikey nefes payı açılır. Yandaki 36px'lik metin kutularıyla
        // ortalı dursun diye combo'lar 2px yukarıdan başlatılır.
        const int comboItemHeight = 34;      // -> combo yüksekliği 40.
        const int comboYukseklik = comboItemHeight + 6;
        const int comboUstFark = (comboYukseklik - 36) / 2; // Metin kutusuyla ortalama.

        // --- 1. satır: Ürün Kodu (combo'ların sağ kenarıyla hizalı) ----------
        AddLabel("Ürün Kodu", labelX, y);
        _txtUrunKodu = new Guna2TextBox
        {
            Location = new Point(labelX, y + 24),
            Size = new Size(sagKenar - labelX, 36),
            BorderRadius = 6,
            Font = new Font("Segoe UI", 10F),
        };
        _txtUrunKodu.Leave += TxtUrunKodu_Leave;
        Controls.Add(_txtUrunKodu);
        y += rowHeight;

        // --- 2. satır: Stok Adeti + Birim combo ------------------------------
        AddLabel("Stok Adeti", labelX, y);
        _txtStokAdeti = new Guna2TextBox
        {
            Location = new Point(labelX, y + 24),
            Size = new Size(solAlanGenislik, 36),
            BorderRadius = 6,
            Font = new Font("Segoe UI", 10F),
        };
        // Stok girişi birime göre: "adet" ise tam sayı, diğer birimlerde ondalık.
        _txtStokAdeti.KeyPress += StokKeyPress;
        Controls.Add(_txtStokAdeti);

        AddLabel("Birim", comboX, y);
        _cmbBirim = new Guna2ComboBox
        {
            Location = new Point(comboX, y + 24 - comboUstFark),
            Size = new Size(comboGenislik, comboYukseklik),
            ItemHeight = comboItemHeight,
            BorderRadius = 6,
            DropDownStyle = ComboBoxStyle.DropDown, // Seçilebilir + yazılabilir.
            Font = new Font("Segoe UI", 10F),
        };
        DoldurCombo(_cmbBirim, SabitBirimler, mevcutUrunler.Select(u => u.Birim));
        Controls.Add(_cmbBirim);
        y += rowHeight;

        // --- 3. satır: Fiyat + Para Birimi combo -----------------------------
        AddLabel("Fiyat", labelX, y);
        _txtFiyat = new Guna2TextBox
        {
            Location = new Point(labelX, y + 24),
            Size = new Size(solAlanGenislik, 36),
            BorderRadius = 6,
            Font = new Font("Segoe UI", 10F),
        };
        // Sadece ondalık sayı: harf engelli, tek ondalık ayıraç serbest.
        _txtFiyat.KeyPress += OndalikKeyPress;
        Controls.Add(_txtFiyat);

        AddLabel("Para Birimi", comboX, y);
        _cmbParaBirimi = new Guna2ComboBox
        {
            Location = new Point(comboX, y + 24 - comboUstFark),
            Size = new Size(comboGenislik, comboYukseklik),
            ItemHeight = comboItemHeight,
            BorderRadius = 6,
            DropDownStyle = ComboBoxStyle.DropDown,
            Font = new Font("Segoe UI", 10F),
        };
        DoldurCombo(_cmbParaBirimi, SabitParaBirimleri, mevcutUrunler.Select(u => u.ParaBirimi));
        Controls.Add(_cmbParaBirimi);
        y += rowHeight - 8;

        // --- Bilgi etiketi (kilit durumunu gösterir) -------------------------
        _lblBilgi = new Guna2HtmlLabel
        {
            Text = string.Empty,
            Location = new Point(labelX, y),
            Size = new Size(392, 20),
            ForeColor = Color.FromArgb(180, 90, 30),
        };
        Controls.Add(_lblBilgi);
        y += 26;

        // --- Butonlar --------------------------------------------------------
        _btnKaydet = new Guna2Button
        {
            Text = "Kaydet",
            Size = new Size(120, 40),
            Location = new Point(labelX + 32, y),
            BorderRadius = 8,
            FillColor = Color.FromArgb(94, 92, 230),
            Font = new Font("Segoe UI Semibold", 9.5F),
        };
        _btnKaydet.Click += BtnKaydet_Click;

        _btnIptal = new Guna2Button
        {
            Text = "İptal",
            Size = new Size(120, 40),
            Location = new Point(labelX + 168, y),
            BorderRadius = 8,
            FillColor = Color.FromArgb(150, 150, 160),
            Font = new Font("Segoe UI Semibold", 9.5F),
        };
        _btnIptal.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };

        Controls.Add(_btnKaydet);
        Controls.Add(_btnIptal);

        AcceptButton = _btnKaydet;
        CancelButton = _btnIptal;
        ActiveControl = _txtUrunKodu;
    }

    private void AddLabel(string text, int x, int y)
    {
        var label = new Label
        {
            Text = text,
            Location = new Point(x, y),
            Size = new Size(180, 20),
            ForeColor = Color.FromArgb(60, 60, 70),
        };
        Controls.Add(label);
    }

    /// <summary>
    /// Combo'yu sabit değerler + Excel'den gelen değerlerle, tekrarsız doldurur.
    /// </summary>
    private static void DoldurCombo(
        Guna2ComboBox combo, IEnumerable<string> sabitler, IEnumerable<string> excelDegerleri)
    {
        var gorulen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sirali = new List<string>();

        void Ekle(string deger)
        {
            string temiz = deger.Trim();
            if (temiz.Length == 0)
                return;
            if (gorulen.Add(temiz))
                sirali.Add(temiz);
        }

        foreach (string s in sabitler)
            Ekle(s);
        foreach (string s in excelDegerleri)
            Ekle(s);

        combo.Items.AddRange(sirali.Cast<object>().ToArray());
    }

    /// <summary>
    /// Ürün Kodu alanından çıkıldığında, kod mevcut bir ürüne aitse
    /// Birim ve Para Birimi'ni otomatik doldurup kilitler; kod boş veya
    /// bilinmeyen bir değerse kilidi açar.
    /// </summary>
    private void TxtUrunKodu_Leave(object? sender, EventArgs e)
    {
        string kod = _txtUrunKodu.Text.Trim();

        if (kod.Length > 0 && _mevcutUrunler.TryGetValue(kod, out Urun? mevcut))
        {
            _cmbBirim.Text = mevcut.Birim;
            _cmbParaBirimi.Text = mevcut.ParaBirimi;
            KilitleSabitAlanlar(true);
            _lblBilgi.Text = "Bu ürün mevcut. Birim ve Para Birimi otomatik dolduruldu.";
        }
        else if (_mevcutUrunKilidi)
        {
            // Kod temizlendi veya değişti: kilidi aç ve otomatik değerleri temizle.
            KilitleSabitAlanlar(false);
            _cmbBirim.Text = string.Empty;
            _cmbParaBirimi.Text = string.Empty;
            _lblBilgi.Text = string.Empty;
        }
    }

    private void KilitleSabitAlanlar(bool kilitli)
    {
        _mevcutUrunKilidi = kilitli;
        _cmbBirim.Enabled = !kilitli;
        _cmbParaBirimi.Enabled = !kilitli;
    }

    /// <summary>
    /// Stok girişi: birim "adet" ise sadece tam sayı; diğer birimlerde
    /// (metre vb.) tek ondalık ayıraçlı sayı serbest.
    /// </summary>
    private void StokKeyPress(object? sender, KeyPressEventArgs e)
    {
        if (char.IsControl(e.KeyChar))
            return;

        bool adetBirimi = _cmbBirim.Text.Trim().Equals("adet", StringComparison.OrdinalIgnoreCase);

        if (char.IsDigit(e.KeyChar))
            return;

        // "adet" değilse tek ondalık ayıraca izin ver.
        if (!adetBirimi && e.KeyChar is '.' or ',')
        {
            if (_txtStokAdeti.Text.Contains('.') || _txtStokAdeti.Text.Contains(','))
                e.Handled = true;
            return;
        }

        e.Handled = true;
    }

    /// <summary>
    /// Rakam ve tek bir ondalık ayıraca (nokta veya virgül) izin verir.
    /// Zaten bir ayıraç varsa ikincisini engeller.
    /// </summary>
    private void OndalikKeyPress(object? sender, KeyPressEventArgs e)
    {
        if (char.IsControl(e.KeyChar))
            return;

        if (char.IsDigit(e.KeyChar))
            return;

        if (e.KeyChar is '.' or ',')
        {
            string mevcut = _txtFiyat.Text;
            if (mevcut.Contains('.') || mevcut.Contains(','))
                e.Handled = true; // Zaten bir ondalık ayıraç var.
            return;
        }

        e.Handled = true; // Diğer her şey engellenir.
    }

    private async void BtnKaydet_Click(object? sender, EventArgs e)
    {
        string urunKodu = _txtUrunKodu.Text.Trim();
        if (urunKodu.Length == 0)
        {
            Uyar("Ürün Kodu boş olamaz.");
            ActiveControl = _txtUrunKodu;
            return;
        }

        if (!TryParseDecimal(_txtStokAdeti.Text, out decimal stokAdeti))
        {
            Uyar("Stok Adeti geçerli bir sayı olmalıdır.");
            ActiveControl = _txtStokAdeti;
            return;
        }

        // "adet" biriminde küsürat kabul edilmez.
        bool adetBirimi = _cmbBirim.Text.Trim().Equals("adet", StringComparison.OrdinalIgnoreCase);
        if (adetBirimi && stokAdeti != decimal.Truncate(stokAdeti))
        {
            Uyar("'adet' biriminde stok tam sayı olmalıdır.");
            ActiveControl = _txtStokAdeti;
            return;
        }

        if (!TryParseDecimal(_txtFiyat.Text, out decimal fiyat))
        {
            Uyar("Fiyat geçerli bir sayı olmalıdır.");
            ActiveControl = _txtFiyat;
            return;
        }

        var urun = new Urun
        {
            UrunKodu = urunKodu,
            StokAdeti = stokAdeti,
            Birim = _cmbBirim.Text.Trim(),
            Fiyat = fiyat,
            ParaBirimi = _cmbParaBirimi.Text.Trim(),
        };

        // Kaydetmeyi dene. Başarısızsa (ör. Excel açık) form açık kalır ve
        // kullanıcının girdiği bilgiler korunur.
        _btnKaydet.Enabled = false;
        try
        {
            bool basarili = await _kaydet(urun);
            if (basarili)
            {
                DialogResult = DialogResult.OK;
                Close();
            }
        }
        finally
        {
            if (!IsDisposed && !Disposing)
                _btnKaydet.Enabled = true;
        }
    }

    private static bool TryParseDecimal(string raw, out decimal value)
    {
        raw = raw.Trim();
        if (raw.Length == 0)
        {
            value = 0m;
            return true; // Boş fiyat 0 kabul edilir.
        }

        // Hem virgül hem nokta ondalık ayıracını tolere et.
        raw = raw.Replace(',', '.');
        return decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
    }

    private void Uyar(string mesaj)
    {
        MessageBox.Show(this, mesaj, "Geçersiz Giriş",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}

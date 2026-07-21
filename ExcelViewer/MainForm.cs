using System.ComponentModel;
using ExcelViewer.Forms;
using ExcelViewer.Models;
using ExcelViewer.Services;
using Guna.UI2.WinForms;

namespace ExcelViewer;

/// <summary>
/// Ana pencere: Excel dosyası seç, ürünleri görüntüle, ara, filtrele,
/// sırala ve yeni ürün ekle.
/// </summary>
public sealed class MainForm : Form
{
    private const int SearchDebounceMs = 200;

    private readonly ExcelReaderService _excelReader = new();
    private readonly LogService _logService = new();
    private readonly ExcelWriterService _excelWriter;
    private readonly BindingSource _bindingSource = new();
    private readonly System.Windows.Forms.Timer _searchDebounceTimer = new();

    // Aktif filtre ve sıralama durumu (kalıcı; panel kapansa da korunur).
    private readonly FiltreKriteri _filtre = new();
    private readonly SiralamaKriteri _siralama = new();

    private List<Urun>? _allProducts;
    private string? _currentFilePath;

    // Kolon hizalama/genişlik ayarı yalnızca ilk veri bağlamada uygulanır;
    // sonrasında kullanıcının elle yaptığı genişlik değişiklikleri korunur.
    private bool _kolonlarAyarlandi;

    private Guna2Button _btnSelectFile = null!;
    private Guna2Button _btnAddProduct = null!;
    private Guna2TextBox _txtSearch = null!;
    private Guna2Button _btnFiltre = null!;
    private Guna2Button _btnSiralama = null!;
    private Guna2DataGridView _grid = null!;
    private Guna2HtmlLabel _lblStatus = null!;

    public MainForm()
    {
        _excelWriter = new ExcelWriterService(_logService);

        BuildUi();

        _searchDebounceTimer.Interval = SearchDebounceMs;
        _searchDebounceTimer.Tick += SearchDebounceTimer_Tick;

        _grid.DataSource = _bindingSource;
    }

    private void BuildUi()
    {
        Text = "Ürün Görüntüleyici";
        Size = new Size(960, 600);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(245, 246, 250);
        Font = new Font("Segoe UI", 9F);

        var topPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 64,
            Padding = new Padding(16, 14, 16, 10),
            BackColor = Color.FromArgb(245, 246, 250),
        };

        _btnSelectFile = new Guna2Button
        {
            Text = "Excel Seç",
            Size = new Size(110, 40),
            Location = new Point(16, 12),
            BorderRadius = 8,
            FillColor = Color.FromArgb(94, 92, 230),
            Font = new Font("Segoe UI Semibold", 9.5F),
        };
        _btnSelectFile.Click += BtnSelectFile_Click;

        _btnAddProduct = new Guna2Button
        {
            Text = "Ürün Ekle",
            Size = new Size(110, 40),
            Location = new Point(136, 12),
            BorderRadius = 8,
            FillColor = Color.FromArgb(46, 170, 120),
            Font = new Font("Segoe UI Semibold", 9.5F),
            Enabled = false,
        };
        _btnAddProduct.Click += BtnAddProduct_Click;

        _txtSearch = new Guna2TextBox
        {
            PlaceholderText = "Ürün kodu veya birime göre ara...",
            Size = new Size(430, 40),
            Location = new Point(256, 12),
            BorderRadius = 8,
            Enabled = false,
            Font = new Font("Segoe UI", 10F),
        };
        _txtSearch.TextChanged += TxtSearch_TextChanged;

        _btnFiltre = new Guna2Button
        {
            Text = "Filtre",
            Size = new Size(80, 40),
            Location = new Point(694, 12),
            BorderRadius = 8,
            FillColor = Color.FromArgb(120, 118, 200),
            Font = new Font("Segoe UI Semibold", 9F),
            Enabled = false,
        };
        _btnFiltre.Click += BtnFiltre_Click;

        _btnSiralama = new Guna2Button
        {
            Text = "Sırala",
            Size = new Size(80, 40),
            Location = new Point(780, 12),
            BorderRadius = 8,
            FillColor = Color.FromArgb(120, 118, 200),
            Font = new Font("Segoe UI Semibold", 9F),
            Enabled = false,
        };
        _btnSiralama.Click += BtnSiralama_Click;

        topPanel.Controls.Add(_btnSelectFile);
        topPanel.Controls.Add(_btnAddProduct);
        topPanel.Controls.Add(_txtSearch);
        topPanel.Controls.Add(_btnFiltre);
        topPanel.Controls.Add(_btnSiralama);

        _lblStatus = new Guna2HtmlLabel
        {
            Text = "Dosya yüklenmedi.",
            Dock = DockStyle.Bottom,
            Height = 28,
            Padding = new Padding(18, 6, 0, 0),
            BackColor = Color.FromArgb(245, 246, 250),
            ForeColor = Color.FromArgb(90, 90, 100),
        };

        _grid = new Guna2DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            RowHeadersVisible = false,
            BorderStyle = BorderStyle.None,
            BackgroundColor = Color.White,
        };
        StyleGrid(_grid);
        _grid.DataBindingComplete += Grid_DataBindingComplete;

        Controls.Add(_grid);
        Controls.Add(topPanel);
        Controls.Add(_lblStatus);
    }

    private static void StyleGrid(Guna2DataGridView grid)
    {
        grid.ColumnHeadersHeight = 40;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9.5F);
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(94, 92, 230);
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        grid.EnableHeadersVisualStyles = false;

        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 223, 250);
        grid.DefaultCellStyle.SelectionForeColor = Color.Black;
        grid.RowTemplate.Height = 34;
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(247, 247, 252);
        grid.GridColor = Color.FromArgb(215, 215, 225);

        // Sütunlar arası dikey çizgi belirgin olsun; satır çizgileri hafif kalsın.
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
    }

    private async void BtnSelectFile_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Excel dosyaları (*.xlsx)|*.xlsx",
            Title = "Bir Excel dosyası seçin",
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        _currentFilePath = dialog.FileName;
        _kolonlarAyarlandi = false; // Yeni dosyada kolon oranları yeniden uygulansın.
        await LoadFileAsync(_currentFilePath);
    }

    private async Task LoadFileAsync(string filePath)
    {
        SetBusy(true, "Yükleniyor...");

        try
        {
            List<Urun> products = await _excelReader.ReadAsync(filePath);
            _allProducts = products;

            ApplyFilter();

            _txtSearch.Enabled = true;
            _btnAddProduct.Enabled = true;
            _btnFiltre.Enabled = true;
            _btnSiralama.Enabled = true;
            _txtSearch.Clear();
            _lblStatus.Text = $"{Path.GetFileName(filePath)} dosyasından {products.Count} ürün yüklendi.";
        }
        catch (Exception ex)
        {
            _allProducts = null;
            _currentFilePath = null;
            _bindingSource.DataSource = null;
            _txtSearch.Enabled = false;
            _btnAddProduct.Enabled = false;
            _btnFiltre.Enabled = false;
            _btnSiralama.Enabled = false;
            _lblStatus.Text = "Dosya yüklenemedi.";

            MessageBox.Show(
                this,
                $"Excel dosyası okunamadı:{Environment.NewLine}{ex.Message}",
                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    private void BtnAddProduct_Click(object? sender, EventArgs e)
    {
        if (_currentFilePath == null || _allProducts == null)
            return;

        using var dialog = new AddProductForm(_allProducts, UrunKaydetAsync);
        dialog.ShowDialog(this);
        // Kaydetme başarılıysa UrunKaydetAsync içinde dosya yeniden okundu.
    }

    private void BtnFiltre_Click(object? sender, EventArgs e)
    {
        if (_allProducts == null)
            return;

        using var panel = new FiltrePanel(_allProducts, _filtre);
        panel.Location = PanelKonumu(_btnFiltre, panel.Size);

        if (panel.ShowDialog(this) == DialogResult.OK && panel.Sonuc != null)
        {
            KopyalaFiltre(panel.Sonuc, _filtre);
            GuncelleFiltreButonu();
            ApplyFilter();
        }
    }

    private void BtnSiralama_Click(object? sender, EventArgs e)
    {
        using var panel = new SiralamaPanel(_siralama);
        panel.Location = PanelKonumu(_btnSiralama, panel.Size);

        if (panel.ShowDialog(this) == DialogResult.OK && panel.Sonuc != null)
        {
            _siralama.Alan = panel.Sonuc.Alan;
            _siralama.Yon = panel.Sonuc.Yon;
            GuncelleSiralamaButonu();
            ApplyFilter();
        }
    }

    /// <summary>Paneli, ilgili butonun hemen altına ekran koordinatında konumlar.</summary>
    private Point PanelKonumu(Control buton, Size panelSize)
    {
        Point ekranSol = buton.PointToScreen(new Point(0, buton.Height + 4));

        // Ekran dışına taşmasın diye sağ kenarı kontrol et.
        Rectangle calisma = Screen.FromControl(this).WorkingArea;
        int x = Math.Min(ekranSol.X, calisma.Right - panelSize.Width - 8);
        int y = Math.Min(ekranSol.Y, calisma.Bottom - panelSize.Height - 8);
        return new Point(Math.Max(x, calisma.Left + 8), Math.Max(y, calisma.Top + 8));
    }

    private static void KopyalaFiltre(FiltreKriteri kaynak, FiltreKriteri hedef)
    {
        hedef.StokMin = kaynak.StokMin;
        hedef.StokMax = kaynak.StokMax;
        hedef.FiyatMin = kaynak.FiyatMin;
        hedef.FiyatMax = kaynak.FiyatMax;

        hedef.SeciliBirimler.Clear();
        foreach (string b in kaynak.SeciliBirimler)
            hedef.SeciliBirimler.Add(b);

        hedef.SeciliParaBirimleri.Clear();
        foreach (string p in kaynak.SeciliParaBirimleri)
            hedef.SeciliParaBirimleri.Add(p);
    }

    private void GuncelleFiltreButonu()
    {
        bool aktif = !_filtre.BosMu;
        _btnFiltre.Text = aktif ? "Filtre ●" : "Filtre";
        _btnFiltre.FillColor = aktif
            ? Color.FromArgb(94, 92, 230)
            : Color.FromArgb(120, 118, 200);
    }

    private void GuncelleSiralamaButonu()
    {
        bool aktif = _siralama.Alan != SiralamaAlani.Yok;
        _btnSiralama.Text = aktif ? "Sırala ●" : "Sırala";
        _btnSiralama.FillColor = aktif
            ? Color.FromArgb(94, 92, 230)
            : Color.FromArgb(120, 118, 200);
    }

    /// <summary>
    /// Ürünü Excel'e yazmayı dener. Başarılıysa true döner (form kapanır);
    /// dosya açıksa veya iptal edilirse false döner (form açık kalır).
    /// AddProductForm tarafından callback olarak çağrılır.
    /// </summary>
    private async Task<bool> UrunKaydetAsync(Urun urun)
    {
        if (_currentFilePath == null)
            return false;

        SetBusy(true, "Kaydediliyor...");

        try
        {
            YazmaSonucu sonuc = await _excelWriter.UrunEkleAsync(
                _currentFilePath, urun, CakismaCozumu.Iptal);

            if (sonuc == YazmaSonucu.CakismaVar)
            {
                CakismaCozumu cozum = CakismaSor(urun.UrunKodu);
                if (cozum == CakismaCozumu.Iptal)
                {
                    _lblStatus.Text = "İşlem iptal edildi.";
                    return false; // Form açık kalsın, kullanıcı yeniden karar verebilsin.
                }

                sonuc = await _excelWriter.UrunEkleAsync(_currentFilePath, urun, cozum);
            }

            await LoadFileAsync(_currentFilePath);

            _lblStatus.Text = sonuc switch
            {
                YazmaSonucu.Eklendi => $"'{urun.UrunKodu}' eklendi.",
                YazmaSonucu.Guncellendi => $"'{urun.UrunKodu}' güncellendi.",
                _ => _lblStatus.Text,
            };

            return true;
        }
        catch (IOException)
        {
            MessageBox.Show(
                this,
                "Excel dosyası şu anda açık görünüyor. Lütfen dosyayı kapatıp tekrar 'Kaydet'e basın." +
                Environment.NewLine + "Girdiğiniz bilgiler korundu.",
                "Dosya Kilitli", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _lblStatus.Text = "Kaydedilemedi: dosya açık.";
            return false; // Form açık kalsın.
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                $"Ürün kaydedilemedi:{Environment.NewLine}{ex.Message}",
                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _lblStatus.Text = "Kaydedilemedi.";
            return false;
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    private CakismaCozumu CakismaSor(string urunKodu)
    {
        DialogResult secim = MessageBox.Show(
            this,
            $"'{urunKodu}' koduna sahip ürün zaten mevcut." + Environment.NewLine + Environment.NewLine +
            "EVET  → Stokları güncelle (mevcut stok yeni değerle değiştirilir)" + Environment.NewLine +
            "HAYIR → Olan stokun üzerine ekle" + Environment.NewLine +
            "İPTAL → İşlemi iptal et",
            "Ürün Zaten Mevcut",
            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

        return secim switch
        {
            DialogResult.Yes => CakismaCozumu.StoklariGuncelle,
            DialogResult.No => CakismaCozumu.UzerineEkle,
            _ => CakismaCozumu.Iptal,
        };
    }

    private void TxtSearch_TextChanged(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();
    }

    private void Grid_DataBindingComplete(object? sender, DataGridViewBindingCompleteEventArgs e)
    {
        // Yalnızca ilk kez: hizalama ve başlangıç genişlik oranlarını uygula.
        // Sonraki bağlamalarda kullanıcının ayarladığı genişlikler korunur.
        if (_kolonlarAyarlandi)
            return;

        AlignColumnRight(nameof(Urun.StokAdeti));
        AlignColumnRight(nameof(Urun.Fiyat));

        // Kolon oranları: Birim ve Para Birimi ~%35 daha dar başlasın.
        SetColumnWeight(nameof(Urun.UrunKodu), 130);
        SetColumnWeight(nameof(Urun.StokAdeti), 110);
        SetColumnWeight(nameof(Urun.Birim), 65);
        SetColumnWeight(nameof(Urun.Fiyat), 110);
        SetColumnWeight(nameof(Urun.ParaBirimi), 65);

        _kolonlarAyarlandi = true;
    }

    private void SetColumnWeight(string propertyName, int weight)
    {
        foreach (DataGridViewColumn column in _grid.Columns)
        {
            if (column.DataPropertyName == propertyName)
            {
                column.FillWeight = weight;
                return;
            }
        }
    }

    private void AlignColumnRight(string propertyName)
    {
        foreach (DataGridViewColumn column in _grid.Columns)
        {
            if (column.DataPropertyName == propertyName)
            {
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                return;
            }
        }
    }

    private void SearchDebounceTimer_Tick(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        ApplyFilter();
    }

    private void ApplyFilter()
    {
        if (_allProducts == null)
            return;

        List<Urun> sonuc = UrunFilterService.Uygula(
            _allProducts, _txtSearch.Text, _filtre, _siralama);

        _bindingSource.DataSource = sonuc;
        _bindingSource.ResetBindings(false);

        _lblStatus.Text = $"{_allProducts.Count} üründen {sonuc.Count} tanesi gösteriliyor.";
    }

    private void SetBusy(bool busy, string? statusText)
    {
        _btnSelectFile.Enabled = !busy;
        _btnAddProduct.Enabled = !busy && _allProducts != null;
        _btnFiltre.Enabled = !busy && _allProducts != null;
        _btnSiralama.Enabled = !busy && _allProducts != null;
        UseWaitCursor = busy;

        if (statusText != null)
            _lblStatus.Text = statusText;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _searchDebounceTimer.Dispose();
            _bindingSource.Dispose();
        }

        base.Dispose(disposing);
    }
}

namespace ExcelViewer.Services;

/// <summary>
/// Ürün ekleme/güncelleme işlemlerini zaman damgalı olarak bir metin
/// dosyasına yazar. Log dosyası, Excel dosyasının bulunduğu klasördeki
/// "log" alt klasöründe tutulur.
/// </summary>
public sealed class LogService
{
    private const string LogFolderName = "log";
    private const string LogFileName = "urun_log.txt";

    private readonly object _writeLock = new();

    /// <summary>
    /// Verilen Excel dosyasının yanındaki log klasörüne bir satır ekler.
    /// Log yazımı asla ana işlemi (Excel'e yazmayı) bozmamalıdır; bu yüzden
    /// çağıran taraf hataları yutmayı tercih edebilir.
    /// </summary>
    public void Append(string excelFilePath, string message)
    {
        if (string.IsNullOrWhiteSpace(excelFilePath))
            throw new ArgumentException("Excel dosya yolu boş olamaz.", nameof(excelFilePath));

        string? directory = Path.GetDirectoryName(excelFilePath);
        if (string.IsNullOrEmpty(directory))
            directory = Directory.GetCurrentDirectory();

        string logFolder = Path.Combine(directory, LogFolderName);
        Directory.CreateDirectory(logFolder);

        string logPath = Path.Combine(logFolder, LogFileName);
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        string line = $"[{timestamp}] {message}{Environment.NewLine}";

        lock (_writeLock)
        {
            File.AppendAllText(logPath, line);
        }
    }
}

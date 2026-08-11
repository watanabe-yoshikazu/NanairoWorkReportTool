using System.Text.Json;

namespace NanairoWorkReportTool.Infrastructure.Persistence;

public sealed class AppSettings
{
    public string ReporterName { get; set; } = string.Empty;
    public string CompanyName { get; set; } = string.Empty;
    public string Destination { get; set; } = "株式会社ナナイロ　御中";
    public string OutputDirectory { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    public string? LastNwrFilePath { get; set; }
    public List<string> RecentFiles { get; set; } = [];
    public List<string> WorkContentHistory { get; set; } = [];
}

public sealed class SettingsStore
{
    private readonly string directory;
    private string SettingsPath => Path.Combine(directory, "settings.json");
    public string RecoveryPath => Path.Combine(directory, "recovery.nwr");

    public SettingsStore(string? directory = null)
    {
        this.directory = directory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NanairoWorkReportTool");
    }

    public async Task<AppSettings> LoadAsync()
    {
        if (!File.Exists(SettingsPath)) return new AppSettings();
        try
        {
            await using var stream = File.OpenRead(SettingsPath);
            return await JsonSerializer.DeserializeAsync<AppSettings>(stream) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public async Task SaveAsync(AppSettings settings)
    {
        Directory.CreateDirectory(directory);
        await using var stream = File.Create(SettingsPath);
        await JsonSerializer.SerializeAsync(stream, settings, new JsonSerializerOptions { WriteIndented = true });
    }
}

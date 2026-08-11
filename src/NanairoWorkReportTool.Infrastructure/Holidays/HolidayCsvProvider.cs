using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using NanairoWorkReportTool.Core.Services;

namespace NanairoWorkReportTool.Infrastructure.Holidays;

public sealed class HolidayCsvProvider : IHolidayProvider, IDisposable
{
    private const string SourceUrl = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv";
    private readonly HttpClient httpClient;
    private readonly string cacheDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NanairoWorkReportTool");
    private string CachePath => Path.Combine(cacheDirectory, "holidays.csv");

    public HolidayCsvProvider(HttpClient? httpClient = null)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        this.httpClient = httpClient ?? new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        this.httpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("NanairoWorkReportTool", "1.0"));
    }

    public async Task<IReadOnlyDictionary<DateOnly, string>> GetHolidaysAsync(int year, CancellationToken cancellationToken = default)
    {
        var text = File.Exists(CachePath)
            ? await File.ReadAllTextAsync(CachePath, DetectEncoding(CachePath), cancellationToken)
            : await ReadEmbeddedAsync(cancellationToken);
        return Parse(text).Where(pair => pair.Key.Year == year).ToDictionary();
    }

    public async Task<bool> UpdateAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var bytes = await httpClient.GetByteArrayAsync(SourceUrl, cancellationToken);
            var text = Decode(bytes);
            var parsed = Parse(text);
            if (parsed.Count == 0) return false;
            Directory.CreateDirectory(cacheDirectory);
            await File.WriteAllTextAsync(CachePath, text, new UTF8Encoding(false), cancellationToken);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static Dictionary<DateOnly, string> Parse(string csv)
    {
        var result = new Dictionary<DateOnly, string>();
        foreach (var rawLine in csv.Replace("\r", string.Empty).Split('\n').Skip(1))
        {
            if (string.IsNullOrWhiteSpace(rawLine)) continue;
            var comma = rawLine.IndexOf(',');
            if (comma < 0) continue;
            var dateText = rawLine[..comma].Trim('"', ' ', '\uFEFF');
            var name = rawLine[(comma + 1)..].Trim('"', ' ');
            if (DateOnly.TryParse(dateText, out var date) && !string.IsNullOrWhiteSpace(name)) result[date] = name;
        }
        return result;
    }

    private static async Task<string> ReadEmbeddedAsync(CancellationToken cancellationToken)
    {
        var assembly = typeof(HolidayCsvProvider).Assembly;
        await using var stream = assembly.GetManifestResourceStream("NanairoWorkReportTool.Infrastructure.Assets.holidays.csv")
                                 ?? throw new FileNotFoundException("同梱の祝日データがありません。");
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return await reader.ReadToEndAsync(cancellationToken);
    }

    private static Encoding DetectEncoding(string path)
    {
        using var stream = File.OpenRead(path);
        Span<byte> bom = stackalloc byte[3];
        var read = stream.Read(bom);
        return read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF
            ? Encoding.UTF8
            : Encoding.UTF8;
    }

    private static string Decode(byte[] bytes)
    {
        try
        {
            var utf8 = new UTF8Encoding(false, true);
            return utf8.GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return Encoding.GetEncoding(932).GetString(bytes);
        }
    }

    public void Dispose() => httpClient.Dispose();
}


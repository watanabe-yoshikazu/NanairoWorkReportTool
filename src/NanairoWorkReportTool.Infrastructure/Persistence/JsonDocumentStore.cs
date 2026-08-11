using System.Text.Json;
using System.Text.Json.Serialization;
using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Core.Services;

namespace NanairoWorkReportTool.Infrastructure.Persistence;

public sealed class JsonDocumentStore : IDocumentStore
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    public async Task SaveAsync(string path, WorkReportDocument document, CancellationToken cancellationToken = default)
    {
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var temporaryPath = fullPath + $".{Guid.NewGuid():N}.tmp";
        try
        {
            await using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, true))
            {
                await JsonSerializer.SerializeAsync(stream, document, Options, cancellationToken);
                await stream.FlushAsync(cancellationToken);
            }
            File.Move(temporaryPath, fullPath, true);
        }
        finally
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    public async Task<WorkReportDocument> LoadAsync(string path, CancellationToken cancellationToken = default)
    {
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true);
        var document = await JsonSerializer.DeserializeAsync<WorkReportDocument>(stream, Options, cancellationToken)
                       ?? throw new InvalidDataException("保存データを読み取れませんでした。");
        if (document.SchemaVersion != WorkReportDocument.CurrentSchemaVersion)
            throw new InvalidDataException($"未対応の保存形式です（バージョン {document.SchemaVersion}）。");
        return document;
    }

    internal static string Serialize(WorkReportDocument document) => JsonSerializer.Serialize(document, Options);
    internal static WorkReportDocument Deserialize(string json)
        => JsonSerializer.Deserialize<WorkReportDocument>(json, Options)
           ?? throw new InvalidDataException("帳票メタデータを読み取れませんでした。");
}

using System.Text;
using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Infrastructure.Persistence;

namespace NanairoWorkReportTool.Infrastructure.Tests;

public sealed class PersistenceTests
{
    [Fact]
    public async Task Nwr_RoundTripsUtf8Json()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"NanairoTests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "作業報告.nwr");
            var source = new WorkReportDocument { TargetMonth = new DateTime(2026, 7, 1), ReporterName = "渡辺", CompanyName = "株式会社リンク" };
            source.Entries.Add(new WorkDayEntry { Date = new DateOnly(2026, 7, 1), WorkContent = "日本語の作業内容", WorkStatus = WorkStatus.Normal });
            var store = new JsonDocumentStore();
            await store.SaveAsync(path, source);
            var restored = await store.LoadAsync(path);
            Assert.Equal(WorkReportDocument.CurrentSchemaVersion, restored.SchemaVersion);
            Assert.Equal("日本語の作業内容", restored.Entries[0].WorkContent);
            Assert.Contains("schemaVersion", await File.ReadAllTextAsync(path, Encoding.UTF8));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task Nwr_RejectsUnknownSchema()
    {
        var path = Path.GetTempFileName();
        try
        {
            await File.WriteAllTextAsync(path, "{\"schemaVersion\":999}", Encoding.UTF8);
            await Assert.ThrowsAsync<InvalidDataException>(() => new JsonDocumentStore().LoadAsync(path));
        }
        finally { File.Delete(path); }
    }
}

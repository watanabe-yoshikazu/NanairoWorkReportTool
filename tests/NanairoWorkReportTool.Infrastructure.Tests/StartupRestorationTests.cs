using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Core.Services;
using NanairoWorkReportTool.Infrastructure.Persistence;
using NanairoWorkReportTool.ViewModels;

namespace NanairoWorkReportTool.Infrastructure.Tests;

public sealed class StartupRestorationTests
{
    [Fact]
    public async Task Initialize_PrefersRecoveryOverLastNwr()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var documentStore = new JsonDocumentStore();
            var settingsStore = new SettingsStore(directory);
            var lastPath = Path.Combine(directory, "last.nwr");
            await documentStore.SaveAsync(lastPath, CreateDocument("保存済み"));
            await documentStore.SaveAsync(settingsStore.RecoveryPath, CreateDocument("自動復旧"));
            await settingsStore.SaveAsync(new AppSettings { LastNwrFilePath = lastPath });

            var viewModel = CreateViewModel(documentStore, settingsStore);
            await viewModel.InitializeAsync();

            Assert.Equal("自動復旧", viewModel.Document.ReporterName);
            Assert.Null(viewModel.CurrentFilePath);
            Assert.True(viewModel.IsDirty);
            Assert.Equal("自動復旧データ（未保存）", viewModel.DocumentStateText);
            Assert.True(viewModel.IsDocumentStateAttention);
        });
    }

    [Fact]
    public async Task Initialize_OpensLastNwrWhenThereIsNoRecovery()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var documentStore = new JsonDocumentStore();
            var settingsStore = new SettingsStore(directory);
            var lastPath = Path.Combine(directory, "作業報告_202608.nwr");
            await documentStore.SaveAsync(lastPath, CreateDocument("前回入力"));
            await settingsStore.SaveAsync(new AppSettings { LastNwrFilePath = lastPath });

            var viewModel = CreateViewModel(documentStore, settingsStore);
            await viewModel.InitializeAsync();

            Assert.Equal("前回入力", viewModel.Document.ReporterName);
            Assert.Equal(lastPath, viewModel.CurrentFilePath);
            Assert.False(viewModel.IsDirty);
            Assert.Equal("編集中: 作業報告_202608.nwr", viewModel.DocumentStateText);
            Assert.Equal(lastPath, viewModel.DocumentStatePath);
            Assert.False(viewModel.IsDocumentStateAttention);
        });
    }

    [Fact]
    public async Task Initialize_FallsBackFromBrokenRecoveryToLastNwr()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var documentStore = new JsonDocumentStore();
            var settingsStore = new SettingsStore(directory);
            var lastPath = Path.Combine(directory, "last.nwr");
            await documentStore.SaveAsync(lastPath, CreateDocument("前回入力"));
            Directory.CreateDirectory(Path.GetDirectoryName(settingsStore.RecoveryPath)!);
            await File.WriteAllTextAsync(settingsStore.RecoveryPath, "broken");
            await settingsStore.SaveAsync(new AppSettings { LastNwrFilePath = lastPath });

            var viewModel = CreateViewModel(documentStore, settingsStore);
            await viewModel.InitializeAsync();

            Assert.Equal("前回入力", viewModel.Document.ReporterName);
            Assert.Equal(lastPath, viewModel.CurrentFilePath);
        });
    }

    [Fact]
    public async Task Initialize_ShowsNewStateWhenLastNwrIsMissing()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var settingsStore = new SettingsStore(directory);
            var missingPath = Path.Combine(directory, "missing.nwr");
            await settingsStore.SaveAsync(new AppSettings { LastNwrFilePath = missingPath });

            var viewModel = CreateViewModel(new JsonDocumentStore(), settingsStore);
            await viewModel.InitializeAsync();

            Assert.Null(viewModel.CurrentFilePath);
            Assert.False(viewModel.IsDirty);
            Assert.Equal("新規・未保存（前回の .nwr はありません）", viewModel.DocumentStateText);
            Assert.True(viewModel.IsDocumentStateAttention);
        });
    }

    [Fact]
    public async Task Initialize_ShowsUnreadableStateWhenLastNwrIsBroken()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var settingsStore = new SettingsStore(directory);
            var brokenPath = Path.Combine(directory, "broken.nwr");
            await File.WriteAllTextAsync(brokenPath, "broken");
            await settingsStore.SaveAsync(new AppSettings { LastNwrFilePath = brokenPath });

            var viewModel = CreateViewModel(new JsonDocumentStore(), settingsStore);
            await viewModel.InitializeAsync();

            Assert.Null(viewModel.CurrentFilePath);
            Assert.Equal("新規・未保存（前回の .nwr を開けませんでした）", viewModel.DocumentStateText);
            Assert.Equal(brokenPath, viewModel.DocumentStatePath);
        });
    }

    [Fact]
    public async Task Initialize_IgnoresLastPathWithWrongExtension()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var documentStore = new JsonDocumentStore();
            var settingsStore = new SettingsStore(directory);
            var jsonPath = Path.Combine(directory, "previous.json");
            await documentStore.SaveAsync(jsonPath, CreateDocument("拡張子不正"));
            await settingsStore.SaveAsync(new AppSettings { LastNwrFilePath = jsonPath });

            var viewModel = CreateViewModel(documentStore, settingsStore);
            await viewModel.InitializeAsync();

            Assert.Null(viewModel.CurrentFilePath);
            Assert.Equal("新規・未保存（前回の .nwr はありません）", viewModel.DocumentStateText);
        });
    }

    [Fact]
    public async Task NewCommand_ChangesLoadedFileToNewState()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var documentStore = new JsonDocumentStore();
            var settingsStore = new SettingsStore(directory);
            var nwrPath = Path.Combine(directory, "previous.nwr");
            await documentStore.SaveAsync(nwrPath, CreateDocument("前回入力"));
            await settingsStore.SaveAsync(new AppSettings { LastNwrFilePath = nwrPath });
            var viewModel = CreateViewModel(documentStore, settingsStore);
            await viewModel.InitializeAsync();

            await viewModel.NewCommand.ExecuteAsync(null);

            Assert.Null(viewModel.CurrentFilePath);
            Assert.True(viewModel.IsDirty);
            Assert.Equal("新規・未保存", viewModel.DocumentStateText);
            Assert.Null(viewModel.DocumentStatePath);
        });
    }

    [Fact]
    public async Task OpenNwr_UpdatesLastNwrButOpeningExcelDoesNot()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var documentStore = new JsonDocumentStore();
            var settingsStore = new SettingsStore(directory);
            var nwrPath = Path.Combine(directory, "opened.nwr");
            await documentStore.SaveAsync(nwrPath, CreateDocument("NWR"));
            var viewModel = CreateViewModel(documentStore, settingsStore, CreateDocument("Excel"));
            await viewModel.InitializeAsync();

            await viewModel.OpenPathAsync(nwrPath);
            Assert.Equal(nwrPath, (await settingsStore.LoadAsync()).LastNwrFilePath);

            var excelPath = Path.Combine(directory, "imported.xlsx");
            await viewModel.OpenPathAsync(excelPath);

            Assert.Equal(nwrPath, (await settingsStore.LoadAsync()).LastNwrFilePath);
            Assert.Equal("Excelから読み込み（.nwr 未保存）", viewModel.DocumentStateText);
            Assert.Equal(excelPath, viewModel.DocumentStatePath);
        });
    }

    [Fact]
    public async Task SaveExistingNwr_PersistsLastPathAndSavedState()
    {
        await WithTestDirectoryAsync(async directory =>
        {
            var documentStore = new JsonDocumentStore();
            var settingsStore = new SettingsStore(directory);
            var nwrPath = Path.Combine(directory, "saved.nwr");
            await documentStore.SaveAsync(nwrPath, CreateDocument("保存前"));
            var viewModel = CreateViewModel(documentStore, settingsStore);
            await viewModel.InitializeAsync();
            await viewModel.OpenPathAsync(nwrPath);
            viewModel.Document.ReporterName = "保存後";

            await viewModel.SaveCommand.ExecuteAsync(null);

            Assert.Equal(nwrPath, (await settingsStore.LoadAsync()).LastNwrFilePath);
            Assert.Equal("編集中: saved.nwr", viewModel.DocumentStateText);
            Assert.False(viewModel.IsDirty);
            Assert.Equal("保存後", (await documentStore.LoadAsync(nwrPath)).ReporterName);
            await Task.Delay(1400);
            Assert.False(File.Exists(settingsStore.RecoveryPath));
        });
    }

    private static MainWindowViewModel CreateViewModel(
        IDocumentStore documentStore,
        SettingsStore settingsStore,
        WorkReportDocument? excelDocument = null)
    {
        var calculator = new WorkReportCalculator();
        return new MainWindowViewModel(
            calculator,
            new ReportValidator(calculator),
            documentStore,
            new FakeExcelReportService(excelDocument ?? CreateDocument("Excel")),
            new FakePdfExportService(),
            new FakeHolidayProvider(),
            settingsStore);
    }

    private static WorkReportDocument CreateDocument(string reporterName)
        => new()
        {
            TargetMonth = new DateTime(2026, 8, 1),
            ReporterName = reporterName,
            OutputDirectory = Path.GetTempPath()
        };

    private static async Task WithTestDirectoryAsync(Func<string, Task> action)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"NanairoStartup-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try { await action(directory); }
        finally { Directory.Delete(directory, true); }
    }

    private sealed class FakeHolidayProvider : IHolidayProvider
    {
        public Task<IReadOnlyDictionary<DateOnly, string>> GetHolidaysAsync(int year, CancellationToken cancellationToken = default)
            => Task.FromResult<IReadOnlyDictionary<DateOnly, string>>(new Dictionary<DateOnly, string>());

        public Task<bool> UpdateAsync(CancellationToken cancellationToken = default) => Task.FromResult(false);
    }

    private sealed class FakeExcelReportService(WorkReportDocument importedDocument) : IExcelReportService
    {
        public Task ExportAsync(string path, WorkReportDocument document, MonthlySummary summary, CancellationToken cancellationToken = default)
            => Task.CompletedTask;

        public Task<WorkReportDocument> ImportAsync(string path, CancellationToken cancellationToken = default)
            => Task.FromResult(importedDocument);

        public string BuildFileName(WorkReportDocument document) => "report.xlsx";
    }

    private sealed class FakePdfExportService : IPdfExportService
    {
        public Task ExportAsync(string excelPath, string pdfPath, CancellationToken cancellationToken = default)
            => Task.CompletedTask;
    }
}

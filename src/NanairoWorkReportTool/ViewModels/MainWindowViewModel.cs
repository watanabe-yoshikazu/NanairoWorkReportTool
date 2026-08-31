using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Core.Services;
using NanairoWorkReportTool.Infrastructure.Persistence;

namespace NanairoWorkReportTool.ViewModels;

public sealed record WorkStatusOption(WorkStatus Value, string DisplayName);

public partial class MainWindowViewModel : ObservableObject
{
    private readonly IWorkReportCalculator calculator;
    private readonly IReportValidator validator;
    private readonly IDocumentStore documentStore;
    private readonly IExcelReportService excelReportService;
    private readonly IPdfExportService pdfExportService;
    private readonly IHolidayProvider holidayProvider;
    private readonly SettingsStore settingsStore;
    private AppSettings settings = new();
    private bool suppressChanges;
    private CancellationTokenSource? recoveryCancellation;

    [ObservableProperty] private WorkReportDocument document = new();
    [ObservableProperty] private MonthlySummary summary = EmptySummary();
    [ObservableProperty] private bool isDirty;
    [ObservableProperty] private bool isBusy;
    [ObservableProperty] private string statusMessage = "準備中です…";
    [ObservableProperty] private string targetMonthText = DateTime.Today.ToString("yyyy/MM");
    [ObservableProperty] private string? currentFilePath;
    [ObservableProperty] private string documentStateText = "準備中です…";
    [ObservableProperty] private string? documentStatePath;
    [ObservableProperty] private bool isDocumentStateAttention = true;

    public ObservableCollection<ValidationIssue> ValidationIssues { get; } = [];
    public ObservableCollection<string> RecentFiles { get; } = [];
    public ObservableCollection<string> WorkContentHistory { get; } = [];
    public IReadOnlyList<WorkStatusOption> WorkStatusOptions { get; } =
    [
        new(WorkStatus.Unset, "未設定"),
        new(WorkStatus.Normal, "通常勤務"),
        new(WorkStatus.CompanyHoliday, "公休日"),
        new(WorkStatus.PaidLeave, "有給休暇")
    ];

    public int ErrorCount => ValidationIssues.Count(x => x.Severity == ValidationSeverity.Error);
    public int WarningCount => ValidationIssues.Count(x => x.Severity == ValidationSeverity.Warning);
    public string SettlementStatus => Summary.IsWithinSettlementRange ? "精算幅内" : "精算幅外";
    public string LeaveGuidance => Summary.AvailablePaidLeaveDays > 0
        ? $"下限まで {FormatHours(Summary.LowerBufferMinutes)} 時間の余裕があります。有給休暇 {Summary.AvailablePaidLeaveDays} 日分の目安です。"
        : $"下限までの余裕は {FormatHours(Summary.LowerBufferMinutes)} 時間です。取得可能な全日休暇の目安は0日です。";

    public MainWindowViewModel(
        IWorkReportCalculator calculator,
        IReportValidator validator,
        IDocumentStore documentStore,
        IExcelReportService excelReportService,
        IPdfExportService pdfExportService,
        IHolidayProvider holidayProvider,
        SettingsStore settingsStore)
    {
        this.calculator = calculator;
        this.validator = validator;
        this.documentStore = documentStore;
        this.excelReportService = excelReportService;
        this.pdfExportService = pdfExportService;
        this.holidayProvider = holidayProvider;
        this.settingsStore = settingsStore;
    }

    public async Task InitializeAsync()
    {
        IsBusy = true;
        try
        {
            settings = await settingsStore.LoadAsync();
            ReplaceCollection(RecentFiles, settings.RecentFiles.Where(File.Exists));
            ReplaceCollection(WorkContentHistory, settings.WorkContentHistory);
            if (File.Exists(settingsStore.RecoveryPath))
            {
                try
                {
                    var recovery = await documentStore.LoadAsync(settingsStore.RecoveryPath);
                    SetDocument(recovery, null, true);
                    SetDocumentState("自動復旧データ（未保存）", settingsStore.RecoveryPath, true);
                    StatusMessage = "前回の自動復旧データを開きました。";
                    _ = UpdateHolidayCacheInBackgroundAsync();
                    return;
                }
                catch { }
            }

            var lastNwrPath = settings.LastNwrFilePath;
            var hasValidLastNwrPath = IsNwrPath(lastNwrPath) && File.Exists(lastNwrPath);
            if (hasValidLastNwrPath)
            {
                try
                {
                    var previous = await documentStore.LoadAsync(lastNwrPath!);
                    SetDocument(previous, lastNwrPath, false);
                    AddRecent(lastNwrPath!);
                    SetDocumentState($"編集中: {Path.GetFileName(lastNwrPath)}", lastNwrPath, false);
                    StatusMessage = $"前回の {Path.GetFileName(lastNwrPath)} を開きました。";
                    _ = UpdateHolidayCacheInBackgroundAsync();
                    return;
                }
                catch { }
            }

            await CreateNewMonthAsync(DateTime.Today, false);
            var lastNwrCouldNotBeOpened = !string.IsNullOrWhiteSpace(lastNwrPath) && hasValidLastNwrPath;
            SetDocumentState(
                lastNwrCouldNotBeOpened
                    ? "新規・未保存（前回の .nwr を開けませんでした）"
                    : "新規・未保存（前回の .nwr はありません）",
                lastNwrPath,
                true);
            StatusMessage = lastNwrCouldNotBeOpened
                ? "前回の .nwr を読み込めないため、新規状態で開きました。"
                : "新規状態で入力を開始できます。";
            _ = UpdateHolidayCacheInBackgroundAsync();
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task NewAsync()
    {
        if (!await ConfirmChangeAsync()) return;
        var month = ParseTargetMonth() ?? DateTime.Today;
        await CreateNewMonthAsync(month, true);
        SetDocumentState("新規・未保存", null, true);
    }

    [RelayCommand]
    private async Task ApplyMonthAsync()
    {
        var month = ParseTargetMonth();
        if (month is null)
        {
            MessageBox.Show("入力月は yyyy/MM 形式で入力してください。", "入力月", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (!await ConfirmChangeAsync()) return;
        await CreateNewMonthAsync(month.Value, true);
        SetDocumentState("新規・未保存", null, true);
    }

    [RelayCommand]
    private void Check()
    {
        RefreshSummaryAndValidation();
        StatusMessage = ErrorCount > 0 ? $"エラーが {ErrorCount} 件あります。" : WarningCount > 0 ? $"警告が {WarningCount} 件あります。" : "問題は見つかりませんでした。";
    }

    [RelayCommand]
    private async Task SaveAsync()
    {
        var path = CurrentFilePath;
        if (string.IsNullOrWhiteSpace(path) || !string.Equals(Path.GetExtension(path), ".nwr", StringComparison.OrdinalIgnoreCase))
        {
            var dialog = new SaveFileDialog { Filter = "ナナイロ作業報告データ (*.nwr)|*.nwr", FileName = $"作業報告_{Document.TargetMonth:yyyyMM}.nwr" };
            if (dialog.ShowDialog() != true) return;
            path = dialog.FileName;
        }
        CancelRecoverySave();
        await documentStore.SaveAsync(path, Document);
        CurrentFilePath = path;
        IsDirty = false;
        AddRecent(path);
        CaptureHistory();
        settings.LastNwrFilePath = path;
        SetDocumentState($"編集中: {Path.GetFileName(path)}", path, false);
        await SaveSettingsAsync();
        if (File.Exists(settingsStore.RecoveryPath)) File.Delete(settingsStore.RecoveryPath);
        StatusMessage = "保存しました。";
    }

    [RelayCommand]
    private async Task OpenAsync()
    {
        if (!await ConfirmChangeAsync()) return;
        var dialog = new OpenFileDialog { Filter = "対応ファイル (*.nwr;*.xlsx)|*.nwr;*.xlsx|すべてのファイル (*.*)|*.*" };
        if (dialog.ShowDialog() == true) await OpenPathAsync(dialog.FileName);
    }

    [RelayCommand]
    private async Task OpenRecentAsync(string? path)
    {
        if (!string.IsNullOrWhiteSpace(path) && await ConfirmChangeAsync()) await OpenPathAsync(path);
    }

    [RelayCommand]
    private void BrowseOutputDirectory()
    {
        var dialog = new OpenFolderDialog { InitialDirectory = Directory.Exists(Document.OutputDirectory) ? Document.OutputDirectory : null };
        if (dialog.ShowDialog() == true) Document.OutputDirectory = dialog.FolderName;
    }

    [RelayCommand]
    private void OpenOutputDirectory()
    {
        if (Directory.Exists(Document.OutputDirectory))
            Process.Start(new ProcessStartInfo(Document.OutputDirectory!) { UseShellExecute = true });
    }

    [RelayCommand]
    private async Task ExportExcelAsync()
    {
        if (!CanExport()) return;
        var path = Path.Combine(Document.OutputDirectory!, excelReportService.BuildFileName(Document));
        if (!ConfirmOverwrite(path)) return;
        await RunBusyAsync(async () =>
        {
            await excelReportService.ExportAsync(path, Document, Summary);
            AddRecent(path);
            CaptureHistory();
            await SaveSettingsAsync();
            StatusMessage = $"Excelを出力しました: {Path.GetFileName(path)}";
            ShowOutputCompleted(path, null);
        });
    }

    [RelayCommand]
    private async Task ExportPdfAsync()
    {
        if (!CanExport()) return;
        var excelPath = Path.Combine(Document.OutputDirectory!, excelReportService.BuildFileName(Document));
        var pdfPath = Path.ChangeExtension(excelPath, ".pdf");
        if ((!ConfirmOverwrite(excelPath)) || (!ConfirmOverwrite(pdfPath))) return;
        await RunBusyAsync(async () =>
        {
            await excelReportService.ExportAsync(excelPath, Document, Summary);
            await pdfExportService.ExportAsync(excelPath, pdfPath);
            AddRecent(excelPath);
            CaptureHistory();
            await SaveSettingsAsync();
            StatusMessage = $"ExcelとPDFを出力しました: {Path.GetFileName(pdfPath)}";
            ShowOutputCompleted(excelPath, pdfPath);
        });
    }

    [RelayCommand]
    private async Task UpdateHolidaysAsync()
    {
        await RunBusyAsync(async () =>
        {
            var updated = await holidayProvider.UpdateAsync();
            StatusMessage = updated ? "祝日情報を更新しました。月を再生成すると反映されます。" : "祝日情報を更新できませんでした。保存済みデータを使用します。";
        });
    }

    public async Task OpenPathAsync(string path)
    {
        try
        {
            var isExcel = string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase);
            var isNwr = IsNwrPath(path);
            var loaded = isExcel
                ? await excelReportService.ImportAsync(path)
                : await documentStore.LoadAsync(path);
            SetDocument(loaded, path, false);
            AddRecent(path);
            if (isExcel)
            {
                SetDocumentState("Excelから読み込み（.nwr 未保存）", path, true);
            }
            else if (isNwr)
            {
                settings.LastNwrFilePath = path;
                SetDocumentState($"編集中: {Path.GetFileName(path)}", path, false);
                await SaveSettingsAsync();
            }
            else
            {
                SetDocumentState("読み込みデータ（.nwr 未保存）", path, true);
            }
            StatusMessage = $"{Path.GetFileName(path)} を開きました。";
        }
        catch (Exception exception)
        {
            MessageBox.Show(exception.Message, "ファイルを開けません", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public async Task TryOpenPathAsync(string path)
    {
        if (await ConfirmChangeAsync()) await OpenPathAsync(path);
    }

    public void SetStandardTime(IEnumerable<WorkDayEntry> entries)
    {
        foreach (var entry in entries) { entry.ApplyStatus(WorkStatus.Normal); entry.SetStandardTime(); }
    }

    public void ClearTime(IEnumerable<WorkDayEntry> entries)
    { foreach (var entry in entries) entry.ClearTime(); }

    public void ResetEntries(IEnumerable<WorkDayEntry> entries)
    {
        foreach (var entry in entries) entry.ResetToInitialState();
    }

    public void CopyPreviousContent(IEnumerable<WorkDayEntry> entries)
    {
        foreach (var entry in entries)
        {
            var previous = Document.Entries.FirstOrDefault(x => x.Date == entry.Date.AddDays(-1));
            if (previous is not null) entry.WorkContent = previous.WorkContent;
        }
    }

    public void ApplyWorkStatus(IEnumerable<WorkDayEntry> entries, WorkStatus status)
    {
        foreach (var entry in entries) entry.ApplyStatus(status);
    }

    public void ApplyWorkContent(IEnumerable<WorkDayEntry> entries, string? content)
    {
        if (string.IsNullOrWhiteSpace(content)) return;
        foreach (var entry in entries.Where(x => x.WorkStatus == WorkStatus.Normal)) entry.WorkContent = content.Trim();
        WorkContentHistory.Remove(content.Trim());
        WorkContentHistory.Insert(0, content.Trim());
        while (WorkContentHistory.Count > 20) WorkContentHistory.RemoveAt(WorkContentHistory.Count - 1);
    }

    public void ApplyCompanyHoliday(IEnumerable<WorkDayEntry> entries, string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return;
        foreach (var entry in entries)
        {
            entry.ApplyStatus(WorkStatus.CompanyHoliday);
            entry.CompanyHolidayName = name.Trim();
        }
    }

    public async Task<bool> TryCloseAsync()
    {
        if (IsDirty)
        {
            var result = MessageBox.Show("未保存の変更があります。終了前に保存しますか？", "終了確認", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Cancel) return false;
            if (result == MessageBoxResult.Yes)
            {
                await SaveAsync();
                if (IsDirty) return false;
            }
        }
        await SaveSettingsAsync();
        return true;
    }

    private async Task CreateNewMonthAsync(DateTime month, bool markDirty)
    {
        var holidays = await holidayProvider.GetHolidaysAsync(month.Year);
        var created = calculator.CreateMonth(month, holidays);
        created.ReporterName = settings.ReporterName;
        created.CompanyName = settings.CompanyName;
        created.Destination = settings.Destination;
        created.OutputDirectory = settings.OutputDirectory;
        SetDocument(created, null, markDirty);
    }

    private void SetDocument(WorkReportDocument value, string? path, bool dirty)
    {
        CancelRecoverySave();
        suppressChanges = true;
        Unsubscribe(Document);
        Document = value;
        Subscribe(Document);
        TargetMonthText = Document.TargetMonth.ToString("yyyy/MM");
        CurrentFilePath = path;
        IsDirty = dirty;
        suppressChanges = false;
        RefreshSummaryAndValidation();
    }

    private void SetDocumentState(string text, string? path, bool attention)
    {
        DocumentStateText = text;
        DocumentStatePath = path;
        IsDocumentStateAttention = attention;
    }

    private void Subscribe(WorkReportDocument value)
    {
        value.PropertyChanged += OnDocumentChanged;
        foreach (var entry in value.Entries) entry.PropertyChanged += OnEntryChanged;
    }
    private void Unsubscribe(WorkReportDocument value)
    {
        value.PropertyChanged -= OnDocumentChanged;
        foreach (var entry in value.Entries) entry.PropertyChanged -= OnEntryChanged;
    }
    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs e) => MarkChanged();
    private void OnEntryChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is WorkDayEntry entry && e.PropertyName == nameof(WorkDayEntry.WorkStatus))
            entry.ApplyStatus(entry.WorkStatus);
        MarkChanged();
    }

    private void MarkChanged()
    {
        if (suppressChanges) return;
        IsDirty = true;
        RefreshSummaryAndValidation();
        ScheduleRecoverySave();
    }

    private void RefreshSummaryAndValidation()
    {
        Summary = calculator.Calculate(Document, DateOnly.FromDateTime(DateTime.Today));
        ReplaceCollection(ValidationIssues, validator.Validate(Document, DateOnly.FromDateTime(DateTime.Today)));
        OnPropertyChanged(nameof(ErrorCount));
        OnPropertyChanged(nameof(WarningCount));
        OnPropertyChanged(nameof(SettlementStatus));
        OnPropertyChanged(nameof(LeaveGuidance));
    }

    private bool CanExport()
    {
        Check();
        if (ErrorCount > 0)
        {
            MessageBox.Show("エラーを修正してから出力してください。", "出力できません", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        return WarningCount == 0 || MessageBox.Show($"警告が {WarningCount} 件あります。出力を続けますか？", "警告", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes;
    }

    private void ScheduleRecoverySave()
    {
        CancelRecoverySave();
        recoveryCancellation = new CancellationTokenSource();
        var token = recoveryCancellation.Token;
        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(1200, token);
                await documentStore.SaveAsync(settingsStore.RecoveryPath, Document, token);
            }
            catch (OperationCanceledException) { }
            catch { }
        }, token);
    }

    private void CancelRecoverySave()
    {
        recoveryCancellation?.Cancel();
        recoveryCancellation = null;
    }

    private async Task RunBusyAsync(Func<Task> action)
    {
        IsBusy = true;
        try { await action(); }
        catch (Exception exception) { MessageBox.Show(exception.Message, "処理に失敗しました", MessageBoxButton.OK, MessageBoxImage.Error); }
        finally { IsBusy = false; }
    }

    private async Task SaveSettingsAsync()
    {
        settings.ReporterName = Document.ReporterName ?? string.Empty;
        settings.CompanyName = Document.CompanyName ?? string.Empty;
        settings.Destination = Document.Destination ?? string.Empty;
        settings.OutputDirectory = Document.OutputDirectory ?? string.Empty;
        settings.RecentFiles = RecentFiles.Take(10).ToList();
        settings.WorkContentHistory = WorkContentHistory.Take(20).ToList();
        await settingsStore.SaveAsync(settings);
    }

    private void CaptureHistory()
    {
        foreach (var text in Document.Entries.Select(x => x.WorkContent).Where(x => !string.IsNullOrWhiteSpace(x)).Reverse())
        {
            WorkContentHistory.Remove(text!);
            WorkContentHistory.Insert(0, text!);
        }
        while (WorkContentHistory.Count > 20) WorkContentHistory.RemoveAt(WorkContentHistory.Count - 1);
    }

    private void AddRecent(string path)
    {
        RecentFiles.Remove(path);
        RecentFiles.Insert(0, path);
        while (RecentFiles.Count > 10) RecentFiles.RemoveAt(RecentFiles.Count - 1);
    }

    private async Task<bool> ConfirmChangeAsync()
    {
        if (!IsDirty) return true;
        var result = MessageBox.Show("未保存の変更があります。続行前に保存しますか？", "保存確認", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
        if (result == MessageBoxResult.Cancel) return false;
        if (result == MessageBoxResult.No) return true;
        await SaveAsync();
        return !IsDirty;
    }
    private static bool ConfirmOverwrite(string path) => !File.Exists(path) || MessageBox.Show($"{Path.GetFileName(path)} は既に存在します。上書きしますか？", "上書き確認", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
    private static bool IsNwrPath(string? path)
        => !string.IsNullOrWhiteSpace(path)
           && string.Equals(Path.GetExtension(path), ".nwr", StringComparison.OrdinalIgnoreCase);
    private DateTime? ParseTargetMonth() => DateTime.TryParseExact(TargetMonthText.Trim(), ["yyyy/MM", "yyyy/M"], null, System.Globalization.DateTimeStyles.None, out var month) ? new DateTime(month.Year, month.Month, 1) : null;

    private static void ShowOutputCompleted(string excelPath, string? pdfPath)
        => new OutputCompletedWindow(excelPath, pdfPath) { Owner = Application.Current.MainWindow }.ShowDialog();

    private async Task UpdateHolidayCacheInBackgroundAsync()
    {
        if (await holidayProvider.UpdateAsync()) StatusMessage = "祝日情報を確認しました。";
    }

    private static string FormatHours(int minutes) => (minutes / 60m).ToString("0.##");
    private static MonthlySummary EmptySummary() => new(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, false);
    private static void ReplaceCollection<T>(ObservableCollection<T> target, IEnumerable<T> values)
    { target.Clear(); foreach (var value in values) target.Add(value); }
}

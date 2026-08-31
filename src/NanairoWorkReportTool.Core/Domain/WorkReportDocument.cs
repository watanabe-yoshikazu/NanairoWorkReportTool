using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace NanairoWorkReportTool.Core.Domain;

public partial class WorkReportDocument : ObservableObject
{
    public const int CurrentSchemaVersion = 1;

    [ObservableProperty] private int schemaVersion = CurrentSchemaVersion;
    [ObservableProperty] private DateTime targetMonth = new(DateTime.Today.Year, DateTime.Today.Month, 1);
    [ObservableProperty] private string? reporterName;
    [ObservableProperty] private string? companyName;
    [ObservableProperty] private string? destination;
    [ObservableProperty] private string? outputDirectory;

    public ObservableCollection<WorkDayEntry> Entries { get; set; } = [];
}

public sealed record HolidayInfo(DateOnly Date, string Name);

public sealed record MonthlySummary(
    int WeekdayWorkDays,
    int BaselineCandidateDays,
    int BaselineCompanyHolidayDays,
    int CompanyHolidayDays,
    int PaidLeaveDays,
    int BaselineMinutes,
    int LowerMinutes,
    int UpperMinutes,
    int ActualToDateMinutes,
    int ForecastMinutes,
    int DifferenceMinutes,
    int LowerBufferMinutes,
    int UpperRemainingMinutes,
    int AvailablePaidLeaveDays,
    bool IsWithinSettlementRange)
{
    public decimal BaselineHours => BaselineMinutes / 60m;
    public decimal LowerHours => LowerMinutes / 60m;
    public decimal UpperHours => UpperMinutes / 60m;
    public decimal ActualToDateHours => ActualToDateMinutes / 60m;
    public decimal ForecastHours => ForecastMinutes / 60m;
    public decimal DifferenceHours => DifferenceMinutes / 60m;
    public decimal LowerBufferHours => LowerBufferMinutes / 60m;
    public decimal UpperRemainingHours => UpperRemainingMinutes / 60m;
}

public sealed record ValidationIssue(
    ValidationSeverity Severity,
    DateOnly? Date,
    string Field,
    string Message)
{
    public string SeverityDisplay => Severity == ValidationSeverity.Error ? "エラー" : "警告";
    public string DateDisplay => Date.HasValue ? $"{Date.Value.Month}月{Date.Value.Day}日" : "月間工数";
}


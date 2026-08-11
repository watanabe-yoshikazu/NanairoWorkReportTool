using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Core.Services;

namespace NanairoWorkReportTool.Core.Tests;

public sealed class ReportValidatorTests
{
    private readonly WorkReportCalculator _calculator = new();

    [Fact]
    public void Validate_MissingPastContentIsErrorButFutureContentIsWarning()
    {
        var document = CompleteDocument(new DateTime(2026, 7, 1));
        var past = new WorkDayEntry { Date = new DateOnly(2026, 7, 1), DayType = DayType.Weekday, WorkStatus = WorkStatus.Normal };
        past.SetStandardTime();
        var future = new WorkDayEntry { Date = new DateOnly(2026, 7, 3), DayType = DayType.Weekday, WorkStatus = WorkStatus.Normal };
        future.SetStandardTime();
        document.Entries.Add(past);
        document.Entries.Add(future);
        var issues = new ReportValidator(_calculator).Validate(document, new DateOnly(2026, 7, 2));
        Assert.Contains(issues, x => x.Date == past.Date && x.Field == "作業内容" && x.Severity == ValidationSeverity.Error);
        Assert.Contains(issues, x => x.Date == future.Date && x.Field == "作業内容" && x.Severity == ValidationSeverity.Warning);
    }

    [Fact]
    public void Validate_FutureContradictoryTimesRemainError()
    {
        var document = CompleteDocument(new DateTime(2026, 7, 1));
        document.Entries.Add(new WorkDayEntry
        {
            Date = new DateOnly(2026, 7, 10), DayType = DayType.Weekday, WorkStatus = WorkStatus.Normal,
            WorkContent = "テスト", StartMinutes = 600, EndMinutes = 540, BreakMinutes = 60
        });
        var issues = new ReportValidator(_calculator).Validate(document, new DateOnly(2026, 7, 2));
        Assert.Contains(issues, x => x.Date == new DateOnly(2026, 7, 10) && x.Field == "終了時刻" && x.Severity == ValidationSeverity.Error);
    }

    [Fact]
    public void Validate_WeekendWorkAndPaidLeaveOnHolidayAreWarnings()
    {
        var document = CompleteDocument(new DateTime(2026, 7, 1));
        document.Entries.Add(new WorkDayEntry
        {
            Date = new DateOnly(2026, 7, 4), DayType = DayType.Saturday, WorkStatus = WorkStatus.Normal,
            WorkContent = "休日対応", StartMinutes = 540, EndMinutes = 1020, BreakMinutes = 60
        });
        document.Entries.Add(new WorkDayEntry { Date = new DateOnly(2026, 7, 5), DayType = DayType.Sunday, WorkStatus = WorkStatus.PaidLeave });
        var issues = new ReportValidator(_calculator).Validate(document, new DateOnly(2026, 7, 31));
        Assert.Contains(issues, x => x.Date == new DateOnly(2026, 7, 4) && x.Severity == ValidationSeverity.Warning);
        Assert.Contains(issues, x => x.Date == new DateOnly(2026, 7, 5) && x.Severity == ValidationSeverity.Warning);
    }

    [Fact]
    public void Validate_PaidLeaveDoesNotRequireOrValidateWorkDetails()
    {
        var document = CompleteDocument(new DateTime(2026, 7, 1));
        var paidLeave = new WorkDayEntry
        {
            Date = new DateOnly(2026, 7, 6), DayType = DayType.Weekday, WorkStatus = WorkStatus.PaidLeave,
            WorkContent = null, StartMinutes = 600, EndMinutes = 540, BreakMinutes = 600
        };
        document.Entries.Add(paidLeave);

        var issues = new ReportValidator(_calculator).Validate(document, new DateOnly(2026, 7, 31));

        Assert.DoesNotContain(issues, x => x.Date == paidLeave.Date && x.Severity == ValidationSeverity.Error);
        Assert.DoesNotContain(issues, x => x.Date == paidLeave.Date && x.Field is "作業内容" or "開始時刻" or "終了時刻" or "休憩時間");
    }

    private static WorkReportDocument CompleteDocument(DateTime month) => new()
    {
        TargetMonth = month, ReporterName = "渡辺", CompanyName = "株式会社リンク",
        Destination = "株式会社ナナイロ 御中", OutputDirectory = Path.GetTempPath()
    };
}

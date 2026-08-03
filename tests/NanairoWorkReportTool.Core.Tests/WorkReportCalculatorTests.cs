using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Core.Services;

namespace NanairoWorkReportTool.Core.Tests;

public sealed class WorkReportCalculatorTests
{
    private readonly WorkReportCalculator _calculator = new();

    [Theory]
    [InlineData(2025, 2, 28)]
    [InlineData(2024, 2, 29)]
    [InlineData(2026, 4, 30)]
    [InlineData(2026, 7, 31)]
    public void CreateMonth_GeneratesEveryCalendarDay(int year, int month, int expected)
    {
        var document = _calculator.CreateMonth(new DateTime(year, month, 1), new Dictionary<DateOnly, string>());
        Assert.Equal(expected, document.Entries.Count);
        Assert.Equal(new DateOnly(year, month, expected), document.Entries[^1].Date);
    }

    [Fact]
    public void CreateMonth_InitializesWeekdaysAndOfficialHolidayCorrectly()
    {
        var holidays = new Dictionary<DateOnly, string> { [new(2026, 7, 20)] = "海の日" };
        var document = _calculator.CreateMonth(new DateTime(2026, 7, 1), holidays);
        var holiday = Assert.Single(document.Entries, x => x.Date == new DateOnly(2026, 7, 20));
        Assert.Equal(DayType.Holiday, holiday.DayType);
        Assert.Equal(WorkStatus.Unset, holiday.WorkStatus);
        Assert.Equal("海の日", holiday.Note);
        Assert.All(document.Entries.Where(x => x.DayType == DayType.Weekday), x => Assert.Equal(450, x.WorkMinutes));
    }

    [Fact]
    public void Calculate_CompanyHolidayReducesBaselineButPaidLeaveDoesNot()
    {
        var document = _calculator.CreateMonth(new DateTime(2026, 7, 1), new Dictionary<DateOnly, string>());
        var initial = _calculator.Calculate(document, new DateOnly(2026, 7, 31));
        var weekdays = document.Entries.Where(x => x.DayType == DayType.Weekday).Take(2).ToArray();
        weekdays[0].ApplyStatus(WorkStatus.CompanyHoliday);
        weekdays[0].CompanyHolidayName = "夏季休業";
        weekdays[1].ApplyStatus(WorkStatus.PaidLeave);
        var result = _calculator.Calculate(document, new DateOnly(2026, 7, 31));
        Assert.Equal(initial.BaselineMinutes - 450, result.BaselineMinutes);
        Assert.Equal(1, result.CompanyHolidayDays);
        Assert.Equal(1, result.PaidLeaveDays);
        Assert.Equal(initial.ForecastMinutes - 900, result.ForecastMinutes);
    }

    [Fact]
    public void Calculate_UsesMinutePrecisionAndFloorsAvailableLeaveDays()
    {
        var document = new WorkReportDocument { TargetMonth = new DateTime(2026, 7, 1) };
        for (var day = 1; day <= 5; day++)
        {
            document.Entries.Add(new WorkDayEntry
            {
                Date = new DateOnly(2026, 7, day + 5), DayType = DayType.Weekday, WorkStatus = WorkStatus.Normal,
                StartMinutes = 0, EndMinutes = day == 5 ? 449 : 450, BreakMinutes = 0
            });
        }
        var result = _calculator.Calculate(document, new DateOnly(2026, 7, 31));
        Assert.Equal(2249, result.ForecastMinutes);
        Assert.Equal(1199, result.LowerBufferMinutes);
        Assert.Equal(2, result.AvailablePaidLeaveDays);
    }
}

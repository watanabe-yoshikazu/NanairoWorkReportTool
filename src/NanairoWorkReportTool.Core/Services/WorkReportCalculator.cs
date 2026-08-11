using NanairoWorkReportTool.Core.Domain;

namespace NanairoWorkReportTool.Core.Services;

public sealed class WorkReportCalculator : IWorkReportCalculator
{
    public WorkReportDocument CreateMonth(DateTime month, IReadOnlyDictionary<DateOnly, string> holidays)
    {
        var first = new DateTime(month.Year, month.Month, 1);
        var document = new WorkReportDocument { TargetMonth = first };
        for (var day = 1; day <= DateTime.DaysInMonth(first.Year, first.Month); day++)
        {
            var date = new DateOnly(first.Year, first.Month, day);
            holidays.TryGetValue(date, out var holidayName);
            var dayType = holidayName is not null ? DayType.Holiday : date.DayOfWeek switch
            {
                DayOfWeek.Saturday => DayType.Saturday,
                DayOfWeek.Sunday => DayType.Sunday,
                _ => DayType.Weekday
            };
            var entry = new WorkDayEntry
            {
                Date = date,
                DayType = dayType,
                HolidayName = holidayName,
                Note = holidayName,
                WorkStatus = dayType == DayType.Weekday ? WorkStatus.Normal : WorkStatus.Unset
            };
            if (entry.WorkStatus == WorkStatus.Normal) entry.SetStandardTime();
            document.Entries.Add(entry);
        }
        return document;
    }

    public MonthlySummary Calculate(WorkReportDocument document, DateOnly today)
    {
        var weekdayWorkDays = document.Entries.Count(IsBaselineWorkday);
        var baseline = weekdayWorkDays * 450;
        var lower = baseline - 1_200;
        var upper = baseline + 1_200;
        var actual = document.Entries.Where(x => x.Date <= today && x.WorkStatus == WorkStatus.Normal).Sum(x => x.WorkMinutes);
        var forecast = document.Entries.Where(x => x.WorkStatus == WorkStatus.Normal).Sum(x => x.WorkMinutes);
        var buffer = forecast - lower;
        return new MonthlySummary(
            weekdayWorkDays,
            document.Entries.Count(x => x.WorkStatus == WorkStatus.CompanyHoliday),
            document.Entries.Count(x => x.WorkStatus == WorkStatus.PaidLeave),
            baseline, lower, upper, actual, forecast, forecast - baseline, buffer, upper - forecast,
            Math.Max(0, buffer / 450), forecast >= lower && forecast <= upper);
    }

    private static bool IsBaselineWorkday(WorkDayEntry entry)
        => entry.Date.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday
           && entry.DayType != DayType.Holiday
           && entry.WorkStatus != WorkStatus.CompanyHoliday;
}


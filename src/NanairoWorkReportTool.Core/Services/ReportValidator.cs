using NanairoWorkReportTool.Core.Domain;

namespace NanairoWorkReportTool.Core.Services;

public sealed class ReportValidator(IWorkReportCalculator calculator) : IReportValidator
{
    public IReadOnlyList<ValidationIssue> Validate(WorkReportDocument document, DateOnly today)
    {
        var issues = new List<ValidationIssue>();
        Required(document.TargetMonth == default, "入力月", "入力月が指定されていません。", issues);
        Required(string.IsNullOrWhiteSpace(document.ReporterName), "入力者名", "入力者名が入力されていません。", issues);
        Required(string.IsNullOrWhiteSpace(document.CompanyName), "会社名", "会社名が入力されていません。", issues);
        Required(string.IsNullOrWhiteSpace(document.Destination), "提出先", "提出先が入力されていません。", issues);
        Required(string.IsNullOrWhiteSpace(document.OutputDirectory), "保存先", "保存先が指定されていません。", issues);
        foreach (var entry in document.Entries) ValidateEntry(entry, today, issues);

        var summary = calculator.Calculate(document, today);
        if (summary.ForecastMinutes < summary.LowerMinutes)
            issues.Add(Warning(null, "工数合計", "月末までの見込み工数が下限工数を下回っています。"));
        else if (summary.ForecastMinutes > summary.UpperMinutes)
            issues.Add(Warning(null, "工数合計", "月末までの見込み工数が上限工数を上回っています。"));
        if (document.TargetMonth.Year == today.Year && document.TargetMonth.Month == today.Month
            && document.Entries.Any(x => x.Date > today && x.WorkStatus == WorkStatus.Normal))
            issues.Add(Warning(null, "見込み工数", "未来日の標準勤務を含む見込み工数です。当日までの実績工数と区別してください。"));
        return issues;
    }

    private static void ValidateEntry(WorkDayEntry entry, DateOnly today, ICollection<ValidationIssue> issues)
    {
        var future = entry.Date > today;
        if (entry.WorkStatus == WorkStatus.Normal)
        {
            EntryRequired(string.IsNullOrWhiteSpace(entry.WorkContent), entry, future, "作業内容", "作業内容が入力されていません。", issues);
            EntryRequired(!entry.StartMinutes.HasValue, entry, future, "開始時刻", "開始時刻が入力されていません。", issues);
            EntryRequired(!entry.EndMinutes.HasValue, entry, future, "終了時刻", "終了時刻が入力されていません。", issues);
            EntryRequired(!entry.BreakMinutes.HasValue, entry, future, "休憩時間", "休憩時間が入力されていません。", issues);
            if (entry.StartMinutes.HasValue ^ entry.EndMinutes.HasValue)
                issues.Add(Error(entry.Date, "勤務時刻", "開始時刻または終了時刻の一方だけが入力されています。"));
            if (entry.StartMinutes.HasValue && entry.EndMinutes.HasValue && entry.EndMinutes <= entry.StartMinutes)
                issues.Add(Error(entry.Date, "終了時刻", "終了時刻は開始時刻より後にしてください。"));
            if (entry.BreakMinutes.HasValue && entry.StartMinutes.HasValue && entry.EndMinutes.HasValue
                && entry.BreakMinutes >= entry.EndMinutes - entry.StartMinutes)
                issues.Add(Error(entry.Date, "休憩時間", "休憩時間は総稼働時間より短くしてください。"));
            if (entry.StartMinutes.HasValue && entry.EndMinutes.HasValue && entry.BreakMinutes.HasValue
                && entry.EndMinutes - entry.StartMinutes - entry.BreakMinutes <= 0)
                issues.Add(Error(entry.Date, "工数", "算出工数が0時間以下です。"));
            if (entry.DayType is DayType.Saturday or DayType.Sunday or DayType.Holiday)
                issues.Add(Warning(entry.Date, "勤務区分", $"{entry.DayTypeDisplay}に通常勤務が入力されています。"));
            if (entry.DayType == DayType.Weekday && entry.WorkMinutes > 0 && entry.WorkMinutes != 450)
                issues.Add(Warning(entry.Date, "工数", "平日の工数が標準勤務時間の7.5時間と異なります。"));
        }
        else
        {
            if (entry.WorkStatus != WorkStatus.PaidLeave
                && (entry.StartMinutes.HasValue || entry.EndMinutes.HasValue || entry.BreakMinutes.HasValue))
                issues.Add(Error(entry.Date, "勤務時刻", "公休日・未設定の日には勤務時間を設定できません。"));
            if (entry.WorkStatus == WorkStatus.CompanyHoliday && string.IsNullOrWhiteSpace(entry.CompanyHolidayName))
                issues.Add(Error(entry.Date, "公休日名", "公休日名が入力されていません。"));
            if (entry.WorkStatus == WorkStatus.Unset && entry.DayType == DayType.Weekday)
                issues.Add(Warning(entry.Date, "勤務区分", "平日の勤務区分が未設定です。"));
            if (entry.WorkStatus == WorkStatus.PaidLeave && entry.DayType is (DayType.Saturday or DayType.Sunday or DayType.Holiday))
                issues.Add(Warning(entry.Date, "勤務区分", "土日または祝日に有給休暇が設定されています。"));
        }
    }

    private static void Required(bool condition, string field, string message, ICollection<ValidationIssue> issues)
    { if (condition) issues.Add(Error(null, field, message)); }
    private static void EntryRequired(bool condition, WorkDayEntry entry, bool future, string field, string message, ICollection<ValidationIssue> issues)
    { if (condition) issues.Add(future ? Warning(entry.Date, field, $"未来日: {message}") : Error(entry.Date, field, message)); }
    private static ValidationIssue Error(DateOnly? date, string field, string message) => new(ValidationSeverity.Error, date, field, message);
    private static ValidationIssue Warning(DateOnly? date, string field, string message) => new(ValidationSeverity.Warning, date, field, message);
}

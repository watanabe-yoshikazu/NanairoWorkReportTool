using System.Text.Json.Serialization;
using CommunityToolkit.Mvvm.ComponentModel;

namespace NanairoWorkReportTool.Core.Domain;

public partial class WorkDayEntry : ObservableObject
{
    [ObservableProperty] private DateOnly date;
    [ObservableProperty] private DayType dayType;
    [ObservableProperty] private string? holidayName;
    [ObservableProperty] private WorkStatus workStatus;
    [ObservableProperty] private string? workContent;
    [ObservableProperty] private int? startMinutes;
    [ObservableProperty] private int? endMinutes;
    [ObservableProperty] private int? breakMinutes;
    [ObservableProperty] private string? companyHolidayName;
    [ObservableProperty] private string? note;

    [JsonIgnore]
    public int GrossMinutes => WorkStatus == WorkStatus.Normal && StartMinutes.HasValue && EndMinutes.HasValue
        ? Math.Max(0, EndMinutes.Value - StartMinutes.Value)
        : 0;

    [JsonIgnore]
    public int WorkMinutes => WorkStatus == WorkStatus.Normal && BreakMinutes.HasValue
        ? Math.Max(0, GrossMinutes - BreakMinutes.Value)
        : 0;

    [JsonIgnore]
    public string DateDisplay => $"{Date.Month}/{Date.Day}";

    [JsonIgnore]
    public string DayOfWeekDisplay => Date.DayOfWeek switch
    {
        DayOfWeek.Sunday => "日",
        DayOfWeek.Monday => "月",
        DayOfWeek.Tuesday => "火",
        DayOfWeek.Wednesday => "水",
        DayOfWeek.Thursday => "木",
        DayOfWeek.Friday => "金",
        DayOfWeek.Saturday => "土",
        _ => string.Empty
    };

    [JsonIgnore]
    public string DayTypeDisplay => DayType switch
    {
        DayType.Weekday => "平日",
        DayType.Saturday => "土曜日",
        DayType.Sunday => "日曜日",
        DayType.Holiday => $"祝日{(string.IsNullOrWhiteSpace(HolidayName) ? string.Empty : $"（{HolidayName}）")}",
        _ => string.Empty
    };

    [JsonIgnore]
    public string WorkStatusDisplay => WorkStatus switch
    {
        WorkStatus.Unset => "未設定",
        WorkStatus.Normal => "通常勤務",
        WorkStatus.CompanyHoliday => "公休日",
        WorkStatus.PaidLeave => "有給休暇",
        _ => string.Empty
    };

    [JsonIgnore]
    public string WorkStatusIcon => WorkStatus switch
    {
        WorkStatus.Normal => "●",
        WorkStatus.CompanyHoliday => "休",
        WorkStatus.PaidLeave => "有",
        _ => "－"
    };

    [JsonIgnore]
    public string GrossHoursDisplay => FormatHours(GrossMinutes);

    [JsonIgnore]
    public string WorkHoursDisplay => FormatHours(WorkMinutes);

    [JsonIgnore]
    public string StartTimeText
    {
        get => FormatTime(StartMinutes);
        set => StartMinutes = ParseTime(value);
    }

    [JsonIgnore]
    public string EndTimeText
    {
        get => FormatTime(EndMinutes);
        set => EndMinutes = ParseTime(value);
    }

    [JsonIgnore]
    public string BreakHoursText
    {
        get => BreakMinutes.HasValue ? FormatHours(BreakMinutes.Value) : string.Empty;
        set => BreakMinutes = ParseHours(value);
    }

    [JsonIgnore]
    public bool IsNormal => WorkStatus == WorkStatus.Normal;

    [JsonIgnore]
    public bool IsCompanyHoliday => WorkStatus == WorkStatus.CompanyHoliday;

    public string GetReportRemark()
    {
        if (WorkStatus == WorkStatus.PaidLeave)
        {
            return "有給休暇";
        }

        if (WorkStatus == WorkStatus.CompanyHoliday)
        {
            return CompanyHolidayName?.Trim() ?? string.Empty;
        }

        return string.Join(" / ", new[] { HolidayName, Note }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct());
    }

    public void ApplyStatus(WorkStatus status, bool restoreStandardTime = true)
    {
        WorkStatus = status;
        if (status == WorkStatus.Normal)
        {
            if (restoreStandardTime)
            {
                SetStandardTime();
            }
        }
        else
        {
            StartMinutes = null;
            EndMinutes = null;
            BreakMinutes = null;
            WorkContent = null;
        }

        if (status != WorkStatus.CompanyHoliday)
        {
            CompanyHolidayName = null;
        }

        OnPropertyChanged(nameof(IsNormal));
        OnPropertyChanged(nameof(IsCompanyHoliday));
    }

    public void SetStandardTime()
    {
        StartMinutes = 9 * 60;
        EndMinutes = 17 * 60 + 30;
        BreakMinutes = 60;
    }

    public void ClearTime()
    {
        StartMinutes = null;
        EndMinutes = null;
        BreakMinutes = null;
    }

    public void ResetToInitialState()
    {
        WorkContent = null;
        CompanyHolidayName = null;
        Note = HolidayName;
        ApplyStatus(DayType == DayType.Weekday ? WorkStatus.Normal : WorkStatus.Unset);
    }

    partial void OnWorkStatusChanged(WorkStatus value)
    {
        OnPropertyChanged(nameof(WorkStatusDisplay));
        OnPropertyChanged(nameof(WorkStatusIcon));
        OnPropertyChanged(nameof(IsNormal));
        OnPropertyChanged(nameof(IsCompanyHoliday));
        RaiseCalculatedProperties();
    }

    partial void OnStartMinutesChanged(int? value)
    {
        OnPropertyChanged(nameof(StartTimeText));
        RaiseCalculatedProperties();
    }

    partial void OnEndMinutesChanged(int? value)
    {
        OnPropertyChanged(nameof(EndTimeText));
        RaiseCalculatedProperties();
    }

    partial void OnBreakMinutesChanged(int? value)
    {
        OnPropertyChanged(nameof(BreakHoursText));
        RaiseCalculatedProperties();
    }

    private void RaiseCalculatedProperties()
    {
        OnPropertyChanged(nameof(GrossMinutes));
        OnPropertyChanged(nameof(WorkMinutes));
        OnPropertyChanged(nameof(GrossHoursDisplay));
        OnPropertyChanged(nameof(WorkHoursDisplay));
    }

    private static string FormatTime(int? minutes) => minutes.HasValue
        ? $"{minutes.Value / 60:00}:{minutes.Value % 60:00}"
        : string.Empty;

    private static int? ParseTime(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (TimeOnly.TryParse(value.Trim(), out var time)) return time.Hour * 60 + time.Minute;
        return null;
    }

    private static string FormatHours(int minutes)
        => (minutes / 60m).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);

    private static int? ParseHours(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (!decimal.TryParse(value.Trim(), System.Globalization.NumberStyles.Number,
                System.Globalization.CultureInfo.InvariantCulture, out var hours)) return null;
        var minutes = hours * 60m;
        return minutes == decimal.Truncate(minutes) && minutes >= 0 && minutes <= 24 * 60
            ? (int)minutes
            : null;
    }
}

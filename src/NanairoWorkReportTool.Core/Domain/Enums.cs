namespace NanairoWorkReportTool.Core.Domain;

public enum DayType
{
    Weekday,
    Saturday,
    Sunday,
    Holiday
}

public enum WorkStatus
{
    Unset,
    Normal,
    CompanyHoliday,
    PaidLeave
}

public enum ValidationSeverity
{
    Error,
    Warning
}


using NanairoWorkReportTool.Core.Domain;

namespace NanairoWorkReportTool.Core.Services;

public interface IHolidayProvider
{
    Task<IReadOnlyDictionary<DateOnly, string>> GetHolidaysAsync(int year, CancellationToken cancellationToken = default);
    Task<bool> UpdateAsync(CancellationToken cancellationToken = default);
}

public interface IWorkReportCalculator
{
    WorkReportDocument CreateMonth(DateTime month, IReadOnlyDictionary<DateOnly, string> holidays);
    MonthlySummary Calculate(WorkReportDocument document, DateOnly today);
}

public interface IReportValidator
{
    IReadOnlyList<ValidationIssue> Validate(WorkReportDocument document, DateOnly today);
}

public interface IDocumentStore
{
    Task SaveAsync(string path, WorkReportDocument document, CancellationToken cancellationToken = default);
    Task<WorkReportDocument> LoadAsync(string path, CancellationToken cancellationToken = default);
}

public interface IExcelReportService
{
    Task ExportAsync(string path, WorkReportDocument document, MonthlySummary summary, CancellationToken cancellationToken = default);
    Task<WorkReportDocument> ImportAsync(string path, CancellationToken cancellationToken = default);
    string BuildFileName(WorkReportDocument document);
}

public interface IPdfExportService
{
    Task ExportAsync(string excelPath, string pdfPath, CancellationToken cancellationToken = default);
}

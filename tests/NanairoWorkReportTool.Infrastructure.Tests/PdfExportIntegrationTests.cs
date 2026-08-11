using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Core.Services;
using NanairoWorkReportTool.Infrastructure.Excel;

namespace NanairoWorkReportTool.Infrastructure.Tests;

public sealed class PdfExportIntegrationTests
{
    [Fact]
    [Trait("Category", "ExcelIntegration")]
    public async Task ExportPdf_CreatesReportsFor28Through31DayMonths()
    {
        if (Type.GetTypeFromProgID("Excel.Application") is null) return;
        var outputDirectory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../artifacts/pdf-validation"));
        Directory.CreateDirectory(outputDirectory);
        var calculator = new WorkReportCalculator();
        var excelService = new OpenXmlExcelReportService();
        var pdfService = new ExcelPdfExportService();

        foreach (var (year, month) in new[] { (2025, 2), (2024, 2), (2026, 4), (2026, 7) })
        {
            var document = calculator.CreateMonth(new DateTime(year, month, 1), new Dictionary<DateOnly, string>());
            document.ReporterName = "PDF検証";
            document.CompanyName = "株式会社リンク";
            document.Destination = "株式会社ナナイロ 御中";
            document.OutputDirectory = outputDirectory;
            foreach (var entry in document.Entries.Where(x => x.WorkStatus == WorkStatus.Normal)) entry.WorkContent = "帳票レイアウト検証";
            var excelPath = Path.Combine(outputDirectory, $"validation-{year}{month:00}.xlsx");
            var pdfPath = Path.ChangeExtension(excelPath, ".pdf");
            await excelService.ExportAsync(excelPath, document, calculator.Calculate(document, new DateOnly(year, month, DateTime.DaysInMonth(year, month))));
            await pdfService.ExportAsync(excelPath, pdfPath);
            Assert.True(new FileInfo(pdfPath).Length > 1_000, $"PDFが正しく生成されていません: {pdfPath}");
        }
    }
}

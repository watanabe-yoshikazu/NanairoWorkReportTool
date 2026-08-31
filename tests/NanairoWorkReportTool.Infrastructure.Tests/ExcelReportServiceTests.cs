using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Core.Services;
using NanairoWorkReportTool.Infrastructure.Excel;

namespace NanairoWorkReportTool.Infrastructure.Tests;

public sealed class ExcelReportServiceTests
{
    [Fact]
    public async Task Export_IsMacroFreePreservesPrintLayoutAndImportsMetadata()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"NanairoExcel-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var document = BuildDocument();
            document.Entries[0].ApplyStatus(WorkStatus.PaidLeave);
            var calculator = new WorkReportCalculator();
            var service = new OpenXmlExcelReportService();
            var path = Path.Combine(directory, service.BuildFileName(document));
            await service.ExportAsync(path, document, calculator.Calculate(document, new DateOnly(2026, 7, 31)));
            using (var archive = ZipFile.OpenRead(path))
                Assert.DoesNotContain(archive.Entries, x => x.FullName.Contains("vbaProject", StringComparison.OrdinalIgnoreCase));
            using (var spreadsheet = SpreadsheetDocument.Open(path, false))
            {
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
                var workbook = spreadsheet.WorkbookPart?.Workbook ?? throw new InvalidDataException("Workbook定義がありません。");
                var sheets = workbook.Sheets ?? throw new InvalidDataException("シート定義がありません。");
                var reportSheet = sheets.Elements<Sheet>().Single(x => x.Name == "作業報告書");
                var metadataSheet = sheets.Elements<Sheet>().Single(x => x.Name == "_NanairoData");
                var reportId = reportSheet.Id?.Value ?? throw new InvalidDataException("作業報告書シートIDがありません。");
                var worksheet = ((WorksheetPart)spreadsheet.WorkbookPart.GetPartById(reportId)).Worksheet
                                ?? throw new InvalidDataException("Worksheet定義がありません。");
                var pageSetup = worksheet.GetFirstChild<PageSetup>();
                Assert.NotNull(pageSetup);
                Assert.DoesNotContain(worksheet.Descendants<Cell>(), x => x.CellReference?.Value?.StartsWith("J", StringComparison.Ordinal) == true);
                Assert.DoesNotContain(worksheet.Descendants<Formula>(), x => x.Text.Contains("$J", StringComparison.Ordinal));
                Assert.Equal(SheetStateValues.VeryHidden, metadataSheet.State?.Value);
                Assert.Equal(OrientationValues.Portrait, pageSetup.Orientation?.Value);
                Assert.Equal(74U, pageSetup.Scale?.Value);
                var definedNames = workbook.DefinedNames ?? throw new InvalidDataException("定義名がありません。");
                Assert.Contains(definedNames.Elements<DefinedName>(), x => x.Name == "_xlnm.Print_Area" && x.Text.Contains("$A$1:$I$40"));
                Assert.Equal("有給休暇", CellText(worksheet, "A9"));
                Assert.Equal(string.Empty, CellText(worksheet, "I9"));
                var paidLeaveRule = Assert.Single(worksheet.Descendants<ConditionalFormattingRule>(), rule =>
                    rule.GetFirstChild<Formula>()?.Text.Contains("PaidLeave", StringComparison.Ordinal) == true);
                Assert.True(paidLeaveRule.StopIfTrue?.Value);
                var holidayRules = worksheet.Descendants<ConditionalFormattingRule>().Where(rule =>
                    rule.GetFirstChild<Formula>()?.Text.Contains("CompanyHoliday", StringComparison.Ordinal) == true).ToArray();
                Assert.NotEmpty(holidayRules);
                Assert.All(holidayRules, rule =>
                    Assert.Contains("=\"Holiday\"", rule.GetFirstChild<Formula>()?.Text, StringComparison.Ordinal));
                var differentialFormats = spreadsheet.WorkbookPart.WorkbookStylesPart?.Stylesheet?.DifferentialFormats
                                          ?? throw new InvalidDataException("条件付き書式のスタイル定義がありません。");
                var paidLeaveFormat = differentialFormats.Elements<DifferentialFormat>()
                    .ElementAt((int)(paidLeaveRule.FormatId?.Value ?? throw new InvalidDataException("有給休暇の書式IDがありません。")));
                var fill = paidLeaveFormat.GetFirstChild<Fill>()?.GetFirstChild<PatternFill>();
                Assert.Equal(PatternValues.Solid, fill?.PatternType?.Value);
                Assert.Equal("FFFFFFFF", fill?.GetFirstChild<ForegroundColor>()?.Rgb?.Value);
                var elements = worksheet.ChildElements.ToList();
                var pageMargins = worksheet.GetFirstChild<PageMargins>() ?? throw new InvalidDataException("ページ余白設定がありません。");
                Assert.True(elements.IndexOf(paidLeaveRule.Parent!) < elements.IndexOf(pageMargins));
            }
            var restored = await service.ImportAsync(path);
            Assert.Equal(document.ReporterName, restored.ReporterName);
            Assert.Equal(document.Entries[0].WorkStatus, restored.Entries[0].WorkStatus);
            Assert.Equal(document.Entries[0].WorkContent, restored.Entries[0].WorkContent);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static WorkReportDocument BuildDocument()
    {
        var calculator = new WorkReportCalculator();
        var document = calculator.CreateMonth(new DateTime(2026, 7, 1), new Dictionary<DateOnly, string> { [new(2026, 7, 20)] = "海の日" });
        document.ReporterName = "渡辺";
        document.CompanyName = "株式会社リンク";
        document.Destination = "株式会社ナナイロ 御中";
        document.OutputDirectory = Path.GetTempPath();
        foreach (var entry in document.Entries.Where(x => x.WorkStatus == WorkStatus.Normal)) entry.WorkContent = "開発・試験";
        return document;
    }

    private static string? CellText(Worksheet worksheet, string reference)
        => worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == reference).InlineString?.InnerText;
}

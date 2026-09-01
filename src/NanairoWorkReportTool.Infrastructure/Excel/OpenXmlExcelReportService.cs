using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.Core.Services;
using NanairoWorkReportTool.Infrastructure.Persistence;

namespace NanairoWorkReportTool.Infrastructure.Excel;

public sealed class OpenXmlExcelReportService : IExcelReportService
{
    private const string ReportSheetName = "作業報告書";
    private const string MetadataSheetName = "_NanairoData";
    private const string Signature = "NanairoWorkReportTool:1";

    public string BuildFileName(WorkReportDocument document)
        => $"作業報告書_{document.TargetMonth:yyyyMM}_{SanitizeFileName(document.ReporterName ?? string.Empty)}.xlsx";

    public async Task ExportAsync(string path, WorkReportDocument document, MonthlySummary summary, CancellationToken cancellationToken = default)
    {
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var temporary = fullPath + $".{Guid.NewGuid():N}.tmp";
        try
        {
            await CopyTemplateAsync(temporary, cancellationToken);
            using (var spreadsheet = SpreadsheetDocument.Open(temporary, true))
            {
                var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidDataException("帳票テンプレートにWorkbookがありません。");
                var workbook = workbookPart.Workbook ?? throw new InvalidDataException("帳票テンプレートのWorkbook定義がありません。");
                var sheet = workbook.Sheets?.Elements<Sheet>().FirstOrDefault(x => x.Name?.Value == ReportSheetName)
                            ?? throw new InvalidDataException($"帳票テンプレートに「{ReportSheetName}」シートがありません。");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                WriteReport(workbookPart, worksheetPart, document, summary);
                EnsurePrintSettings(workbookPart, worksheetPart, sheet);
                WriteMetadata(workbookPart, document);
                workbook.CalculationProperties ??= new CalculationProperties();
                workbook.CalculationProperties.FullCalculationOnLoad = true;
                workbook.CalculationProperties.ForceFullCalculation = true;
                workbook.Save();
            }
            ValidateMacroFree(temporary);
            File.Move(temporary, fullPath, true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    public Task<WorkReportDocument> ImportAsync(string path, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("初版で読み込める帳票は本ツールが生成した .xlsx のみです。");

        using var spreadsheet = SpreadsheetDocument.Open(path, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidDataException("Workbookを読み取れません。");
        var workbook = workbookPart.Workbook ?? throw new InvalidDataException("Workbook定義を読み取れません。");
        var metadata = workbook.Sheets?.Elements<Sheet>().FirstOrDefault(x => x.Name?.Value == MetadataSheetName)
                       ?? throw new InvalidDataException("本ツールの帳票メタデータがありません。");
        var metadataPart = (WorksheetPart)workbookPart.GetPartById(metadata.Id!);
        var signature = ReadCellText(workbookPart, metadataPart, "A1");
        if (signature != Signature) throw new InvalidDataException("本ツールの帳票形式と一致しません。");
        var chunks = GetWorksheet(metadataPart).Descendants<Cell>()
            .Where(cell => CellRow(cell.CellReference?.Value) >= 3 && CellColumn(cell.CellReference?.Value) == "A")
            .OrderBy(cell => CellRow(cell.CellReference?.Value))
            .Select(cell => ReadCellText(workbookPart, cell))
            .Where(value => value is not null);
        var document = JsonDocumentStore.Deserialize(string.Concat(chunks));
        OverlayVisibleCells(workbookPart, document);
        return Task.FromResult(document);
    }

    private static void WriteReport(WorkbookPart workbookPart, WorksheetPart worksheetPart, WorkReportDocument document, MonthlySummary summary)
    {
        SetText(worksheetPart, "A2", document.Destination?.Trim() ?? string.Empty);
        SetNumber(worksheetPart, "A4", document.TargetMonth.ToOADate());
        SetText(worksheetPart, "G4", document.CompanyName?.Trim() ?? string.Empty);
        SetText(worksheetPart, "G5", document.ReporterName?.Trim() ?? string.Empty);
        for (var index = 0; index < 31; index++)
        {
            var row = 9 + index;
            var entry = index < document.Entries.Count ? document.Entries[index] : null;
            ClearContents(worksheetPart, row, "ABCDEFGHI");
            if (entry is null) continue;
            SetText(worksheetPart, $"A{row}", entry.WorkStatus switch
            {
                WorkStatus.PaidLeave => "有給休暇",
                WorkStatus.Normal => entry.WorkContent ?? string.Empty,
                _ => string.Empty
            });
            SetFormulaNumber(worksheetPart, $"B{row}", $"G{row}-H{row}", entry.WorkMinutes / 60d);
            SetNumber(worksheetPart, $"C{row}", entry.Date.ToDateTime(TimeOnly.MinValue).ToOADate());
            SetFormulaNumber(worksheetPart, $"D{row}", $"C{row}", entry.Date.ToDateTime(TimeOnly.MinValue).ToOADate());
            if (entry.StartMinutes.HasValue) SetNumber(worksheetPart, $"E{row}", entry.StartMinutes.Value / 1440d);
            if (entry.EndMinutes.HasValue) SetNumber(worksheetPart, $"F{row}", entry.EndMinutes.Value / 1440d);
            SetFormulaNumber(worksheetPart, $"G{row}", $"(F{row}-E{row})*24", entry.GrossMinutes / 60d);
            if (entry.BreakMinutes.HasValue) SetNumber(worksheetPart, $"H{row}", entry.BreakMinutes.Value / 60d);
            SetText(worksheetPart, $"I{row}", entry.WorkStatus == WorkStatus.PaidLeave ? string.Empty : entry.GetReportRemark());
        }

        SetFormulaNumber(worksheetPart, "B40", "SUM(B9:B39)", summary.ForecastMinutes / 60d);
        SetFormulaNumber(worksheetPart, "G40", "SUM(G9:G39)", document.Entries.Sum(x => x.GrossMinutes) / 60d);
        SetFormulaNumber(worksheetPart, "H40", "SUM(H9:H39)", document.Entries.Sum(x => x.BreakMinutes ?? 0) / 60d);
        SetText(worksheetPart, "A41", BuildVerificationSummary(summary));
        for (var column = 'B'; column <= 'I'; column++) ClearContents(worksheetPart, 41, column.ToString());
        EnsureMergedCell(worksheetPart, "A41:I41");
        var verificationRow = GetWorksheet(worksheetPart).GetFirstChild<SheetData>()?
            .Elements<Row>().FirstOrDefault(row => row.RowIndex?.Value == 41U);
        if (verificationRow is not null)
        {
            verificationRow.Height = 22.5D;
            verificationRow.CustomHeight = true;
        }
        UpdateConditionalFormatting(workbookPart, worksheetPart);
        GetWorksheet(worksheetPart).Save();
    }

    private static string BuildVerificationSummary(MonthlySummary summary)
        => $"基準日数：{summary.WeekdayWorkDays}日（稼働対象日{summary.BaselineCandidateDays}日－公休日{summary.BaselineCompanyHolidayDays}日）、" +
           $"基準時間：{FormatHours(summary.BaselineMinutes, true)}H±{FormatHours(summary.UpperMinutes - summary.BaselineMinutes, false)}H、" +
           $"稼働実績：{FormatHours(summary.ForecastMinutes, true)}H";

    private static string FormatHours(int minutes, bool minimumOneDecimal)
        => (minutes / 60d).ToString(minimumOneDecimal ? "0.0#" : "0.##", CultureInfo.InvariantCulture);

    private static void OverlayVisibleCells(WorkbookPart workbookPart, WorkReportDocument document)
    {
        var workbook = workbookPart.Workbook ?? throw new InvalidDataException("Workbook定義を読み取れません。");
        var sheet = workbook.Sheets?.Elements<Sheet>().FirstOrDefault(x => x.Name?.Value == ReportSheetName)
                    ?? throw new InvalidDataException("作業報告書シートがありません。");
        var part = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        document.Destination = ReadCellText(workbookPart, part, "A2") ?? document.Destination;
        document.CompanyName = ReadCellText(workbookPart, part, "G4") ?? document.CompanyName;
        document.ReporterName = ReadCellText(workbookPart, part, "G5") ?? document.ReporterName;
        for (var i = 0; i < Math.Min(31, document.Entries.Count); i++)
        {
            var row = 9 + i;
            var entry = document.Entries[i];
            if (entry.WorkStatus == WorkStatus.Normal)
            {
                entry.WorkContent = ReadCellText(workbookPart, part, $"A{row}");
                entry.StartMinutes = ReadTimeMinutes(workbookPart, part, $"E{row}");
                entry.EndMinutes = ReadTimeMinutes(workbookPart, part, $"F{row}");
                entry.BreakMinutes = ReadHourMinutes(workbookPart, part, $"H{row}");
            }
            else
            {
                entry.ClearTime();
            }
            var remark = ReadCellText(workbookPart, part, $"I{row}");
            if (entry.WorkStatus == WorkStatus.CompanyHoliday) entry.CompanyHolidayName = remark;
            else if (entry.WorkStatus == WorkStatus.Normal) entry.Note = remark;
        }
    }

    private static void WriteMetadata(WorkbookPart workbookPart, WorkReportDocument document)
    {
        var workbook = workbookPart.Workbook ?? throw new InvalidDataException("Workbook定義を読み取れません。");
        var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());
        var sheet = sheets.Elements<Sheet>().FirstOrDefault(x => x.Name?.Value == MetadataSheetName);
        WorksheetPart part;
        if (sheet is null)
        {
            part = workbookPart.AddNewPart<WorksheetPart>();
            part.Worksheet = new Worksheet(new SheetData());
            var id = sheets.Elements<Sheet>().Select(x => x.SheetId?.Value ?? 0).DefaultIfEmpty().Max() + 1;
            sheet = new Sheet { Id = workbookPart.GetIdOfPart(part), SheetId = id, Name = MetadataSheetName, State = SheetStateValues.VeryHidden };
            sheets.Append(sheet);
        }
        else
        {
            part = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            part.Worksheet = new Worksheet(new SheetData());
            sheet.State = SheetStateValues.VeryHidden;
        }

        SetText(part, "A1", Signature);
        SetNumber(part, "A2", WorkReportDocument.CurrentSchemaVersion);
        var json = JsonDocumentStore.Serialize(document);
        const int chunkSize = 30_000;
        for (var offset = 0; offset < json.Length; offset += chunkSize)
            SetText(part, $"A{3 + offset / chunkSize}", json.Substring(offset, Math.Min(chunkSize, json.Length - offset)));
        SetText(part, "B2", "workStatusCodes");
        for (var index = 0; index < 31; index++)
        {
            var code = index < document.Entries.Count ? BuildHelperCode(document.Entries[index]) : "Empty";
            SetText(part, $"B{3 + index}", code);
        }
        EnsureWorkStatusDefinedName(workbook);
        GetWorksheet(part).Save();
    }

    private static async Task CopyTemplateAsync(string target, CancellationToken cancellationToken)
    {
        var assembly = typeof(OpenXmlExcelReportService).Assembly;
        await using var input = assembly.GetManifestResourceStream("NanairoWorkReportTool.Infrastructure.Assets.ReportTemplate.xlsx")
                                ?? throw new FileNotFoundException("同梱の帳票テンプレートがありません。");
        await using var output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, true);
        await input.CopyToAsync(output, cancellationToken);
    }

    private static void EnsureWorkStatusDefinedName(Workbook workbook)
    {
        var definedNames = workbook.DefinedNames ?? workbook.AppendChild(new DefinedNames());
        var definedName = definedNames.Elements<DefinedName>()
            .FirstOrDefault(x => x.Name?.Value == "_NanairoWorkStatus");
        if (definedName is null)
        {
            definedName = new DefinedName { Name = "_NanairoWorkStatus" };
            definedNames.Append(definedName);
        }
        definedName.Text = $"'{MetadataSheetName}'!$B$3:$B$33";
    }

    private static void EnsurePrintSettings(WorkbookPart workbookPart, WorksheetPart worksheetPart, Sheet sheet)
    {
        var worksheet = GetWorksheet(worksheetPart);
        var margins = worksheet.GetFirstChild<PageMargins>();
        if (margins is null)
        {
            margins = new PageMargins();
            var setupPosition = worksheet.GetFirstChild<PageSetup>();
            if (setupPosition is null) worksheet.Append(margins); else worksheet.InsertBefore(margins, setupPosition);
        }
        margins.Left = 0.7874D;
        margins.Right = 0.3937D;
        margins.Top = 0.3937D;
        margins.Bottom = 0.3937D;
        margins.Header = 0.1969D;
        margins.Footer = 0.1969D;

        var pageSetup = worksheet.GetFirstChild<PageSetup>();
        if (pageSetup is null)
        {
            pageSetup = new PageSetup();
            worksheet.InsertAfter(pageSetup, margins);
        }
        pageSetup.PaperSize = 9U;
        pageSetup.Orientation = OrientationValues.Portrait;
        pageSetup.Scale = 72U;

        var workbook = workbookPart.Workbook ?? throw new InvalidDataException("Workbook定義を読み取れません。");
        var sheets = workbook.Sheets?.Elements<Sheet>().ToList()
                     ?? throw new InvalidDataException("帳票テンプレートにシート情報がありません。");
        var localSheetId = (uint)sheets.IndexOf(sheet);
        var definedNames = workbook.DefinedNames ?? workbook.AppendChild(new DefinedNames());
        var printArea = definedNames.Elements<DefinedName>().FirstOrDefault(x =>
            x.Name?.Value == "_xlnm.Print_Area" && x.LocalSheetId?.Value == localSheetId);
        if (printArea is null)
        {
            printArea = new DefinedName { Name = "_xlnm.Print_Area", LocalSheetId = localSheetId };
            definedNames.Append(printArea);
        }
        printArea.Text = $"'{ReportSheetName}'!$A$1:$I$41";
    }

    private static void EnsureMergedCell(WorksheetPart part, string reference)
    {
        var worksheet = GetWorksheet(part);
        var mergeCells = worksheet.GetFirstChild<MergeCells>();
        if (mergeCells is null)
        {
            mergeCells = new MergeCells();
            var conditionalFormatting = worksheet.GetFirstChild<ConditionalFormatting>();
            if (conditionalFormatting is null) worksheet.Append(mergeCells);
            else worksheet.InsertBefore(mergeCells, conditionalFormatting);
        }

        if (!mergeCells.Elements<MergeCell>().Any(cell => cell.Reference?.Value == reference))
        {
            mergeCells.Append(new MergeCell { Reference = reference });
            mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
        }
    }

    private static void UpdateConditionalFormatting(WorkbookPart workbookPart, WorksheetPart part)
    {
        var worksheet = GetWorksheet(part);
        foreach (var rule in worksheet.Descendants<ConditionalFormattingRule>())
        {
            var formula = rule.GetFirstChild<Formula>();
            if (formula?.Text?.Contains("祝日", StringComparison.Ordinal) == true)
                formula.Text = "OR(INDEX(_NanairoWorkStatus,ROW()-8)=\"CompanyHoliday\",INDEX(_NanairoWorkStatus,ROW()-8)=\"Holiday\")";
        }

        var styles = workbookPart.WorkbookStylesPart?.Stylesheet
                     ?? throw new InvalidDataException("帳票テンプレートのスタイル定義がありません。");
        var differentialFormats = styles.DifferentialFormats ?? styles.AppendChild(new DifferentialFormats());
        var differentialFormatId = (uint)differentialFormats.Elements<DifferentialFormat>().Count();
        differentialFormats.Append(new DifferentialFormat(
            new Fill(new PatternFill(
                new ForegroundColor { Rgb = "FFFFFFFF" },
                new BackgroundColor { Rgb = "FFFFFFFF" })
            { PatternType = PatternValues.Solid })));
        differentialFormats.Count = (uint)differentialFormats.Elements<DifferentialFormat>().Count();
        styles.Save();

        foreach (var rule in worksheet.Descendants<ConditionalFormattingRule>())
            rule.Priority = (rule.Priority?.Value ?? 0) + 1;

        var paidLeaveRule = new ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.Expression,
            FormatId = differentialFormatId,
            Priority = 1,
            StopIfTrue = true
        };
        paidLeaveRule.Append(new Formula("INDEX(_NanairoWorkStatus,ROW()-8)=\"PaidLeave\""));
        var paidLeaveFormatting = new ConditionalFormatting
        {
            SequenceOfReferences = new ListValue<StringValue> { InnerText = "A9:I39" }
        };
        paidLeaveFormatting.Append(paidLeaveRule);
        var pageMargins = worksheet.GetFirstChild<PageMargins>();
        if (pageMargins is null) worksheet.Append(paidLeaveFormatting);
        else worksheet.InsertBefore(paidLeaveFormatting, pageMargins);
    }

    private static string BuildHelperCode(WorkDayEntry entry) => entry.WorkStatus switch
    {
        WorkStatus.CompanyHoliday => "CompanyHoliday",
        WorkStatus.PaidLeave => "PaidLeave",
        _ when entry.DayType == DayType.Holiday => "Holiday",
        _ when entry.DayType == DayType.Saturday => "Saturday",
        _ when entry.DayType == DayType.Sunday => "Sunday",
        _ => "Weekday"
    };

    private static void ClearContents(WorksheetPart part, int row, string columns)
    {
        foreach (var column in columns)
        {
            var cell = GetOrCreateCell(part, $"{column}{row}");
            cell.RemoveAllChildren<CellFormula>();
            cell.RemoveAllChildren<CellValue>();
            cell.RemoveAllChildren<InlineString>();
            cell.DataType = null;
        }
    }

    private static void SetText(WorksheetPart part, string reference, string value)
    {
        var cell = GetOrCreateCell(part, reference);
        cell.RemoveAllChildren();
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static void SetNumber(WorksheetPart part, string reference, double value)
    {
        var cell = GetOrCreateCell(part, reference);
        cell.RemoveAllChildren();
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(value.ToString("G17", CultureInfo.InvariantCulture));
    }

    private static void SetFormulaNumber(WorksheetPart part, string reference, string formula, double value)
    {
        var cell = GetOrCreateCell(part, reference);
        cell.RemoveAllChildren();
        cell.DataType = CellValues.Number;
        cell.CellFormula = new CellFormula(formula);
        cell.CellValue = new CellValue(value.ToString("G17", CultureInfo.InvariantCulture));
    }

    private static Cell GetOrCreateCell(WorksheetPart part, string reference)
    {
        var worksheet = GetWorksheet(part);
        var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
        var rowIndex = (uint)CellRow(reference);
        var row = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex?.Value == rowIndex);
        if (row is null)
        {
            row = new Row { RowIndex = rowIndex };
            var next = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex?.Value > rowIndex);
            if (next is null) sheetData.Append(row); else sheetData.InsertBefore(row, next);
        }
        var existing = row.Elements<Cell>().FirstOrDefault(x => x.CellReference?.Value == reference);
        if (existing is not null) return existing;
        var cell = new Cell { CellReference = reference };
        var nextCell = row.Elements<Cell>().FirstOrDefault(x => ColumnNumber(x.CellReference?.Value) > ColumnNumber(reference));
        if (nextCell is null) row.Append(cell); else row.InsertBefore(cell, nextCell);
        return cell;
    }

    private static string? ReadCellText(WorkbookPart workbookPart, WorksheetPart part, string reference)
    {
        var cell = GetWorksheet(part).Descendants<Cell>().FirstOrDefault(x => x.CellReference?.Value == reference);
        return cell is null ? null : ReadCellText(workbookPart, cell);
    }

    private static string? ReadCellText(WorkbookPart workbookPart, Cell cell)
    {
        if (cell.DataType?.Value == CellValues.InlineString) return cell.InlineString?.InnerText;
        var value = cell.CellValue?.InnerText;
        if (cell.DataType?.Value == CellValues.SharedString && int.TryParse(value, out var index))
            return workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(index)?.InnerText;
        return value;
    }

    private static int? ReadTimeMinutes(WorkbookPart workbookPart, WorksheetPart part, string reference)
    {
        var text = ReadCellText(workbookPart, part, reference);
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? (int)Math.Round(value * 1440) : null;
    }

    private static int? ReadHourMinutes(WorkbookPart workbookPart, WorksheetPart part, string reference)
    {
        var text = ReadCellText(workbookPart, part, reference);
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? (int)Math.Round(value * 60) : null;
    }

    private static int CellRow(string? reference)
        => int.TryParse(new string((reference ?? string.Empty).SkipWhile(char.IsLetter).ToArray()), out var row) ? row : 0;
    private static string CellColumn(string? reference) => new((reference ?? string.Empty).TakeWhile(char.IsLetter).ToArray());
    private static int ColumnNumber(string? reference)
    {
        var number = 0;
        foreach (var character in CellColumn(reference)) number = number * 26 + character - 'A' + 1;
        return number;
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sanitized = new string(value.Trim().Select(character => invalid.Contains(character) ? '_' : character).ToArray()).TrimEnd('.', ' ');
        return string.IsNullOrWhiteSpace(sanitized) ? "入力者" : sanitized;
    }

    private static void ValidateMacroFree(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        if (archive.Entries.Any(entry => entry.FullName.Equals("xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException("生成帳票にVBAが残っています。");
    }

    private static Worksheet GetWorksheet(WorksheetPart part)
        => part.Worksheet ?? throw new InvalidDataException("帳票のWorksheet定義を読み取れません。");
}

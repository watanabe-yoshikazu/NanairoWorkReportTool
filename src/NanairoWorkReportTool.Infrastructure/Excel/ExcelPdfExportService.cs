using System.Runtime.InteropServices;
using NanairoWorkReportTool.Core.Services;

namespace NanairoWorkReportTool.Infrastructure.Excel;

public sealed class ExcelPdfExportService : IPdfExportService
{
    public Task ExportAsync(string excelPath, string pdfPath, CancellationToken cancellationToken = default)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            object? application = null;
            object? workbooks = null;
            object? workbook = null;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                var excelType = Type.GetTypeFromProgID("Excel.Application")
                                ?? throw new InvalidOperationException("デスクトップ版Microsoft Excelがインストールされていません。");
                application = Activator.CreateInstance(excelType)
                              ?? throw new InvalidOperationException("Microsoft Excelを起動できませんでした。");
                dynamic excel = application;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.AutomationSecurity = 3;
                workbooks = excel.Workbooks;
                dynamic books = workbooks;
                workbook = books.Open(Path.GetFullPath(excelPath), 0, true);
                dynamic book = workbook;
                Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(pdfPath))!);
                book.ExportAsFixedFormat(0, Path.GetFullPath(pdfPath), 0, true, false);
                book.Close(false);
                workbook = null;
                excel.Quit();
                completion.TrySetResult();
            }
            catch (Exception exception)
            {
                completion.TrySetException(new InvalidOperationException($"PDF出力に失敗しました。{exception.Message}", exception));
            }
            finally
            {
                Release(workbook);
                Release(workbooks);
                if (application is not null)
                {
                    try { ((dynamic)application).Quit(); } catch { }
                }
                Release(application);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        });
        thread.IsBackground = true;
        thread.Name = "Nanairo Excel PDF Export";
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
}

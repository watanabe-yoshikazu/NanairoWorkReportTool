using System.Windows;
using NanairoWorkReportTool.Core.Services;
using NanairoWorkReportTool.Infrastructure.Excel;
using NanairoWorkReportTool.Infrastructure.Holidays;
using NanairoWorkReportTool.Infrastructure.Persistence;
using NanairoWorkReportTool.ViewModels;

namespace NanairoWorkReportTool;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        var calculator = new WorkReportCalculator();
        var documentStore = new JsonDocumentStore();
        var settingsStore = new SettingsStore();
        var holidayProvider = new HolidayCsvProvider();
        var viewModel = new MainWindowViewModel(
            calculator,
            new ReportValidator(calculator),
            documentStore,
            new OpenXmlExcelReportService(),
            new ExcelPdfExportService(),
            holidayProvider,
            settingsStore);
        var window = new MainWindow { DataContext = viewModel };
        MainWindow = window;
        window.Show();
    }
}

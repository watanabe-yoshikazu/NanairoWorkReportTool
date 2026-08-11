using System.Diagnostics;
using System.IO;
using System.Windows;

namespace NanairoWorkReportTool;

public partial class OutputCompletedWindow : Window
{
    public string ExcelPath { get; }
    public string? PdfPath { get; }

    public OutputCompletedWindow(string excelPath, string? pdfPath)
    {
        ExcelPath = excelPath;
        PdfPath = pdfPath;
        InitializeComponent();
        DataContext = this;
        OpenPdfButton.IsEnabled = File.Exists(pdfPath);
    }

    private static void Open(string path) => Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    private void OpenExcel_Click(object sender, RoutedEventArgs e) => Open(ExcelPath);
    private void OpenPdf_Click(object sender, RoutedEventArgs e) { if (PdfPath is not null) Open(PdfPath); }
    private void OpenFolder_Click(object sender, RoutedEventArgs e) => Open(Path.GetDirectoryName(ExcelPath)!);
    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}

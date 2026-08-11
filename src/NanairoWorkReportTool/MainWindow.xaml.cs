using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using NanairoWorkReportTool.Core.Domain;
using NanairoWorkReportTool.ViewModels;

namespace NanairoWorkReportTool;

public partial class MainWindow : Window
{
    private bool closeApproved;

    public MainWindow() => InitializeComponent();

    private MainWindowViewModel ViewModel => (MainWindowViewModel)DataContext;

    private async void Window_Loaded(object sender, RoutedEventArgs e) => await ViewModel.InitializeAsync();

    private async void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (closeApproved) return;
        e.Cancel = true;
        if (await ViewModel.TryCloseAsync())
        {
            closeApproved = true;
            Close();
        }
    }

    private async void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is not string[] files || files.Length == 0) return;
        var extension = Path.GetExtension(files[0]);
        if (extension.Equals(".nwr", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            await ViewModel.TryOpenPathAsync(files[0]);
        else
            MessageBox.Show(".nwr または本ツールが生成した .xlsx をドロップしてください。", "未対応ファイル", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private IEnumerable<WorkDayEntry> SelectedEntries()
        => EntriesGrid.SelectedItems.Cast<WorkDayEntry>().ToArray();

    private void StandardTime_Click(object sender, RoutedEventArgs e) => ViewModel.SetStandardTime(SelectedEntries());
    private void ClearTime_Click(object sender, RoutedEventArgs e) => ViewModel.ClearTime(SelectedEntries());
    private void CopyPrevious_Click(object sender, RoutedEventArgs e) => ViewModel.CopyPreviousContent(SelectedEntries());
    private void BulkStatus_Click(object sender, RoutedEventArgs e)
    {
        if (BulkStatusCombo.SelectedValue is WorkStatus status) ViewModel.ApplyWorkStatus(SelectedEntries(), status);
    }
    private void BulkContent_Click(object sender, RoutedEventArgs e) => ViewModel.ApplyWorkContent(SelectedEntries(), BulkContentCombo.Text);
    private void BulkHoliday_Click(object sender, RoutedEventArgs e) => ViewModel.ApplyCompanyHoliday(SelectedEntries(), BulkHolidayNameText.Text);

    private void ValidationGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (ValidationGrid.SelectedItem is not ValidationIssue issue || issue.Date is null) return;
        var entry = ViewModel.Document.Entries.FirstOrDefault(item => item.Date == issue.Date.Value);
        if (entry is null) return;
        EntriesGrid.SelectedItem = entry;
        EntriesGrid.ScrollIntoView(entry);
        EntriesGrid.Focus();
    }
}

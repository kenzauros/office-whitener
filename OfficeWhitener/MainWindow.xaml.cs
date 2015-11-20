using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace OfficeWhitener
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void Window_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files != null && files.Length > 0)
            {
                files = files.Where(f => OfficeDocumentWhitener.AvailableExtensions.Any(ex => ex == Path.GetExtension(f).Replace(".", ""))).ToArray();
                var vm = this.DataContext as MainWindowVM;

                var dispatcher = Application.Current.Dispatcher;
                foreach (var file in files)
                {
                    var item = new OfficeFileItemVM(file);
                    vm.Items.Add(item);
                    Task.Run(() =>
                    {
                        try
                        {
                            OfficeDocumentWhitener.RemovePersonalInfo(file);
                            dispatcher.Invoke(() => { item.State = "OK"; });
                        }
                        catch (Exception ex)
                        {
                            dispatcher.Invoke(() =>
                            {
                                item.State = "NG";
                                item.ErrorMessage = ex.Message;
                            });
                        }
                    });
                }
            }
        }
    }
}

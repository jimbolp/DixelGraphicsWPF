using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Threading;

namespace DixelGraphics
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static bool _cancel = false;
        public static bool Cancel { get { return _cancel; } }
        private string loadedFile = "";
        private bool isRunning = false;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void filePathTextBox_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void filePathTextBox_Drop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                filePathTextBox.Text = files[0];
                if (isExcelFile(filePathTextBox.Text))
                {
                    loadedFile = filePathTextBox.Text;
                }
                else
                {
                    filePathTextBox.Text = "";
                }
            }
        }

        private bool isExcelFile(string text)
        {
            if(System.IO.Path.GetExtension(text) == ".xls" || System.IO.Path.GetExtension(text) == ".xlsx")
            {
                return true;
            }
            return false;
        }

        private void filePathTextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            progBarConvert.Value = 0;
            progBarChart.Value = 0;
            if (!isExcelFile(filePathTextBox.Text) || isRunning)
            {
                return;
            }
            StartWorking();
        }

        /// <summary>
        /// Here starts the whole process of working on the file!
        /// </summary>
        private void StartWorking()
        {
            if ((!graphicsCheckBox.IsChecked ?? false) && (!printChckBox.IsChecked ?? false) && (!alterChckBox.IsChecked ?? false))
            {
                MessageBox.Show(
                        "Не сте избрали опция за промяна на стойности, създаване или принтиране на графики!",
                        "Внимание!", MessageBoxButton.OK);
                return;
            }
            ExcelFile excelFile = null;
            try
            {
                excelFile = new ExcelFile(filePathTextBox.Text, printChckBox.IsChecked);
                Thread workThread = new Thread(() =>
                {
                    isRunning = true;
                    if (alterChckBox.Dispatcher.Invoke(() => alterChckBox.IsChecked ?? false))
                    {
                        excelFile.AlterValues();
                    }
                    if (graphicsCheckBox.Dispatcher.Invoke(() => graphicsCheckBox.IsChecked ?? false))
                    {
                        excelFile.CreateGraphics(tempChckBox.Dispatcher.Invoke(() => tempChckBox.IsChecked ?? false),
                            humidChckBox.Dispatcher.Invoke(() => humidChckBox.IsChecked ?? false));
                    }
                    if(printChckBox.Dispatcher.Invoke(() => printChckBox.IsChecked ?? false))
                    {
                        excelFile.PrintGraphics();
                    }
                    excelFile.SaveAs();
                    isRunning = false;
                    excelFile.Dispose();
                });
                workThread.Start();
            }
            catch (Exception)
            {
                if (excelFile != null)
                    excelFile.Dispose();
                //TODO...
            }
        }
        private void graphicsCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            tempChckBox.IsEnabled = true;
            tempChckBox.IsChecked = true;
            humidChckBox.IsEnabled = true;
            humidChckBox.IsChecked = true;
        }

        private void graphicsCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            tempChckBox.IsChecked = false;
            humidChckBox.IsChecked = false;
            tempChckBox.IsEnabled = false;
            humidChckBox.IsEnabled = false;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cancel = true;
        }
    }
}

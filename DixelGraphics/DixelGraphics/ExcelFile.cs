using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Core;
using System.Globalization;
using System.Threading;
using System.Windows;

namespace DixelGraphics
{
    internal class ExcelFile
    {
        private readonly String SaveDir = "";
        Application xlApp = new Application();
        Workbooks xlWBooks = null;
        Workbook xlWBook = null;

        public ExcelFile(string filePath, bool? printGraphics)
        {
            try
            {
                xlApp.DisplayAlerts = false;
                xlApp.ScreenUpdating = false;
                xlApp.Visible = false;
                xlApp.UserControl = false;
                xlApp.Interactive = false;
                xlApp.FileValidation = MsoFileValidationMode.msoFileValidationSkip;
                SaveDir = Path.GetDirectoryName(filePath);
                InitializeExcelObjs(filePath);
            }
            catch(COMException ex)
            {
                MessageBox.Show(ex.Message, "Excel Error!");
                Dispose();
            }
            catch(Exception e)
            {
                Dispose();
                //TODO... 
            }
        }

        private void InitializeExcelObjs(string filePath)
        {
            xlWBooks = xlApp.Workbooks;
            xlWBook = xlWBooks.Open(filePath, IgnoreReadOnlyRecommended: true, ReadOnly: true, Editable: false);
        }

        public void Dispose()
        {
            try
            {
                xlApp.Quit();
            }
            catch (InvalidComObjectException)
            {
                //File probably already closed :D :D
            }
            catch (Exception)
            {
                //MessageBox.Show("Unable to close the application or it's already closed! Check Task Manager :D :D");
            }
            Dispose(xlWBook);
            Dispose(xlWBooks);
            Dispose(xlApp);
        }

        private void Dispose(object obj)
        {
            try
            {
                if(obj.GetType() == typeof(Application))
                {
                    try
                    {
                        (obj as Application).Quit();
                    }
                    catch (Exception)
                    {
                        //Just Release the Application object
                    }
                }
                while (Marshal.ReleaseComObject(obj) > 0) { }
                obj = null;
            }
            catch (COMException)
            {
                obj = null;
                //MessageBox.Show("Com Exception Occured while releasing object " + cEx.ToString());
            }
            catch (Exception)
            {
                obj = null;
                //MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        internal void CreateGraphics(bool temperature, bool humidity)
        {            
            List<Thread> chartThreads = new List<Thread>();
            Sheets xlWorkSheets = null;
            try
            {
                xlWorkSheets = xlWBook.Worksheets;
                int sheetCount = xlWorkSheets.Count;
                int sheetNumber = 1;
                ResetProgBarConvert(sheetCount);
                foreach(Worksheet sheet in xlWorkSheets)
                {
                    if (CancelRequest())
                        return;
                    if (sheet.UsedRange.Value == null)
                        continue;
                    ConvertDateTimeToString(sheet, sheetNumber, sheetCount);
                    Thread t = new Thread(() =>
                    {
                        try
                        {
                            if (temperature)
                            {
                                CreateTemperatureGraphs(sheet);
                            }
                            if (humidity)
                            {
                                CreateHumidityGraphs(sheet);
                            }
                        }
                        catch (NotImplementedException)
                        {
                            
                        }
                    });
                    chartThreads.Add(t);
                    sheetNumber++;                    
                }
                foreach(Thread t in chartThreads)
                {
                    t.Start();
                    t.Join();
                }
            }
            catch (Exception)
            {
                Dispose(xlWorkSheets);
                Dispose();
                //TODO...
            }
        }

        private void SetLabelNote(int sheetNumber, int sheetCount)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            window.labelNotification.Dispatcher.Invoke(() => window.labelNotification.Content = $"Страница {sheetNumber}/{sheetCount}");
        }

        private void ConvertDateTimeToString(Worksheet sheet, int sheetNumber, int sheetCount)
        {
            SetLabelNote(sheetNumber, sheetCount);
            Range range = sheet.UsedRange;
            object[,] usedRange = range.Value;
            int rowCount = range.Rows.Count;
            for (int i = 1; i <= rowCount; ++i)
            {
                UpdateProgBarConvert(i, rowCount);
                string currentCell = FixDateTimeFormat(usedRange[i, 1]);
                if (DateTime.TryParse(currentCell, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime d))
                    usedRange[i, 1] = "\'" + currentCell;
            }
            range.Value = usedRange;
        }

        private string FixDateTimeFormat(object date)
        {
            if (!date.ToString().Contains("."))
            {
                return date.ToString();
            }
            if (date.ToString().IndexOf('.') > 10)
            {
                date = date.ToString().Replace('.', ':');
            }
            else
            {
                string smalldate = date.ToString().Substring(0, 10);
                string fixedSmallDate = "";
                if (smalldate.Contains("."))
                {
                    fixedSmallDate = smalldate.Replace('.', '/');
                    date = date.ToString().Replace(smalldate, fixedSmallDate);
                    FixDateTimeFormat(date);
                }
                
            }
            
            return date.ToString();
        }

        private void ResetProgBarConvert(int max)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Value = 0);
            window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Maximum = max);
            window.progBarConvertText.Dispatcher.Invoke(() => window.progBarConvertText.Text = "");
            
        }

        private void ResetProgBarChart(int max)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value = 0);
            window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum = max);
            window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = "");

        }

        private void UpdateProgBarConvert(int val, int max)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            if (window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Maximum < max))
            {
                window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Maximum = max);
            }
            window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Value = val);
            window.progBarConvertText.Dispatcher.Invoke(() => window.progBarConvertText.Text = $"{val}/{max}");
        }

        private void UpdateProgBarConvert(int val)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Value = val);
            window.progBarConvertText.Dispatcher.Invoke(() => window.progBarConvertText.Text = $"{val}/{window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Maximum)}");
        }

        private void UpdateProgBarChart(int val, int max)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            if (window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum < max))
            {
                window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum = max);
            }
            window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value = val);
            window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"{val}/{max}");
        }

        private void UpdateProgBarChart(int val)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value = val);
            window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"{val}/{window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum)}");
        }

        private void CreateHumidityGraphs(Worksheet sheet) => throw new NotImplementedException();
        private void CreateTemperatureGraphs(Worksheet sheet)
        {
            if (CancelRequest())
            {
                return;
            }

        }

        private bool CancelRequest()
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            if (MainWindow.Cancel)
            {
                window.labelNotification.Dispatcher.Invoke(() => window.labelNotification.Content = "Work canceled...!");
                return true;
            }
            return false;
        }

        internal void AlterValues()
        {
            
        }

        internal void PrintGraphics()
        {

        }

        internal void SaveAs()
        {

        }
    }
}
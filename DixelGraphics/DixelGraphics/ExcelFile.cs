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
using Microsoft.Win32;

namespace DixelGraphics
{
    internal class ExcelFile
    {
        const double COLD_MIN = 2.5;
        const double COLD_MAX = 7.5;
        const double TEMP_MIN = 16.0;
        const double TEMP_MAX = 24.0;
        const double HUMID_MIN = 35.0;
        const double HUMID_MAX = 55.0;

        private readonly String SaveDir = "";
        private readonly String SaveFileName = "";
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
                SaveFileName = Path.GetFileName(filePath);
                InitializeExcelObjs(filePath);
            }
            catch(COMException ex)
            {
                MessageBox.Show(ex.Message, "Excel Error!");
                Dispose();
                throw ex;
            }
            catch(Exception e)
            {
                Dispose();
                throw e;
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
                        catch (Exception)
                        {
                            Dispose(sheet);
                        }
                    });
                    chartThreads.Add(t);
                    sheetNumber++;
                    Thread.Sleep(1);
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
            CultureInfo cInfo = new CultureInfo("bg-BG");
            cInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            cInfo.DateTimeFormat.ShortTimePattern = "hh:mm:ss";
            cInfo.DateTimeFormat.DateSeparator = "/";

            SetLabelNote(sheetNumber, sheetCount);
            Range range = sheet.UsedRange;
            object[,] usedRange = range.Value;
            int rowCount = range.Rows.Count;
            for (int i = 1; i <= rowCount; ++i)
            {
                try
                {
                    UpdateProgBarConvert(i, rowCount);
                    string currentCell = FixDateTimeFormat(usedRange[i, 1]);
                    if (DateTime.TryParse(currentCell, cInfo, DateTimeStyles.None, out DateTime d))
                    {
                        if (usedRange[i, 1] != null)
                            usedRange[i, 1] = "\'" + currentCell;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
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
                string separatedHour = date.ToString().Substring(11);
                if (separatedHour.Contains("."))
                {
                    string fixedHours = separatedHour.Replace('.', ':');
                    date = date.ToString().Replace(separatedHour, fixedHours);
                }
            }
            //else
            //{
            //    string smalldate = date.ToString().Substring(0, 10);
            //    string fixedSmallDate = "";
            //    if (smalldate.Contains("."))
            //    {
            //        fixedSmallDate = smalldate.Replace('.', '/');
            //        date = date.ToString().Replace(smalldate, fixedSmallDate);
            //        FixDateTimeFormat(date);
            //    }
                
            //}
            
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
            window.progBarConvertText.Dispatcher.Invoke(() => window.progBarConvertText.Text = $"Проверка на датите: {(int)(((decimal)val / max) * 100)}%");
        }

        private void UpdateProgBarConvert(int val)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Value = val);
            window.progBarConvertText.Dispatcher.Invoke(() => window.progBarConvertText.Text = $"Проверка на датите: {(int)((val / (decimal)window.progBarConvert.Dispatcher.Invoke(() => window.progBarConvert.Maximum) * 100))}%");
        }

        private void UpdateProgBarChart(int val, int max, bool print = false)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            if (window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum < max))
            {
                window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum = max);
            }
            if(window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value < val))
                window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value = val);
            if (print)
            {
                window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"Принтиране на графики: {(int)(((decimal)val / max) * 100)}%");
            }
            else
            {
                window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"Създаване на графики: {(int)(((decimal)val / max) * 100)}%");
            }
        }

        private void UpdateProgBarChart(int val, bool print = false)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value = val);
            if (print)
            {
                window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"Принтиране на графики: {(int)((val / (decimal)window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum) * 100))}%");
            }
            else
            {
                window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"Създаване на графики: {(int)((val / (decimal)window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum) * 100))}%");
            }
        }

        private void CreateHumidityGraphs(Worksheet sheet)
        {
            if (CancelRequest())
                return;

            int totalRows = sheet.UsedRange.Rows.Count;
            UpdateProgBarChart(0, totalRows);
            ExcelChart xlChart;
            try
            {
                xlChart = new ExcelChart(sheet, false);
                xlChart.SetChartRange();
                ResetProgBarChart(totalRows);
                Thread.Sleep(1);
            }
            catch (Exception)
            {
                Dispose(sheet);
            }
        }
        private void CreateTemperatureGraphs(Worksheet sheet)
        {
            if (CancelRequest())
                return;

            int totalRows = sheet.UsedRange.Rows.Count;
            UpdateProgBarChart(0, totalRows);
            ExcelChart xlChart;
            try
            {
                xlChart = new ExcelChart(sheet);
                xlChart.SetChartRange();
                ResetProgBarChart(totalRows);
                Thread.Sleep(1);
            }
            catch (Exception)
            {
                Dispose(sheet);
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
            try
            {
                Sheets sheets = xlWBook.Sheets;
                List<Thread> sheetToAlter = new List<Thread>();
                foreach(Worksheet sheet in sheets)
                {
                    Thread t = new Thread(() => ConvertValues(sheet));
                    sheetToAlter.Add(t);
                }
                foreach(Thread t in sheetToAlter)
                {
                    t.Start();
                    t.Join();
                }
            }
            catch (Exception)
            {
                //TODO...
            }
        }

        private void ConvertValues(Worksheet sheet)
        {
            if (sheet.UsedRange.Columns.Count == 2)
            {
                TempOrHumidConvert(sheet.UsedRange.Columns[2]);               
            }
            else if(sheet.UsedRange.Columns.Count == 3)
            {
                TempOrHumidConvert(sheet.UsedRange.Columns[2]);
                TempOrHumidConvert(sheet.UsedRange.Columns[3]);
            }
            else
            {
                MessageBox.Show("Програмата не може да прецени в коя колона са стойностите. Няма направени промени.");
                return;
            }
            
        }
        private void TempOrHumidConvert(Range cells)
        {
            object[,] values = cells.Value;
            double count = 0.0;
            for (int i = 1; i <= (10 < values.Length ? 10 : values.Length); ++i)
            {
                object value = values[i,1];
                if (Double.TryParse(value.ToString(), out Double val))
                {
                    count += val;
                }
            }
            double average = count / (10 < values.Length ? 10 : values.Length);
            if (average < COLD_MAX && average > COLD_MIN)
            {
                for (int i = 1; i <= values.Length; ++i)
                {
                    if (Double.TryParse(values[i, 1].ToString(), out Double d))
                    {
                        if (d > COLD_MAX)
                        {
                            values[i, 1] = (int)d - ((int)d - (int)COLD_MAX) + ((double)d - (int)d);
                        }
                        else if (d < COLD_MIN)
                        {
                            values[i, 1] = (int)d + ((int)COLD_MIN - (int)d) + ((double)d - (int)d);
                        }
                    }
                }
            }
            cells.Value = values;
        }

        /// <summary>
        /// Print all Charts in the Workbook
        /// </summary>
        internal void PrintGraphics()
        {
            try
            {
                Sheets sheets = xlWBook.Worksheets;
                int sheetCount = 1;
                foreach (Worksheet sheet in sheets)
                {
                    if(!(sheet.UsedRange.Rows.Count <= 1))
                        SetLabelNote(sheetCount++, sheets.Count);
                    ChartObjects charts = sheet.ChartObjects();
                    ResetProgBarChart(charts.Count);
                    if(charts != null)
                    {
                        int chartCount = 1;
                        foreach(ChartObject chartObj in charts)
                        {
                            UpdateProgBarChart(chartCount++, charts.Count, true);
                            Chart chart = chartObj.Chart;
                            if(chart != null)
                            {
                                //chart.PrintOut();
                                Thread.Sleep(10);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                Dispose();
            }
        }

        internal void SaveAs()
        {
            try
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls";
                switch (Path.GetExtension(SaveFileName))
                {
                    case ".xls":
                        saveFile.DefaultExt = "xls";
                        break;
                    case ".xlsx":
                        saveFile.DefaultExt = "xlsx";
                        break;
                    default:
                        break;
                }
                saveFile.AddExtension = true;
                saveFile.FileName = SaveFileName;
                saveFile.OverwritePrompt = false;
                bool fileSaved = false;
                while (saveFile.ShowDialog() ?? false)
                {
                    if (File.Exists(saveFile.FileName))
                    {
                        MessageBox.Show("Файлът вече съществува. Моля изберете друго име!");
                        continue;
                    }
                    xlWBook.SaveAs(saveFile.FileName,
                                    Type.Missing,
                                    Type.Missing,
                                    Type.Missing,
                                    false,
                                    false,
                                    XlSaveAsAccessMode.xlExclusive,
                                    false,
                                    false,
                                    Type.Missing,
                                    Type.Missing,
                                    Type.Missing);
                    while (!xlWBook.Saved) { }

                    MessageBox.Show("File saved successfully in \"" + saveFile.FileName + "\"");
                    fileSaved = true;
                    xlWBook.Close(false);
                    xlWBooks.Close();
                    Dispose();

                    break;
                }
                if (!fileSaved)
                {
                    MessageBox.Show("Файлът не беше запазен!");
                }
            }
            catch (COMException comEx)
            {
                if(comEx.Message.Contains("Cannot save as that name."))
                {
                    MessageBox.Show("Файлът е отворен за четене. Нямате права да го променяте. Моля изберете друго име!");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("There was a problem saving the file!");
                xlWBook.Close(false);
                xlWBooks.Close();
                Dispose();
            }
        }
    }
}
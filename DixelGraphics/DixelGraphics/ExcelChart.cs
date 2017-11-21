using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace DixelGraphics
{
    internal class ExcelChart
    {
        const double chartHeigth = 521.0134; //18.23cm * 28.58
        const double chartWidth = 867.9746;  //30.37cm * 28.58
        private readonly bool temperature = true;
        string topDateCell, topValueCell, bottomDateCell, bottomValueCell;
        private readonly char humidValueColumn = 'B';
        Worksheet sheet;
        Range usedRange;
        public int ChartNumber { get; set; } = 1;

        public ExcelChart(Worksheet sheet, bool isTemperature = true)
        {
            temperature = isTemperature;
            this.sheet = sheet;
            usedRange = sheet.UsedRange;
            if (!temperature)
            {
                MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
                if(window.humidColumnCorrectionCheckBox.Dispatcher.Invoke(() => window.humidColumnCorrectionCheckBox.IsChecked ?? false))
                {
                    if(this.usedRange.Columns.Count > 2)
                        humidValueColumn = 'C';
                    else
                    {
                        humidValueColumn = 'B';
                    }
                }
            }
            bottomDateCell = topDateCell = "A1";
            bottomValueCell = topValueCell = humidValueColumn.ToString() + 1;
        }

        public void ExpandRange(int row)
        {
            bottomDateCell = "A" + row;
            bottomValueCell = humidValueColumn.ToString() + row;
        }

        public void StartNewRange(int row)
        {
            bottomDateCell = topDateCell = "A" + row;
            bottomValueCell = topValueCell = humidValueColumn.ToString() + row;
        }

        public void SetChartRange()
        {
            double startPositionLeft = 100;
            double startPositionTop = 100;
            bool startRange = true;
            int totalRows = usedRange.Rows.Count;
            object[,] range = usedRange.Value;
            string currentValue;
            CultureInfo cInfo = new CultureInfo("bg-BG");
            cInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            cInfo.DateTimeFormat.ShortTimePattern = "hh:mm:ss";
            cInfo.DateTimeFormat.DateSeparator = "/";

            for (int i = 1; i <= totalRows; ++i)
            {
                if (CancelRequest())
                {
                    return;
                }
                if (range[i, 1] == null)
                {
                    if(i == totalRows)
                        CreateChart(startPositionLeft, startPositionTop);
                    continue;
                }
                UpdateProgBarChart(i);

                currentValue = Convert.ToString(range[i, 1]).Trim();
                if (currentValue.Contains("\'"))
                    currentValue = currentValue.Remove(currentValue.IndexOf('\''), 1);
                DateTime date;
                if (DateTime.TryParse(currentValue, out date) || DateTime.TryParse(currentValue, cInfo, DateTimeStyles.None, out date))
                {
                    if ((temperature && date.DayOfWeek == DayOfWeek.Monday) || (!temperature && IsFirstDayOfMonth(currentValue, cInfo)))
                    {
                        if (startRange && i != totalRows)
                        {
                            ExpandRange(i);
                        }
                        else
                        {
                            CreateChart(startPositionLeft, startPositionTop);
                            startPositionTop += chartHeigth;
                            StartNewRange(i);
                            startRange = true;
                        }
                    }
                    else
                    {
                        ExpandRange(i);
                        startRange = false;

                        if(i == totalRows)
                        {
                            CreateChart(startPositionLeft, startPositionTop);
                            startPositionTop += chartHeigth;
                            //StartNewRange(i);
                        }
                        else
                        {
                            string nextCell = Convert.ToString(range[i + 1, 1]);
                            if ((temperature && DateTime.TryParse(nextCell, out DateTime d) && d.DayOfWeek == DayOfWeek.Monday) || (!temperature && IsFirstDayOfMonth(currentValue, cInfo)))
                            {
                                CreateChart(startPositionLeft, startPositionTop);
                                startPositionTop += chartHeigth;
                                StartNewRange(i + 1);
                                startRange = true;
                            }
                        }
                    }
                }
                else
                {
                    if (EnoughDataForChart())
                    {
                        CreateChart(startPositionLeft, startPositionTop);
                        startPositionTop += chartHeigth;
                        startRange = true;
                    }
                    StartNewRange(i + 1);
                }
            }
        }

        private bool IsFirstDayOfMonth(string currentValue, CultureInfo cInfo)
        {
            DateTime d;
            if (DateTime.TryParse(currentValue, cInfo, DateTimeStyles.None, out d) && d.Day == 1)
            {
                return true;
            }
            return false;
        }

        private void CreateChart(double startPositionLeft, double startPositionTop)
        {
            try
            {
                ChartObjects charts = sheet.ChartObjects();

                //In case we create both temperature and humdity graphs in one sheet we need to separate them. Otherwise they will be on top of eachother
                if (!temperature)
                {
                    startPositionLeft += 100;
                    startPositionTop += 50;
                }
                string chartTitle = sheet.Name + (temperature ? "_T" : "_H");
                Range DateRange = usedRange.Range[topDateCell, bottomDateCell];
                Range ValueRange = usedRange.Range[topValueCell, bottomValueCell];
                ChartObject chart = charts.Add(startPositionLeft, startPositionTop, chartWidth, chartHeigth);

                Chart xlChartPage = chart.Chart;
                Series xlChartSeries = xlChartPage.SeriesCollection().Add(ValueRange);
                xlChartSeries.XValues = DateRange;
                xlChartPage.ChartType = XlChartType.xlLine;
                xlChartPage.HasTitle = true;
                xlChartPage.ChartTitle.Text = chartTitle;
                xlChartPage.Legend.Delete();

                DateRange = null;
                ValueRange = null;
                Marshal.ReleaseComObject(chart);
                chart = null;
            }
            catch (Exception)
            {

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

        private void UpdateProgBarChart(int val, int max)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            if (window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum < max))
            {
                window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum = max);
            }
            if (window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value < val))
                window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value = val);
            window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"Създаване на графики: {(int)((val / max) * 100)}%");
        }

        private void UpdateProgBarChart(int val)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            if(window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value < val))
                window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value = val);
            window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"Създаване на графики: {(int)((val / window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum) * 100))} %");
        }

        private bool EnoughDataForChart()
        {
            if (topDateCell == bottomDateCell || topValueCell == bottomValueCell)
            {
                return false;
            }
            return true;
        }
    }
}
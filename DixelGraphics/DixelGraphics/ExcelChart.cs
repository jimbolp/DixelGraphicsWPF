using System;
using System.Runtime.InteropServices;
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
            int startPositionLeft = 100;
            int startPositionTop = 100;
            bool startRange = true;
            int totalRows = usedRange.Rows.Count;
            object[,] range = usedRange.Value;
            string currentValue;
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
                if (DateTime.TryParse(currentValue, out DateTime date))
                {
                    if (date.DayOfWeek == DayOfWeek.Monday)
                    {
                        if (startRange && i != totalRows)
                        {
                            ExpandRange(i);
                        }
                        else
                        {
                            CreateChart(startPositionLeft, startPositionTop);
                            startPositionTop += 600;
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
                            startPositionTop += 600;
                            //StartNewRange(i);
                        }
                        else
                        {
                            string nextCell = Convert.ToString(range[i + 1, 1]);
                            if (DateTime.TryParse(nextCell, out DateTime d) && d.DayOfWeek == DayOfWeek.Monday)
                            {
                                CreateChart(startPositionLeft, startPositionTop);
                                startPositionTop += 600;
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
                        startPositionTop += 600;
                        startRange = true;
                    }
                    StartNewRange(i + 1);
                }
            }
        }

        private void CreateChart(int startPositionLeft, int startPositionTop)
        {
            try
            {
                ChartObjects charts = sheet.ChartObjects();
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
            window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"{val}/{max}");
        }

        private void UpdateProgBarChart(int val)
        {
            MainWindow window = System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow as MainWindow);
            if(window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value < val))
                window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Value = val);
            window.progBarChartText.Dispatcher.Invoke(() => window.progBarChartText.Text = $"{val}/{window.progBarChart.Dispatcher.Invoke(() => window.progBarChart.Maximum)}");
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
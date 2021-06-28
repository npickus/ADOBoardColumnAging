using ExcelController.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelController
{
    public class StoryPointCycleTimeReportBuilder : IDisposable
    {
        private ExcelHelper _excelController;
        public string TargetFile { get; set; }

        public StoryPointCycleTimeReportBuilder(string filename)
        {
            TargetFile = filename;
            _excelController = new ExcelHelper();
            _excelController.CreateOpenExcelFile(TargetFile);
        }

        public void BuildCycleTimeReport(List<CycleTimeData> ctData)
        {
            try
            {
                var sheet = _excelController.GetWorkSheet("Sheet1");


                int currRow = 1;

                SetPageHeaders(currRow, sheet);
                _excelController.SaveWorkBook();
                currRow++;

                foreach (var wi in ctData)
                {
                    if (!wi.ClosedDate.Equals("9999-01-01T00:00:00Z"))
                    {
                        ((Range)sheet.Cells[currRow, 1]).Value2 = wi.WorkItemId;
                        ((Range)sheet.Cells[currRow, 2]).Value2 = wi.WorkItemType;
                        ((Range)sheet.Cells[currRow, 3]).Value2 = wi.StoryPoints;
                        ((Range)sheet.Cells[currRow, 4]).Value2 = DateTime.Parse(wi.ActiveDate, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal).ToString();
                        ((Range)sheet.Cells[currRow, 5]).Value2 = DateTime.Parse(wi.ClosedDate, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal).ToString();
                        string formula = $"=(INT(E{currRow}-D{currRow})*24)+ HOUR(E{currRow}-D{currRow})";
                        ((Range)sheet.Cells[currRow, 6]).Formula = formula;
                        currRow++;
                    }
                }


                CreateChartTest(sheet, currRow);
                //CreateChart(sheet, currRow);

                _excelController.SaveWorkBook();
                _excelController.DisposeExcel();
            }
            catch (Exception exception)
            {

                throw exception;
            }
        }

        private static void CreateChartTest(Worksheet sheet, int currRow)
        {
            try
            {
                object misValue = System.Reflection.Missing.Value;
                var missing = Type.Missing;
                ChartObjects xlCharts = (ChartObjects)sheet.ChartObjects(missing);
                ChartObject myChart = (ChartObject)xlCharts.Add(400, 80, 300, 250);
                Chart chartPage = myChart.Chart;
                chartPage.ChartWizard($"Sheet1!$F$1:$F${currRow}", XlChartType.xlXYScatter, missing, missing, missing, missing, missing, missing, "Story Points", missing, missing);
                Series series = (Series)chartPage.SeriesCollection(1);
                series.Values = sheet.get_Range("F2", $"F{currRow}");
                series.XValues = sheet.get_Range("C2", $"C{currRow}"); //sheet.Range[$"$C$1:$C${currRow}"];
                chartPage.Refresh();
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private static void CreateChart(Worksheet sheet, int currRow)
        {
            object misValue = System.Reflection.Missing.Value;
            ChartObjects xlCharts = (ChartObjects)sheet.ChartObjects(Type.Missing);
            ChartObject myChart = (ChartObject)xlCharts.Add(100, 80, 500, 250);
            Chart chartPage = myChart.Chart;
            //=Sheet1!$C$1:$C$498,Sheet1!$F$1:$F$498
            var chartRange = sheet.get_Range($"Sheet1!$C$1:$C${currRow},Sheet1!$F$1:$F${currRow}");
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = XlChartType.xlXYScatter;
        }

        private void SetPageHeaders(int headerRow, Worksheet sheet)
        {
            ((Range)sheet.Cells[headerRow, 1]).Value2 = "Work Item Id";
            ((Range)sheet.Cells[headerRow, 2]).Value2 = "Work Item Type";
            ((Range)sheet.Cells[headerRow, 3]).Value2 = "Story Points";
            ((Range)sheet.Cells[headerRow, 4]).Value2 = "Active Date";
            ((Range)sheet.Cells[headerRow, 5]).Value2 = "Done Date";
            ((Range)sheet.Cells[headerRow, 6]).Value2 = "Difference In Hours";
        }



        #region IDisposable Support

        private bool disposedValue; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (_excelController != null)
                    {
                        _excelController.DisposeExcel();
                        _excelController = null;
                    }
                }

                disposedValue = true;
            }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}

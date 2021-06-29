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
    public class ColumnAgingReportBuilder : IDisposable
    {
        private ExcelHelper _excelController;
        public string TargetFile { get; set; }

        public ColumnAgingReportBuilder(string filename)
        {
            TargetFile = filename;
            _excelController = new ExcelHelper();
            _excelController.CreateOpenExcelFile(TargetFile);
        }

        public void BuildCycleTimeReport(List<ColumnAgingData> ctData)
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

                    /*
                     * Id, type, Title, area path, Url, Current Column, Column entry, Column Exit, Age
                     */
                    //if (!wi.ColumnExitDate.Equals("9999-01-01T00:00:00Z"))

                    ((Range)sheet.Cells[currRow, 1]).Value2 = wi.WorkItemId;
                    ((Range)sheet.Cells[currRow, 2]).Value2 = wi.WorkItemType;
                    ((Range)sheet.Cells[currRow, 3]).Value2 = wi.Title;
                    ((Range)sheet.Cells[currRow, 4]).Value2 = wi.AreaPath;
                    ((Range)sheet.Cells[currRow, 5]).Value2 = wi.Url;
                    ((Range)sheet.Cells[currRow, 6]).Value2 = wi.CurrentColumn;
                    ((Range)sheet.Cells[currRow, 7]).Value2 = DateTime.Parse(wi.ColumnEntryDate, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal).ToString();

                    if (!string.IsNullOrWhiteSpace(wi.ColumnExitDate))
                    {
                        ((Range)sheet.Cells[currRow, 8]).Value2 = DateTime.Parse(wi.ColumnExitDate, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal).ToString();
                    }
                    else
                    {
                        ((Range)sheet.Cells[currRow, 8]).Value2 = DateTime.UtcNow.ToString();
                    }
                    
                    string formula = $"=(H{currRow}-G{currRow})";
                    ((Range)sheet.Cells[currRow, 9]).Formula = formula;
                    currRow++;
                }


                //CreateChartTest(sheet, currRow);
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
            /*
            * Id, type, Title, Url, Current Column, Column entry, Column Exit, Age
            */
            ((Range)sheet.Cells[headerRow, 1]).Value2 = "Work Item Id";
            ((Range)sheet.Cells[headerRow, 2]).Value2 = "Work Item Type";
            ((Range)sheet.Cells[headerRow, 3]).Value2 = "Title";
            ((Range)sheet.Cells[headerRow, 4]).Value2 = "Area Path";
            ((Range)sheet.Cells[headerRow, 5]).Value2 = "URL";
            ((Range)sheet.Cells[headerRow, 6]).Value2 = "BoardColumn";
            ((Range)sheet.Cells[headerRow, 7]).Value2 = "Column Entry";
            ((Range)sheet.Cells[headerRow, 8]).Value2 = "Column Exit";
            ((Range)sheet.Cells[headerRow, 9]).Value2 = "Age";
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

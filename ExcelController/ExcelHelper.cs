using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelController
{
    public class ExcelHelper
    {
        private readonly Excel.Application _excelApp;
        private Excel.Workbook _workBook;

        public ExcelHelper()
        {

            _excelApp = ExcelHelper.GetExcel();
        }

        public Excel.Application GetExcelApplication()
        {
            return _excelApp;
        }

        public void CreateOpenExcelFile(string excelFile)
        {
            _excelApp.Visible = true;
            if (File.Exists(excelFile))
            {
                _workBook = _excelApp.Workbooks.Open(excelFile);
            }
            else
            {
                _workBook = _excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                _workBook.SaveAs(excelFile);
            }
        }

        public Excel.Worksheet CreateWorksheet(string name)
        {
            var sheet = (Excel.Worksheet)_workBook.Sheets.Add();
            sheet.Name = name;
            return sheet;
        }

        public Excel.Worksheet GetWorkSheet(string name)
        {
            var sheet = (Excel.Worksheet)_workBook.Worksheets.Item[name];
            return sheet;
        }

        public void SaveWorkBook()
        {
            _workBook.Save();
        }

        public static Excel.Application GetExcel()
        {
            return new Excel.Application();
            //Excel.Application xl;
            //try
            //{
            //    xl = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            //}
            //catch (Exception)
            //{
            //    xl = new Excel.Application();
            //}

            //if (xl == null) throw new ApplicationException("Couldn't Open Excel");
            //return xl;
        }

        public void DisposeExcel()
        {
            // Cleanup 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            _workBook.Close(Type.Missing, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(_workBook);

            _excelApp.Quit();
            Marshal.FinalReleaseComObject(_excelApp);
        }
    }
}
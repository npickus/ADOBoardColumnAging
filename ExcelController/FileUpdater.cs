using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelController
{
    public class FileUpdater : IDisposable
    {

        private ExcelHelper _sourceExcelController;
        public string SourceFile { get; set; }

        public FileUpdater(string sourceFile)
        {
            SourceFile = sourceFile;
            _sourceExcelController = new ExcelHelper();
            _sourceExcelController.CreateOpenExcelFile(sourceFile);
        }

        public void RemoveWorkItemsFromExcelIfNotInList(List<int> workItemIds, string worksheetName, string outputFileName)
        {
            var sheet = _sourceExcelController.GetWorkSheet(worksheetName);
            var currRow = 2;
            while (((Range)sheet.Cells[currRow, 1]).Value2 != null)
            {
                if (!workItemIds.Contains((int)((Range)sheet.Cells[currRow, 1]).Value2))
                {
                    ((Range)sheet.Rows[currRow]).Delete(Type.Missing);
                }
                else
                {
                    currRow++;
                }
            }
            sheet.SaveAs(outputFileName);

        }

        #region IDisposable Support

        private bool disposedValue; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (_sourceExcelController != null)
                    {
                        _sourceExcelController.DisposeExcel();
                        _sourceExcelController = null;
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

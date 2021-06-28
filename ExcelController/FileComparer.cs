using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelController
{


    public class FileComparer : IDisposable
    {
        private ExcelHelper _sourceExcelController;
        private ExcelHelper _targetExcelController;
        public string SourceFile { get; set; }
        public string TargetFile { get; set; }
        public FileComparer(string sourceFile, string targetFile)
        {
            SourceFile = sourceFile;
            TargetFile = targetFile;
            _sourceExcelController = new ExcelHelper();
            _sourceExcelController.CreateOpenExcelFile(sourceFile);
            _targetExcelController = new ExcelHelper();
            _targetExcelController.CreateOpenExcelFile(targetFile);

        }

        public List<int> FindMissingSourceIdsFromTarget()
        {
            //Source = Changed Items
            //Target = Original Items
            var sourceIds = ParseSourceWorkItemIds("Features");
            var targetIds = ParseTargetWorkItemIds("Features");

            //return all ids in source that aren't in target
            return sourceIds.Except(targetIds).ToList();
            //return new List<int>();
        }

        public List<int> FindMissingTargetIdsInSource()
        {
            //Source = Changed Items
            //Target = Original Items
            var sourceIds = ParseSourceWorkItemIds("Features");
            var targetIds = ParseTargetWorkItemIds("Features");

            //return all ids in source that aren't in target
            return targetIds.Except(sourceIds).ToList();
            //return new List<int>();
        }

        public List<int> FindIntersectionOfIds()
        {
            //Source = Changed Items
            //Target = Original Items
            var sourceIds = ParseSourceWorkItemIds("Features");
            var targetIds = ParseTargetWorkItemIds("Features");

            //return all ids in source that aren't in target
            return sourceIds.Intersect(targetIds).ToList();
            //return new List<int>();
        }


        private List<int> ParseSourceWorkItemIds(string workSheetName, int startingrow = 2, int startingcol = 1)
        {
            var sheet = _sourceExcelController.GetWorkSheet(workSheetName);
            var currRow = startingrow;
            var ids = new List<int>();
            while (((Range)sheet.Cells[currRow, 1]).Value2 != null)
            {
                ids.Add((int)((Range)sheet.Cells[currRow, 1]).Value2);
                currRow++;
            }
            return ids;
        }

        private List<int> ParseTargetWorkItemIds(string workSheetName, int startingrow = 2, int startingcol = 1)
        {
            var sheet = _targetExcelController.GetWorkSheet(workSheetName);
            var currRow = startingrow;
            var ids = new List<int>();
            while (((Range)sheet.Cells[currRow, 1]).Value2 != null)
            {
                ids.Add((int)((Range)sheet.Cells[currRow, 1]).Value2);
                currRow++;
            }
            return ids;
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

                    if (_targetExcelController != null)
                    {
                        _targetExcelController.DisposeExcel();
                        _targetExcelController = null;
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

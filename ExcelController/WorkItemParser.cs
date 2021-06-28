using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelController
{
    public class WorkItemParser : IDisposable
    {
        private ExcelHelper _excelController;

        public WorkItemParser(string filename)
        {
            FileName = filename;
            _excelController = new ExcelHelper();
            _excelController.CreateOpenExcelFile(filename);
        }

        public string FileName { get; set; }

        public Dictionary<int, List<int>> ParseWorkItems(string worksheetname, Dictionary<int, List<int>> workItemData,
            int startingrow = 2, int startingcol = 1)
        {
            var sheet = _excelController.GetWorkSheet(worksheetname);
            var currRow = startingrow;

            while (((Range)sheet.Cells[currRow, 1]).Value2 != null)
            {
            }

            return new Dictionary<int, List<int>>();
        }

        public Dictionary<int, DateTime> ParseCreatedDateData(string worksheetname,
            Dictionary<int, DateTime> workItemData, int startingrow = 2, int startingcol = 1)
        {
            var columnData = GetBugDataColumnInfo();
            var sheet = _excelController.GetWorkSheet(worksheetname);
            var currRow = startingrow;

            while (((Range)sheet.Cells[currRow, 1]).Value2 != null)
            {
                var workItemId = (int)((Range)sheet.Cells[currRow, columnData["ID"]]).Value2;

                DateTime createdDate =
                    DateTime.FromOADate(((Range)sheet.Cells[currRow, columnData["Created Date"]]).Value2);
                //DateTime createdDate = DateTime.Parse(;
                if (!workItemData.ContainsKey(workItemId)) workItemData.Add(workItemId, createdDate);
                currRow++;
            }

            return workItemData;
        }

        public Dictionary<int, List<int>> ParseTestData(string worksheetname, Dictionary<int, List<int>> workItemData,
            int startingrow = 2, int startingcol = 1)
        {
            var columnData = GetTestDataColumnInfo();
            var sheet = _excelController.GetWorkSheet(worksheetname);
            var currRow = startingrow;

            while (((Range)sheet.Cells[currRow, 1]).Value2 != null)
            {
                string entityType = ((Range)sheet.Cells[currRow, columnData["Entity Type"]]).Value2;
                var workItemId = (int)((Range)sheet.Cells[currRow, columnData["ID"]]).Value2;
                //Determine Entity Type
                if (entityType == "TestCase")
                    if (!workItemData.ContainsKey(workItemId))
                        workItemData.Add(workItemId, new List<int>());

                //Add to correct parent
                if (entityType == "TestStep")
                {
                    var testCaseId = (int)((Range)sheet.Cells[currRow, columnData["Test Case ID"]]).Value2;
                    workItemData[testCaseId].Add(workItemId);
                }


                currRow++;
            }

            return workItemData;
        }

        private static IDictionary<string, string> GetBugDataColumnInfo()
        {
            var bugDataColumns = new Dictionary<string, string>
            {
                {"ID", "A"},
                {"Work Item Type", "B"},
                {"Title", "C"},
                {"Created Date", "D"},
                {"Modified Date", "E"}
            };
            return bugDataColumns;
        }

        private static IDictionary<string, string> GetTestDataColumnInfo()
        {
            var testDataColumns = new Dictionary<string, string>
            {
                {"ProjectID", "A"},
                {"Entity Type", "B"},
                {"Test Case ID", "C"},
                {"ID", "D"},
                {"Name", "E"},
                {"Description", "F"},
                {"Result", "G"},
                {"Start Date", "H"},
                {"End Date", "I"},
                {"Create Date", "J"},
                {"Modify Date", "K"},
                {"Last Comment Date", "L"},
                {"Last Run Date", "M"},
                {"Project", "N"},
                {"User Story", "O"},
                {"Last Commented User", "P"},
                {"Tags", "Q"},
                {"Numeric Priority", "R"},
                {"Last Editor", "S"},
                {"Owner", "T"},
                {"Linked Test Plan", "U"},
                {"Last Status", "V"},
                {"Last Run Status", "W"},
                {"Last Failure Comment", "X"},
                {"Priority", "Y"},
                {"Test Case", "Z"},
                {"PL Auto Dev Status", "AA"},
                {"PL Auto Dev Iteration", "AB"},
                {"Device Type", "AC"},
                {"PL AutoDeveloper", "AD"},
                {"Validation Group", "AE"},
                {"VL Auto Dev Iteration", "AF"},
                {"VL Auto Dev Status", "AG"},
                {"VL AutoDeveloper", "AH"},
                {"Test in PLorVL", "AI"},
                {"AutoMethod", "AJ"},
                {"VL UF Updated", "AK"},
                {"RevampTCID", "AL"}
            };

            return testDataColumns;
        }

        #region IDisposable Support

        private bool disposedValue; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                    if (_excelController != null)
                    {
                        _excelController.DisposeExcel();
                        _excelController = null;
                    }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~WorkItemParsert() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
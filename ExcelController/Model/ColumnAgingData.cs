using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelController.Model
{
    public class ColumnAgingData
    {
        public int WorkItemId { get; set; }
        public string WorkItemType { get; set; }
        public string Title { get; set; }
        public string Url { get; set; }
        public string CurrentColumn { get; set; }
        public string ColumnEntryDate { get; set; }
        public string ColumnExitDate { get; set; }
    }
}

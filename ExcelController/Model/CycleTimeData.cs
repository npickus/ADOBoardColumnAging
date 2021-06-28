using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelController.Model
{
    public class CycleTimeData
    {
        public int WorkItemId { get; set; }
        public string WorkItemType { get; set; }
        public decimal StoryPoints { get; set; }
        public string ActiveDate { get; set; }
        public string ClosedDate { get; set; }
    }
}

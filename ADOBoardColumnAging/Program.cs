using AzDOController;
using AzDOController.JsonData;
using ExcelController;
using ExcelController.Model;
using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace ADOBoardColumnAging
{
    static class Settings
    {
        public const bool UPDATEDATA = true;
        public const bool REPORTDATA = false;
        public const bool FIELDDATA = false;
        public const string BASE = "https://dev.azure.com";
        public const string AUTHUSER = "nick.m.pickus@hotmail.com";
        public const string PAT = "5s5qzsg7knp5b4i7oapibxwt6bokhkqatcwstm3b3gdcykqdfd5q";
        public const string API = "api-version=6.0";
        public const string ORG = "nickp";
        public const string PROJ = "testASD";
        public const string EXCELFILE = @"C:\ADOExport\ADOAging";
    }
    class Program
    {
        static void Main(string[] args)
        {
            XmlConfigurator.Configure();
            var logger = LogManager.GetLogger(typeof(Program));
            try
            {
                var startTime = DateTime.Now;
                var processedItemsDictionary = new Dictionary<int, int>();
                if (logger.IsInfoEnabled)
                    logger.InfoFormat("{0}: BEGIN ", DateTime.Now.ToShortTimeString());

                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.ASCII.GetBytes(string.Format("{0}:{1}", Settings.AUTHUSER, Settings.PAT))));

                var azdoController = new Controller(Settings.BASE, Settings.PAT, Settings.AUTHUSER, logger);
                var results = GetWorkItemListByQuery(azdoController);
                CreateExcelFile(results, $"{Settings.EXCELFILE} {DateTime.Now.Ticks}.xlsx", azdoController, logger);

            }
            catch (Exception exception)
            {
                logger.ErrorFormat(
                           $"Unable to complete Data Collection and Reporting {exception.Message} - {exception.StackTrace}");
            }
        }


        private static QueryResponseJson GetWorkItemListByQuery(Controller azdoController)
        {
            //WIQL Reference: https://docs.microsoft.com/en-us/azure/devops/boards/queries/wiql-syntax?view=azure-devops

            string wiql = @"{""query"": ""SELECT
    [System.Id],
    [System.WorkItemType],
    [System.Title],
    [System.AssignedTo],
    [System.State],
    [System.Tags],
    [System.AreaPath],
    [System.BoardColumn]
FROM workitems
WHERE
    [System.TeamProject] = @project
    AND [System.WorkItemType] = 'Product Backlog Item'
    AND NOT [System.State] IN ('Closed', 'Removed')""}";
            var results = azdoController.GetAzDOQueryResults(wiql, Settings.ORG, Settings.PROJ);
            return results;
        }

        private static void CreateExcelFile(QueryResponseJson results, string excelfile, Controller azdoController, ILog logger)
        {
            try
            {
                var items = GetAgingDataList(results, azdoController, logger);
                //var items = GetTestCycleTimeData();
                using (var excel = new ColumnAgingReportBuilder(excelfile))
                {
                    excel.BuildCycleTimeReport(items);
                }
            }
            catch (Exception exception)
            {
                logger.ErrorFormat(
                           $"Unable to CreateExcelFile {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        private static List<ColumnAgingData> GetAgingDataList(QueryResponseJson results, Controller azdoController, ILog logger)
        {
            try
            {
                var items = new List<ColumnAgingData>();

                for (int i = 0; i < results.workItems.Length; i++)
                {
                    var wiInfo = results.workItems[i];
                    var workItem = azdoController.GetWorkItemAgingData(Settings.ORG, Settings.PROJ, wiInfo.id);
                    if (!string.IsNullOrEmpty(workItem.fields.ColumnEntryDate))
                    {
                        var cad = new ColumnAgingData
                        {
                            WorkItemId = workItem.id,
                            Url = workItem.url,
                            WorkItemType = workItem.fields.WorkItemType,
                            Title = workItem.fields.Title,
                            CurrentColumn = (!string.IsNullOrEmpty(workItem.fields.BoardColumn)) ? workItem.fields.BoardColumn : "New",
                            ColumnEntryDate = (!string.IsNullOrEmpty(workItem.fields.ColumnEntryDate)) ? workItem.fields.ColumnEntryDate : workItem.fields.ChangedDate
                        };

                        items.Add(cad);
                    }
                }
                
                return items;
            }
            catch (Exception exception)
            {
                logger.ErrorFormat(
                           $"Unable to GetCycleTimeDataList {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        private static List<CycleTimeData> GetCycleTimeDataList(QueryResponseJson results, Controller azdoController, ILog logger)
        {
            try
            {
                var items = new List<CycleTimeData>();
                foreach (var wiInfo in results.workItems)
                {
                    var workItem = azdoController.GetWorkItemCycleTimeData(Settings.ORG, Settings.PROJ, wiInfo.id);
                    if (!string.IsNullOrEmpty(workItem.fields.ActiveDate))
                    {
                        var ct = new CycleTimeData
                        {
                            WorkItemId = workItem.id,
                            WorkItemType = workItem.fields.WorkItemType,
                            StoryPoints = workItem.fields.StoryPoints,
                            ActiveDate = workItem.fields.ActiveDate,
                            //ClosedDate = (!string.IsNullOrEmpty(workItem.fields.DoneDate)) ? workItem.fields.DoneDate : workItem.fields.ResolvedDate,
                            ClosedDate = (!string.IsNullOrEmpty(workItem.fields.DoneDate)) ? workItem.fields.DoneDate : (!string.IsNullOrEmpty(workItem.fields.ClosedDate)) ? workItem.fields.ClosedDate : workItem.fields.ResolvedDate,
                        };
                        items.Add(ct);
                    }

                }
                return items;
            }
            catch (Exception exception)
            {
                logger.ErrorFormat(
                           $"Unable to GetCycleTimeDataList {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }
        private static List<CycleTimeData> GetTestCycleTimeData()
        {
            var testData = new List<CycleTimeData>
            {
                new CycleTimeData {
                    WorkItemId = 238008, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "2/21/2019 18:04", ClosedDate = "3/18/2019 15:40",},
new CycleTimeData {WorkItemId = 241502, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "2/26/2019 22:49", ClosedDate = "3/18/2019 15:32",},
new CycleTimeData {WorkItemId = 258837, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "3/18/2019 15:16", ClosedDate = "3/18/2019 23:11",},
new CycleTimeData {WorkItemId = 258839, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "3/27/2019 20:56", ClosedDate = "8/20/2019 7:09",},
new CycleTimeData {WorkItemId = 644514, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "11/10/2020 16:36", ClosedDate = "1/5/2021 12:53",},
new CycleTimeData {WorkItemId = 664360, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "10/5/2020 12:28", ClosedDate = "10/15/2020 18:10",},
new CycleTimeData {WorkItemId = 772687, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "5/26/2020 11:50", ClosedDate = "10/9/2020 13:26",},
new CycleTimeData {WorkItemId = 801836, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "8/20/2020 12:49", ClosedDate = "9/29/2020 16:17",},
new CycleTimeData {WorkItemId = 802601, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "7/6/2020 5:37", ClosedDate = "10/15/2020 20:36",},
new CycleTimeData {WorkItemId = 811432, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "8/3/2020 13:01", ClosedDate = "11/11/2020 14:58",},
new CycleTimeData {WorkItemId = 811579, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "7/13/2020 14:36", ClosedDate = "11/11/2020 14:58",},
new CycleTimeData {WorkItemId = 855176, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "12/15/2020 6:13", ClosedDate = "1/6/2021 14:08",},
new CycleTimeData {WorkItemId = 864064, WorkItemType ="User Story", StoryPoints =13, ActiveDate = "7/22/2020 12:03", ClosedDate = "10/23/2020 13:02",},
new CycleTimeData {WorkItemId = 864595, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/15/2020 13:22", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 887013, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "12/15/2020 14:03", ClosedDate = "1/5/2021 12:53",},
new CycleTimeData {WorkItemId = 892164, WorkItemType ="User Story", StoryPoints =13, ActiveDate = "9/29/2020 18:34", ClosedDate = "10/15/2020 18:10",},
new CycleTimeData {WorkItemId = 905699, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "9/30/2020 14:22", ClosedDate = "10/6/2020 14:24",},
new CycleTimeData {WorkItemId = 911447, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "8/19/2020 18:15", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 919761, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "9/8/2020 12:36", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 920592, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "9/2/2020 13:47", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 921231, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "9/9/2020 18:44", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 927733, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "9/15/2020 14:29", ClosedDate = "11/11/2020 14:58",},
new CycleTimeData {WorkItemId = 934212, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/1/2020 16:14", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 934637, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "9/23/2020 14:51", ClosedDate = "10/15/2020 18:10",},
new CycleTimeData {WorkItemId = 936466, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/22/2020 14:36", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 941871, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "10/23/2020 15:37", ClosedDate = "11/12/2020 17:02",},
new CycleTimeData {WorkItemId = 941877, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/23/2020 15:39", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 941997, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "11/9/2020 13:39", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 942171, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "9/17/2020 17:42", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 942204, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "9/17/2020 17:42", ClosedDate = "9/30/2020 20:12",},
new CycleTimeData {WorkItemId = 947660, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/5/2020 19:46", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 947702, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "11/6/2020 15:57", ClosedDate = "1/5/2021 12:53",},
new CycleTimeData {WorkItemId = 947904, WorkItemType ="User Story", StoryPoints =1, ActiveDate = "10/20/2020 15:36", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 947917, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "11/5/2020 13:43", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 948136, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "9/16/2020 14:00", ClosedDate = "11/11/2020 14:58",},
new CycleTimeData {WorkItemId = 948140, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "9/28/2020 13:40", ClosedDate = "11/11/2020 14:58",},
new CycleTimeData {WorkItemId = 948155, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "9/18/2020 22:41", ClosedDate = "11/11/2020 14:58",},
new CycleTimeData {WorkItemId = 948343, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "11/4/2020 10:13", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 950311, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/15/2020 13:15", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 960037, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/19/2020 15:35", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 960038, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "9/18/2020 18:24", ClosedDate = "9/29/2020 16:17",},
new CycleTimeData {WorkItemId = 961890, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "2/10/2021 16:33", ClosedDate = "2/25/2021 13:59",},
new CycleTimeData {WorkItemId = 961986, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "11/24/2020 13:49", ClosedDate = "12/9/2020 12:44",},
new CycleTimeData {WorkItemId = 962008, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "11/6/2020 16:37", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 962009, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "10/30/2020 15:34", ClosedDate = "11/17/2020 17:01",},
new CycleTimeData {WorkItemId = 962010, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "11/5/2020 14:01", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 962030, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/22/2020 15:31", ClosedDate = "10/30/2020 13:02",},
new CycleTimeData {WorkItemId = 963465, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/13/2020 12:52", ClosedDate = "10/15/2020 18:10",},
new CycleTimeData {WorkItemId = 963589, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/7/2020 11:06", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 963597, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/15/2020 13:11", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 963655, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "10/5/2020 15:36", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 963663, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "10/1/2020 16:10", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 963698, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "10/29/2020 12:50", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 963727, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "2/10/2021 16:47", ClosedDate = "2/25/2021 13:59",},
new CycleTimeData {WorkItemId = 963739, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "12/14/2020 16:36", ClosedDate = "1/5/2021 12:53",},
new CycleTimeData {WorkItemId = 963740, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "11/9/2020 12:58", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 963743, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "11/6/2020 16:34", ClosedDate = "1/5/2021 12:53",},
new CycleTimeData {WorkItemId = 965162, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "9/22/2020 13:28", ClosedDate = "10/15/2020 20:36",},
new CycleTimeData {WorkItemId = 965201, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "9/22/2020 13:28", ClosedDate = "10/14/2020 10:56",},
new CycleTimeData {WorkItemId = 965288, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "2/17/2021 14:30", ClosedDate = "5/11/2021 16:53",},
new CycleTimeData {WorkItemId = 965674, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/13/2020 12:21", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 965733, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/27/2020 14:45", ClosedDate = "11/11/2020 14:58",},
new CycleTimeData {WorkItemId = 965736, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/15/2020 13:12", ClosedDate = "11/29/2020 16:16",},
new CycleTimeData {WorkItemId = 965757, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/1/2020 14:14", ClosedDate = "10/9/2020 13:26",},
new CycleTimeData {WorkItemId = 965817, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/14/2020 12:34", ClosedDate = "10/15/2020 18:10",},
new CycleTimeData {WorkItemId = 965818, WorkItemType ="User Story", StoryPoints =13, ActiveDate = "10/16/2020 12:53", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 965825, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "9/30/2020 9:20", ClosedDate = "10/15/2020 18:10",},
new CycleTimeData {WorkItemId = 965949, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "9/24/2020 11:27", ClosedDate = "9/29/2020 16:17",},
new CycleTimeData {WorkItemId = 965950, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "9/23/2020 13:37", ClosedDate = "9/24/2020 11:59",},
new CycleTimeData {WorkItemId = 968973, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "9/30/2020 20:12", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 968978, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/13/2020 15:42", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 968985, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "10/1/2020 16:00", ClosedDate = "11/29/2020 16:19",},
new CycleTimeData {WorkItemId = 971735, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/13/2020 12:04", ClosedDate = "10/15/2020 18:10",},
new CycleTimeData {WorkItemId = 974249, WorkItemType ="User Story", StoryPoints =5, ActiveDate = "10/6/2020 14:54", ClosedDate = "10/15/2020 20:36",},
new CycleTimeData {WorkItemId = 974250, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "9/30/2020 15:01", ClosedDate = "10/15/2020 20:36",},
new CycleTimeData {WorkItemId = 974251, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "10/7/2020 14:09", ClosedDate = "11/29/2020 16:16",},
new CycleTimeData {WorkItemId = 974258, WorkItemType ="User Story", StoryPoints =2, ActiveDate = "10/5/2020 14:10", ClosedDate = "11/11/2020 14:58",},
new CycleTimeData {WorkItemId = 977481, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "10/19/2020 12:35", ClosedDate = "11/2/2020 13:43",},
new CycleTimeData {WorkItemId = 977986, WorkItemType ="User Story", StoryPoints =8, ActiveDate = "10/14/2020 14:29", ClosedDate = "11/29/2020 16:16",},
new CycleTimeData {WorkItemId = 978208, WorkItemType ="User Story", StoryPoints =3, ActiveDate = "9/30/2020 20:12", ClosedDate = "10/15/2020 18:10",},
            };
            return testData;
        }
    }
}


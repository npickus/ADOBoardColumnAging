using AzDOController.JsonData;
using log4net;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Linq;

namespace AzDOController
{
    public class Controller
    {
        private readonly ILog _logger;
        private readonly HttpClient _client;
        private readonly string _baseUrl;
        //private string 
        #region Constructors
        public Controller()
        {
            throw new NotImplementedException();
        }

        public Controller(string baseAzDOUrl, HttpClient client, ILog logger)
        {
            _logger = logger;
            _client = client;
            _baseUrl = baseAzDOUrl;
        }

        public Controller(string baseAzDOUrl, string PAT, string AuthenticationUser, ILog logger)
        {
            _baseUrl = baseAzDOUrl;
            _logger = logger;
            _client = CreateHttpClient(PAT, AuthenticationUser);

        }

        private HttpClient CreateHttpClient(string PAT, string AuthenticationUser)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.ASCII.GetBytes(string.Format("{0}:{1}", AuthenticationUser, PAT))));
            return client;
        }
        #endregion



        #region GetAzDO Data

        public WorkItemJson GetWorkItemCycleTimeData(string orgName, string projectName, int workItemId)
        {
            try
            {
                var workItem = GetWorkItemData(orgName, projectName, workItemId);
                //https://dev.azure.com/{organization}/{project}/_apis/wit/workItems/{id}/updates?api-version=6.0
                string uri = String.Join("?", String.Join("/", _baseUrl, orgName, projectName, "_apis/wit/workitems", workItemId, "updates"), "&api-version=6.0");
                var result = SendRequest(uri).Result;
                var wi = JsonConvert.DeserializeObject<WorkItemUpdates>(result);
                foreach (var update in wi.value)
                {
                    if (update.fields != null && update.fields.State != null)
                    {
                        if (update.fields.State.newValue == "Active")
                        {
                            workItem.fields.ActiveDate = update.fields.RevisedDate.newValue;
                        }

                        if (update.fields.State.newValue == "Closed")
                        {
                            workItem.fields.ClosedDate = update.fields.RevisedDate.newValue;
                        }

                        if (update.fields.State.newValue == "Resolved")
                        {
                            workItem.fields.ResolvedDate = update.fields.RevisedDate.newValue;
                        }

                        if (update.fields.State.newValue == "Done")
                        {
                            workItem.fields.DoneDate = update.fields.RevisedDate.newValue;
                        }
                    }
                }

                return workItem;

            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetWorkItemCycleTimeData {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        public WorkItemJson GetWorkItemAgingData(string orgName, string projectName, int workItemId)
        {
            try
            {
                var workItem = GetWorkItemData(orgName, projectName, workItemId);
                //https://dev.azure.com/{organization}/{project}/_apis/wit/workItems/{id}/updates?api-version=6.0
                string uri = String.Join("?", String.Join("/", _baseUrl, orgName, projectName, "_apis/wit/workitems", workItemId, "updates"), "&api-version=6.0");
                var result = SendRequest(uri).Result;
                var wi = JsonConvert.DeserializeObject<WorkItemUpdates>(result);
                string bc = workItem.fields.BoardColumn;
                foreach (var update in wi.value)
                {
                    //Find the most recent entry into the current column
                    //Get the current Column
                    if (string.IsNullOrWhiteSpace(bc))
                    {
                        bc = workItem.fields.State;
                        if (update.fields != null && update.fields.State != null)
                        {
                            if (string.Equals(update.fields.State.newValue, bc))
                            {
                                workItem.fields.ColumnEntryDate = update.fields.ChangedDate.newValue;
                            }
                        }
                    }
                    else
                    {
                        if (update.fields != null && update.fields.BoardColumn != null)
                        {
                            if (string.Equals(update.fields.BoardColumn.newValue, bc))
                            {
                                workItem.fields.ColumnEntryDate = update.fields.ChangedDate.newValue;
                            }
                        }

                    }
                }

                return workItem;

            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetWorkItemAgingData {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }
        public string GetWorkItemStateChangeDate(string orgName, string projName, int workItemId, string state)
        {
            try
            {
                //https://dev.azure.com/{organization}/{project}/_apis/wit/workItems/{id}/updates?api-version=6.0
                string uri = String.Join("?", String.Join("/", _baseUrl, orgName, projName, "_apis/wit/workitems", workItemId, "updates"), "&api-version=6.0");
                var result = SendRequest(uri).Result;
                var wi = JsonConvert.DeserializeObject<WorkItemUpdates>(result);
                foreach (var update in wi.value)
                {
                    if (update.fields != null && update.fields.State != null)
                    {
                        if (update.fields.State.newValue == state)
                        {
                            return update.fields.RevisedDate.newValue;
                        }
                    }
                }

                return String.Empty;

            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetWorkItemActiveDate {exception.Message} - {exception.StackTrace}");
                throw;
            }

        }
        public string GetWorkItemActiveDate(string orgName, string projName, int workItemId)
        {
            try
            {
                return GetWorkItemStateChangeDate(orgName, projName, workItemId, "Active");

            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetWorkItemActiveDate {exception.Message} - {exception.StackTrace}");
                throw;
            }

        }
        public string GetMostRecentWorkItemChangedDate(string orgName, string projName)
        {
            try
            {
                int workItemId = GetMostRecentlyChangedWorkItemId(orgName, projName);
                var changedDate = GetWorkItemChangedDate(orgName, projName, workItemId);
                return changedDate;

            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetMostRecentWorkItemChangedDate {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        public WorkItemJson GetWorkItemData(string orgName, string projectName, int workItemId)
        {
            //https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{id}?api-version=6.0
            try
            {
                string uri = String.Join("?", String.Join("/", _baseUrl, orgName, projectName, "_apis/wit/workitems", workItemId), "$top=1&api-version=6.0");
                var result = SendRequest(uri).Result;
                var wi = JsonConvert.DeserializeObject<WorkItemJson>(result);
                return wi;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetWorkItemChangedDate {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        public int GetMostRecentlyChangedWorkItemId(string orgName, string projName)
        {
            try
            {
                string wiql = "{\"query\": \"Select[System.Id] From WorkItems order by [System.ChangedDate] desc\"}";
                var response = GetAzDOQueryResults(wiql, orgName, projName); ;
                return response.workItems.Length > 0 ? response.workItems[0].id : 0;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetMostRecentlyChangedWorkItemId {exception.Message} - {exception.StackTrace}");
                throw;
            }

        }

        public string GetWorkItemChangedDate(string orgName, string projectName, int workItemId)
        {
            //https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{id}?api-version=6.0
            try
            {
                string uri = String.Join("?", String.Join("/", _baseUrl, orgName, projectName, "_apis/wit/workitems", workItemId), "$top=1&api-version=6.0");
                var result = SendRequest(uri).Result;
                var wi = JsonConvert.DeserializeObject<WorkItemJson>(result);
                return wi.fields.ChangedDate;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetWorkItemChangedDate {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        public string GetLastBuildDate(string orgName, string projName)
        {
            try
            {
                var builds = GetBuildInfoFromAzDOForProject(orgName, projName);

                return builds.count > 0 ? builds.value[0].queueTime : string.Empty;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetLastBuildDate {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        public QueryResponseJson GetAzDOQueryResults(string wiql, string orgName, string projName)
        {
            try
            {
                var content = new StringContent(wiql, Encoding.UTF8, "application/json");
                string uri = String.Join("?", String.Join("/", _baseUrl, orgName, projName, "_apis/wit/wiql"), "$top=10000&api-version=6.0");
                var result = PostRequest(uri, content).Result;
                var response = JsonConvert.DeserializeObject<QueryResponseJson>(result);
                return response;
            }
            catch (Exception exception)
            {

                _logger.ErrorFormat($"Unable to complete GetAzDOQueryResults {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        private void GetBuildList(string organization, string project)
        {
            /*
             Get Most Recent Build: https://docs.microsoft.com/en-us/rest/api/azure/devops/build/latest/get?view=azure-devops-rest-6.0
            GET https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{id}?api-version=6.0
        https://dev.azure.com/{org}accenturecio08/{proj}/_apis/build/builds?$top=1&queryOrder=queueTimeDescending&api-version=6.1-preview.6
            */
            try
            {
                string buildsUrl = String.Join("?", String.Join("/", _baseUrl, organization, project, "_apis/build/builds"), String.Join("&", "$top=1", "queryOrder=queueTimeDescending", "api -version=6.0-preview.1"));
                string result = SendRequest(buildsUrl).Result;
                var buildList = JsonConvert.DeserializeObject<BuildList>(result);

            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetWorkItem {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        private BuildList GetBuildInfoFromAzDOForProject(string organization, string project)
        {
            /*
             Get Most Recent Build: https://docs.microsoft.com/en-us/rest/api/azure/devops/build/latest/get?view=azure-devops-rest-6.0
            https://dev.azure.com/{org}accenturecio08/{proj}/_apis/build/builds?$top=1&queryOrder=queueTimeDescending&api-version=6.1-preview.6
             */
            try
            {
                string buildsUrl = String.Join("?", String.Join("/", _baseUrl, organization, project, "_apis/build/builds"), String.Join("&", "$top=1", "queryOrder=queueTimeDescending", "api -version=6.0-preview.1"));
                string result = SendRequest(buildsUrl).Result;
                var buildList = JsonConvert.DeserializeObject<BuildList>(result);
                return buildList;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetBuildInfoFromAzDOForProject {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }


        public int GetProjectPipelineCount(string orgName, string projId)
        {
            try
            {
                var pipeLines = GetPipelineListFromProjectId(orgName, projId);

                return pipeLines.count;

            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetProjectPipelineCount {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        private PipelineList GetPipelineListFromProjectId(string orgName, string projId)
        {
            try
            {
                //ttps://docs.microsoft.com/en-us/rest/api/azure/devops/pipelines/pipelines/list?view=azure-devops-rest-6.0#pipelineconfiguration
                //https://dev.azure.com/{organization}/{project}/_apis/pipelines?api-version=6.0-preview.1
                //string projectUrl = String.Join("?", String.Join("/", AzDO.BASE, AzDO.ORG, "_apis/projects", id, "properties"),String.Join("&", "api-version=6.0-preview.1", "keys=*ProcessTemplate*"));
                string projectUrl = String.Join("?", String.Join("/", _baseUrl, orgName, projId, "_apis/pipelines"), "api-version=6.0-preview.1");
                string result = SendRequest(projectUrl).Result;
                var pipelines = JsonConvert.DeserializeObject<PipelineList>(result);

                return pipelines;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetPipelineListFromProjectId {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        private ProcessList ProcessOrganization(string organization)
        {
            var processList = GetOrganizationProcessTemplates(organization);
            return processList;
        }

        public ProjectList GetCurrentOrganizationProjectList(string org)
        {
            //GET https://dev.azure.com/{organization}/_apis/projects?api-version=6.0
            string projectsUrl = String.Join("?", String.Join("/", _baseUrl, org, "_apis/projects"), "?api-version=6.0");

            //Console.WriteLine(projectsUrl);
            try
            {
                string result = SendRequest(projectsUrl).Result;
                ProjectList projectList = JsonConvert.DeserializeObject<ProjectList>(result);
                return projectList;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetCurrentOrganizationProjectList {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        public ProjectPropertyList GetProjectInfoFromProjId(string organization, string id)
        {
            //https://docs.microsoft.com/en-us/rest/api/azure/devops/core/projects/get%20project%20properties?view=azure-devops-rest-6.0
            //https://dev.azure.com/{organization}/_apis/projects/{projectId}/properties?api-version=6.0-preview.1
            //string projectUrl = String.Join("?", String.Join("/", AzDO.BASE, AzDO.ORG, "_apis/projects", id, "properties"),String.Join("&", "api-version=6.0-preview.1", "keys=*ProcessTemplate*"));
            try
            {
                string projectUrl = String.Join("?", String.Join("/", _baseUrl, organization, "_apis/projects", id, "properties"), String.Join("&", "api-version=6.0-preview.1"));
                string result = SendRequest(projectUrl).Result;
                ProjectPropertyList props = JsonConvert.DeserializeObject<ProjectPropertyList>(result);
                return props;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetProjectInfoFromProjId {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        public void GetOrganizationUsers(string organization)
        {
            //https://vssps.dev.azure.com/{organization}/_apis/graph/users?api-version=6.0-preview.1
            try
            {
                string processUrl = String.Join("?", String.Join("/", _baseUrl, organization, "_apis/graph/users"), "api-version=6.0-preview.1");

                string result = SendRequest(processUrl).Result;
                var processList = JsonConvert.DeserializeObject<ProcessList>(result);
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetOrganizationUsers {exception.Message} - {exception.StackTrace}");
                //return null;
            }

        }

        public int GetOrganizationUserCount(string organization)
        {
            try
            {
                var memberList = GetOrganizationUserEntitlements(organization);
                return memberList.TotalCount;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetOrganizationUserCount {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }
        public MemberList GetOrganizationUserEntitlements(string organization)
        {
            //https://vsaex.dev.azure.com/$OrganizationName/_apis/userentitlements?api-version=5.1-preview.2
            //api-version=6.0-preview.3
            try
            {
                string processUrl = String.Join("?", String.Join("/", "https://vsaex.dev.azure.com", organization, "_apis/userentitlements"), "api-version=5.1-preview.2");
                string result = SendRequest(processUrl).Result;
                var memberList = JsonConvert.DeserializeObject<MemberList>(result);
                return memberList;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetOrganizationUsers {exception.Message} - {exception.StackTrace}");
                return null;
            }

        }


        public ProcessList GetOrganizationProcessTemplates(string organization)
        {
            //https://docs.microsoft.com/en-us/rest/api/azure/devops/core/processes/list?view=azure-devops-rest-6.0
            //https://docs.microsoft.com/en-us/rest/api/azure/devops/wit/templates/get?view=azure-devops-rest-6.0
            //https://dev.azure.com/{organization}/_apis/process/processes?api-version=6.0
            try
            {
                string processUrl = String.Join("?", String.Join("/", _baseUrl, organization, "_apis/process/processes"), "api-version=6.0");

                string result = SendRequest(processUrl).Result;
                var processList = JsonConvert.DeserializeObject<ProcessList>(result);

                return processList;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetOrganizationProcessTemplates {exception.Message} - {exception.StackTrace}");
                return null;
            }
        }

        public bool DoesProjectExist(string organization, string project)
        {
            try
            {
                string testUrl = String.Join("?", String.Join("/", _baseUrl, organization, "_apis/projects", project, "properties"), String.Join("&", "api-version=6.0-preview.1"));
                var testResult = SendRequest(testUrl);
                var res = testResult.Result;
                //var good = testResult.IsCompletedSuccessfully;
                return true;
            }
            catch (Exception exception)
            {
                if (exception.Message.Contains("is an Invalid Url"))
                    return false;
                _logger.ErrorFormat($"Unable to complete DoesProjectExist {exception.Message} - {exception.StackTrace}");
                throw;
            }

        }

        public FieldList GetFieldsFromAzDOForProject(string organization, string project)
        {
            /*
             https://docs.microsoft.com/en-us/rest/api/azure/devops/wit/fields?view=azure-devops-rest-4.1
            GET https://dev.azure.com/{organization}/{project}/_apis/wit/fields?api-version=4.1
             */
            try
            {
                string fieldsUrl = String.Join("?", String.Join("/", _baseUrl, organization, project, "_apis/wit/fields"), "api-version=4.1");
                string result = SendRequest(fieldsUrl).Result;
                var fieldList = JsonConvert.DeserializeObject<FieldList>(result);
                return fieldList;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetFieldsFromAzDOForProject {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        public void GetConnectedOrganizations(string organization)
        {
            //https://stackoverflow.com/questions/54762368/get-all-organizations-in-azure-devops-using-rest-api
            //Post https://dev.azure.com/{organization1}/_apis/Contribution/HierarchyQuery?api-version=5.0-preview.1

            try
            {
                string organizationsUrl = String.Join("?", String.Join("/", _baseUrl, organization, "_apis/Contribution/HierarchyQuery"), "api-version=5.0-preview.1");
                string postData = "{\"contributionIds\": [\"ms.vss-features.my-organizations-data-provider\"],\"dataProviderContext\":{\"properties\":{}}}";
                var content = new StringContent(postData, Encoding.UTF8, "application/json");

                var result = PostRequest(organizationsUrl, content).Result;
                var orgsList = JsonConvert.DeserializeObject(result);
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete GetConnectedOrganizations {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }


        #endregion

        #region HTTP Methods
        private async Task<string> PostRequest(string uri, StringContent content)
        {
            try
            {
                using (HttpResponseMessage response = await _client.PostAsync(uri, content))
                {
                    response.EnsureSuccessStatusCode();
                    return (await response.Content.ReadAsStringAsync());
                }
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete Post Request {exception.Message} - {exception.StackTrace}");
                throw;
            }
        }

        private async Task<string> SendRequest(string uri)
        {
            try
            {
                using (HttpResponseMessage response = await _client.GetAsync(uri))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        response.EnsureSuccessStatusCode();
                        return (await response.Content.ReadAsStringAsync());
                    }
                    else
                    {
                        throw new ArgumentException($"{uri} is an Invalid Url");
                    }
                }
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception exception)
            {
                _logger.ErrorFormat($"Unable to complete Send Request {exception.Message} - {exception.StackTrace}");
                throw;
            }
        } // End of Send 

        #endregion
    }
}

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzDOController.JsonData
{
    public class QueryResponseJson
    {
        public string queryType { get; set; }
        public string asOf { get; set; }
        public WorkItemJson[] workItems { get; set; }
    }

    public class WorkItemJson
    {
        public int id { get; set; }
        public string url { get; set; }
        public WorkItemFieldsJson fields { get; set; }
    }

    public class WorkItemFieldsJson
    {
        [JsonProperty("System.ChangedDate")]
        public string ChangedDate { get; set; }

        [JsonProperty("System.Title")]
        public string Title { get; set; }

        [JsonProperty("Microsoft.VSTS.Scheduling.StoryPoints")]
        public decimal StoryPoints { get; set; }

        [JsonProperty("System.WorkItemType")]
        public string WorkItemType { get; set; }

        [JsonProperty("System.BoardColumn")]
        public string BoardColumn { get; set; }
        
        [JsonProperty("System.AreaPath")]
        public string AreaPath { get; set; }
        
        [JsonProperty("System.State")]
        public string State { get; set; }
        
        public string ActiveDate { get; set; }

        public string ClosedDate { get; set; }

        public string ResolvedDate { get; set; }

        public string DoneDate { get; set; }

        public string ColumnEntryDate { get; set; }
        public string ColumnExitDate { get; set; }
  
        public string Url { get; set; }



        //System.WorkItemType

        //[JsonProperty("System.ChangedBy")]
        //public string ChangedBy { get; set; }
    }

    public class WorkItemUpdates
    {
        public int count { get; set; }
        public WorkItemUpdateJson[] value { get; set; }
    }

    public class WorkItemUpdateJson
    {
        public int id { get; set; }
        public int workItemId { get; set; }
        public string url { get; set; }
        public WorkItemUpdateFieldJson fields { get; set; }
    }

    public class WorkItemUpdateFieldJson
    {
        [JsonProperty("System.ChangedDate")]
        public ChangedDate ChangedDate { get; set; }
        [JsonProperty("System.RevisedDate")]
        public RevisedDate RevisedDate { get; set; }
        [JsonProperty("System.State")]
        public State State { get; set; }
        [JsonProperty("System.BoardColumn")]
        public BoardColumn BoardColumn { get; set; }
    }

    [JsonObject("System.RevisedDate")]
    public class RevisedDate
    {
        public string oldValue { get; set; }
        public string newValue { get; set; }
    }

    [JsonObject("System.ChangedDate")]
    public class ChangedDate
    {
        public string oldValue { get; set; }
        public string newValue { get; set; }
    }

    [JsonObject("System.State")]
    public class State
    {
        public string oldValue { get; set; }
        public string newValue { get; set; }
    }

    [JsonObject("System.BoardColumn")]
    public class BoardColumn
    {
        public string oldValue { get; set; }
        public string newValue { get; set; }
    }


    public class BuildList
    {
        public int count { get; set; }
        public Build[] value { get; set; }
    }
    public class Build
    {
        public int id { get; set; }
        public string queueTime { get; set; }
        public BuildDefinition definition { get; set; }
    }

    public class BuildDefinition
    {
        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
        public string uri { get; set; }
        public string path { get; set; }
        public string type { get; set; }
        public string queueStatus { get; set; }
        public string revision { get; set; }
    }
    public class Field
    {
        public string name { get; set; }
        public string referenceName { get; set; }
        public string description { get; set; }
        public string type { get; set; }
        public string usage { get; set; }
    }

    public class FieldList
    {
        public int count { get; set; }
        public Field[] value { get; set; }
    }

    public class Pipeline
    {
        public string id { get; set; }
        public string revision { get; set; }
        public string name { get; set; }
        public string folder { get; set; }
        public string url { get; set; }

    }

    public class PipelineList
    {
        public int count { get; set; }
        public Pipeline[] value { get; set; }
    }
    public class Project
    {
        public string id { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string url { get; set; }
        public string state { get; set; }
        public string revision { get; set; }
        public string visibility { get; set; }
        public string lastUpdateTime { get; set; }
        public ProjectPropertyList properties { get; set; }
    }
    public class ProjectList
    {
        public int count { get; set; }
        public Project[] value { get; set; }
    }

    public class Process
    {
        public string id { get; set; }
        public string description { get; set; }
        public string isDefault { get; set; }
        public string type { get; set; }
        public string url { get; set; }
        public string name { get; set; }
    }

    public class UserList
    {
        public int count { get; set; }
        public User[] value { get; set; }
    }
    public class User
    {
        public string displayName { get; set; }

    }
    public class ProcessList
    {
        public int count { get; set; }
        public Process[] value { get; set; }
    }
    [JsonObject("members")]
    public class MemberList
    {
        public Member[] Members { get; set; }

        [JsonProperty("totalCount")]
        public int TotalCount { get; set; }

        [JsonProperty("continuationToken")]
        public string ContinuationToken { get; set; }
    }
    public class AccessLevel
    {
        [JsonProperty("licensingSource")]
        public string LicensingSource { get; set; }

        [JsonProperty("accountLicenseType")]
        public string AccountLicenseType { get; set; }

        [JsonProperty("msdnLicenseType")]
        public string MsdnLicenseType { get; set; }

        [JsonProperty("licenseDisplayName")]
        public string LicenseDisplayName { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("statusMessage")]
        public string StatusMessage { get; set; }

        [JsonProperty("assignmentSource")]
        public string AssignmentSource { get; set; }
    }


    public class Member
    {
        [JsonProperty("id")]
        public string Id { get; set; }
        [JsonProperty("lastAccessedDate")]
        public string LastAccessedDate { get; set; }

        [JsonProperty("accessLevel")]
        public AccessLevel AccessLevel { get; set; }
    }

    public class ProjectPropertyList
    {
        public int count { get; set; }
        public Property[] value { get; set; }
    }

    public class Property
    {
        public string name { get; set; }
        public string value { get; set; }
    }
}

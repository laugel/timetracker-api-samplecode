using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace TimetrackerOdataClient
{
    public class TFSExtender
    {
        private RestClient _client;
        private Dictionary<int, Dictionary<string, string>> _cache = new Dictionary<int, Dictionary<string, string>>();
        private string vstsToken;

        public TFSExtender(string tfsUrl)
        {
            _client = new RestClient(tfsUrl);
            _client.AddDefaultHeader("Accept", "application/json");
        }

        public TFSExtender(string tfsUrl, string vstsToken) : this(tfsUrl)
        {
            this.vstsToken = vstsToken;
        }

        public List<WorkItem> GetMultipleTfsItemsDataWithoutCache(IEnumerable<int> workItemIds)
        {
            var stringifiedIds = string.Join(",", workItemIds);
            var request = new RestRequest("/_apis/wit/workItems?ids=" + stringifiedIds + "&$expand=relations&api-version=4.1");
            if (!string.IsNullOrEmpty(vstsToken))
            {
                var auth = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($"{""}:{vstsToken}"));

                request.AddHeader("Authorization", "Basic" + auth);
            }
            else
            {
                request.UseDefaultCredentials = true;
            }

            List<WorkItem> result = new List<WorkItem>();


            Program.WriteLogLine($"Getting workitems from VSTS (ids={stringifiedIds}) ...");
            var response = _client.Get(request);

            #region Exemple de réponse
            // Exemple de response (cf https://docs.microsoft.com/en-us/rest/api/vsts/wit/work%20items/list?view=vsts-rest-4.1 )
            // {
            //  "count": 3,
            //  "value": [
            //    {
            //      "id": 297,
            //      "rev": 1,
            //      "fields": {
            //        "System.AreaPath": "Fabrikam-Fiber-Git",
            //        "System.TeamProject": "Fabrikam-Fiber-Git",
            //        "System.IterationPath": "Fabrikam-Fiber-Git",
            //        "System.WorkItemType": "Product Backlog Item",
            //        "System.State": "New",
            //        "System.Reason": "New backlog item",
            //        "System.CreatedDate": "2014-12-29T20:49:20.77Z",
            //        "System.CreatedBy": "Jamal Hartnett ",
            //        "System.ChangedDate": "2014-12-29T20:49:20.77Z",
            //        "System.ChangedBy": "Jamal Hartnett ",
            //        "System.Title": "Customer can sign in using their Microsoft Account",
            //        "Microsoft.VSTS.Scheduling.Effort": 8,
            //        "WEF_6CB513B6E70E43499D9FC94E5BBFB784_Kanban.Column": "New",
            //        "System.Description": "Our authorization logic needs to allow for users with Microsoft accounts (formerly Live Ids) - http://msdn.microsoft.com/en-us/library/live/hh826547.aspx"
            //      },
            //      "url": "https://dev.azure.com/fabrikam/_apis/wit/workItems/297"
            //    },
            //    {
            //      "id": 299,
            //      "rev": 7,
            //      "fields": {
            //        "System.AreaPath": "Fabrikam-Fiber-Git\\Website",
            //        "System.TeamProject": "Fabrikam-Fiber-Git",
            //        "System.IterationPath": "Fabrikam-Fiber-Git",
            //        "System.WorkItemType": "Task",
            //        "System.State": "To Do",
            //        "System.Reason": "New task",
            //        "System.AssignedTo": "Johnnie McLeod ",
            //        "System.CreatedDate": "2014-12-29T20:49:21.617Z",
            //        "System.CreatedBy": "Jamal Hartnett ",
            //        "System.ChangedDate": "2014-12-29T20:49:28.74Z",
            //        "System.ChangedBy": "Jamal Hartnett ",
            //        "System.Title": "JavaScript implementation for Microsoft Account",
            //        "Microsoft.VSTS.Scheduling.RemainingWork": 4,
            //        "System.Description": "Follow the code samples from MSDN",
            //        "System.Tags": "Tag1; Tag2"
            //      },
            //      "url": "https://dev.azure.com/fabrikam/_apis/wit/workItems/299"
            //    },
            //    {
            //      "id": 300,
            //      "rev": 1,
            //      "fields": {
            //        "System.AreaPath": "Fabrikam-Fiber-Git",
            //        "System.TeamProject": "Fabrikam-Fiber-Git",
            //        "System.IterationPath": "Fabrikam-Fiber-Git",
            //        "System.WorkItemType": "Task",
            //        "System.State": "To Do",
            //        "System.Reason": "New task",
            //        "System.CreatedDate": "2014-12-29T20:49:22.103Z",
            //        "System.CreatedBy": "Jamal Hartnett ",
            //        "System.ChangedDate": "2014-12-29T20:49:22.103Z",
            //        "System.ChangedBy": "Jamal Hartnett ",
            //        "System.Title": "Unit Testing for MSA login",
            //        "Microsoft.VSTS.Scheduling.RemainingWork": 3,
            //        "System.Description": "We need to ensure we have coverage to prevent regressions"
            //      },
            //      "url": "https://dev.azure.com/fabrikam/_apis/wit/workItems/300"
            //    }
            //  ]
            //}
            #endregion

            dynamic obj = JObject.Parse(response.Content);
            foreach (var workItemDescription in obj.value)
            {
                result.Add(new WorkItem()
                {
                    Id = workItemDescription.id,
                    Fields = workItemDescription.fields.ToObject<Dictionary<string, object>>(),
                    ParentId = FindParentId(workItemDescription),
                });
            }



            return result;
        }

        private int? FindParentId(dynamic workItemDescription)
        {
            var relations = ((IEnumerable)workItemDescription.relations).Cast<dynamic>();
            var parentUrl = (string)relations.SingleOrDefault(x => x.rel == "System.LinkTypes.Hierarchy-Reverse")?.url;

            if (parentUrl == null)
                return null;

            return int.Parse(parentUrl.Substring(parentUrl.LastIndexOf('/') + 1));
        }

        public Dictionary<string, string> GetTfsItemData(int id, string[] fields)
        {
            if (_cache.ContainsKey(id))
            {
                return _cache[id];
            }

            var request = new RestRequest("/_apis/wit/workItems/" + id);
            if (!string.IsNullOrEmpty(vstsToken))
            {
                var auth = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($"{""}:{vstsToken}"));

                request.AddHeader("Authorization", "Basic" + auth);
            }
            else
            {
                request.UseDefaultCredentials = true;
            }
            var fieldValues = new Dictionary<string, string>();

            try
            {
                var objTemplate = new { fields = new Dictionary<string, string>() };

                Program.WriteLogLine($"Calling {request} ...");
                var response = _client.Get(request);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeAnonymousType(response.Content, objTemplate);

                if (obj == null || obj.fields == null)
                {
                    return fieldValues;
                }

                //get fields
                foreach (var name in fields)
                {
                    if (obj.fields.ContainsKey(name))
                    {
                        fieldValues[name] = obj.fields[name];
                    }
                }

                _cache[id] = fieldValues;
            }
            catch (Exception e)
            {
                _cache[id] = fieldValues;
                //handle errors here
                Program.WriteLogLine("failed getting info for tfs#" + id + "; Exception: " + e);
            }

            return fieldValues;
        }
    }
}
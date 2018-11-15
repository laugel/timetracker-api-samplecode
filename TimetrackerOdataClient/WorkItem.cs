using System;
using System.Collections.Generic;
using System.Diagnostics;
using TimetrackerOnline.BusinessLayer.Models;

namespace TimetrackerOdataClient
{
    [DebuggerDisplay("WorkItem {Id} : {Title}")]
    public class WorkItem
    {
        public WorkItem()
        {
        }

        public WorkItem(int workItemId)
        {
            Id = workItemId;
        }

        public WorkItem(ExportItemViewModelApi trackedTimeRow)
        {
            if (trackedTimeRow == null)
                throw new ArgumentNullException(nameof(trackedTimeRow));
            Id = trackedTimeRow.TFSID.Value;
            Title = trackedTimeRow.TFSTitle;
            WorkItemType = trackedTimeRow.WorkItemType;
            TeamProject = trackedTimeRow.TeamProject;
            ParentId = trackedTimeRow.ParentTFSID;
        }

        public int Id { get; set; }

        public string Title
        {
            get
            {
                if (Fields.TryGetValue("System.Title", out object val))
                    return val?.ToString();
                return null;
            }
            set
            {
                Fields["System.Title"] = value;
            }
        }

        public Dictionary<string, object> Fields { get; set; } = new Dictionary<string, object>();
        public int? ParentId { get; internal set; }
        public string WorkItemType
        {
            get
            {
                if (Fields.TryGetValue("System.WorkItemType", out object val))
                    return val?.ToString();
                return null;
            }
            set
            {
                Fields["System.WorkItemType"] = value;
            }
        }

        public string TeamProject
        {
            get
            {
                if (Fields.TryGetValue("System.TeamProject", out object val))
                    return val?.ToString();
                return null;
            }
            set
            {
                Fields["System.TeamProject"] = value;
            }
        }
    }
}
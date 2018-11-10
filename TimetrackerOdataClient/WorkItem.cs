using System.Collections.Generic;
using System.Diagnostics;

namespace TimetrackerOdataClient
{
    [DebuggerDisplay("WorkItem {Id} : {Title}")]
    public class WorkItem
    {
        public int Id { get; set; }

        public string Title
        {
            get
            {
                if (Fields.TryGetValue("System.Title", out object val))
                    return val?.ToString();
                return null;
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
        }

        public string TeamProject {
            get
            {
                if (Fields.TryGetValue("System.TeamProject", out object val))
                    return val?.ToString();
                return null;
            }
        }
    }
}
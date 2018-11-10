using System.Collections.Generic;
using TimetrackerOnline.BusinessLayer.Models;

namespace TimetrackerOdataClient
{
    internal class TrackedTimeNode
    {
        public ExportItemViewModelApi FirstRow { get; internal set; }
        public List<ExportItemViewModelApi> Rows { get; internal set; } = new List<ExportItemViewModelApi>();
        public int TotalDurationWithChildrenInMin { get; internal set; }
        public int TotalDurationWithoutChildrenInMin { get; internal set; }
        public WorkItem WorkItem { get; internal set; }

        public string Title
        {
            get
            {
                return WorkItem?.Title ?? FirstRow?.TFSTitle;
            }
        }

        public int? WorkItemId
        {
            get
            {
                return WorkItem?.Id ?? FirstRow?.TFSID;
            }
        }

        public int? ParentId
        {
            get
            {
                return WorkItem?.ParentId ?? FirstRow?.ParentTFSID;
            }
        }

        public string TimeForTeamMember
        {
            get
            {
                return FirstRow?.TeamMember;
            }
        }
    }
}
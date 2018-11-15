using System;
using System.Collections.Generic;
using System.Linq;
using TimetrackerOnline.BusinessLayer.Models;


namespace TimetrackerOdataClient
{
    /// <summary>
    /// Représente un WorkItem avec ses saisies des temps (1 instance de ExportItemViewModelApi = 1 saisie des temps sur ce WI).
    /// </summary>
    internal class TrackedTimeNode
    {
        public ExportItemViewModelApi FirstTrackedTimeRow { get; internal set; }
        public List<ExportItemViewModelApi> DirectTrackedTimeRows { get; private set; } = new List<ExportItemViewModelApi>();

        public List<TrackedTimeNode> Childs { get; private set; } = new List<TrackedTimeNode>();


        public int TotalDurationWithChildrenInMin
        {
            get
            {
                return TotalDurationWithoutChildrenInMin + Childs.Sum(x => x.TotalDurationWithChildrenInMin);
            }
        }
        public int TotalDurationWithoutChildrenInMin
        {
            get
            {
                return DirectTrackedTimeRows.Sum(x => (int)x.DurationInSeconds / 60);
            }

        }
        public WorkItem WorkItem { get; internal set; }

        public string Title
        {
            get
            {
                return WorkItem?.Title ?? FirstTrackedTimeRow?.TFSTitle;
            }
        }

        public int? WorkItemId
        {
            get
            {
                return WorkItem?.Id ?? FirstTrackedTimeRow?.TFSID;
            }
        }

        public int? ParentId
        {
            get
            {
                return WorkItem?.ParentId ?? FirstTrackedTimeRow?.ParentTFSID;
            }
        }

        public string TimeForTeamMember
        {
            get
            {
                return FirstTrackedTimeRow?.TeamMember;
            }
        }

        public string WorkItemType
        {
            get
            {
                return WorkItem?.WorkItemType ?? FirstTrackedTimeRow?.WorkItemType;
            }
        }

        public string Project
        {
            get
            {
                return WorkItem?.TeamProject ?? FirstTrackedTimeRow?.TeamProject;
            }
        }

        internal void ForceParentId(int newParentId)
        {
            if (this.ParentId != null)
                throw new InvalidOperationException("Not allowed when ParentId is already set.");
            if (WorkItem != null)
                WorkItem.ParentId = newParentId;
            if (FirstTrackedTimeRow != null)
                FirstTrackedTimeRow.ParentTFSID = newParentId;
        }

        internal bool ContainsMember(string member)
        {
            return DirectTrackedTimeRows.Any(x => x.TeamMember == member) || Childs.Any(x => x.ContainsMember(member));
        }
    }
}
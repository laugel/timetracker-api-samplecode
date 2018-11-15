using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimetrackerOdataClient
{
    class GroupedTimeRecords
    {
        public IEnumerable<TrackedTimeNode> GroupedByWorkItem { get; set; }

        public IEnumerable<TeamMemberRecords> GroupedByTeamMember{ get; set; }
    }

    class TeamMemberRecords
    {
        public string TeamMember { get; set; }

        public IEnumerable<TrackedTimeNode> GroupedByWorkItem { get; set; }
    }
}

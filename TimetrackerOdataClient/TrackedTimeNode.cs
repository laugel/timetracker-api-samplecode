using System.Collections.Generic;

namespace TimetrackerOdataClient
{
    internal class TrackedTimeNode
    {
        public ExtendedTimetrackerRow FirstRow { get; internal set; }
        public List<ExtendedTimetrackerRow> Rows { get; internal set; } = new List<ExtendedTimetrackerRow>();
        public int TotalDurationWithChildrenInMin { get; internal set; }
        public int TotalDurationWithoutChildrenInMin { get; internal set; }
    }
}
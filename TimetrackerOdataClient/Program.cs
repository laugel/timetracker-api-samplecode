using CommandLine;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using TimetrackerOnline.BusinessLayer.Models;

namespace TimetrackerOdataClient
{
    internal class Program
    {
        private const string DateParametersFormat = @"yyyy-MM-dd";

        private static TFSExtender tfsExtender;

        public static DateTime StartDate { get; set; }
        public static DateTime EndDate { get; set; }

        private static void Main(string[] args)
        {
            bool parsed = false;
            CommandLineOptions cmd = null;
            // Get parameters
            CommandLine.Parser.Default.ParseArguments<CommandLineOptions>(args).WithParsed(x =>
                                                                                           {
                                                                                               parsed = true;
                                                                                               cmd = x;
                                                                                           })
                       .WithNotParsed(x => { Program.WriteLogLine("Check https://github.com/7pace/timetracker-api-samplecode to get samples of usage"); });

            if (!parsed)
            {
                Console.ReadLine();
                return;
            }

            tfsExtender = new TFSExtender(cmd.TfsUrl, cmd.VstsToken);

            // Create OData service context
            var context = cmd.IsWindowsAuth
                ? new TimetrackerOdataContext(cmd.ServiceUri)
                : new TimetrackerOdataContext(cmd.ServiceUri, cmd.Token);


            if (!string.IsNullOrEmpty(cmd.StartDate) && !string.IsNullOrEmpty(cmd.EndDate))
            {
                StartDate = DateTime.ParseExact(cmd.StartDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                EndDate = DateTime.ParseExact(cmd.EndDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            else
            { // fallback

                StartDate = new DateTime(2018, 6, 1);
                //StartDate = DateTime.Today.AddDays(-7);
                //StartDate = DateTime.Today.AddMonths(-6);
                EndDate = new DateTime(2018, 10, 31);
                //EndDate = DateTime.Today;
            }
            // tests only
            //var workItems = tfsExtender.GetMultipleTfsItemsDataWithoutCache(new int[] { 11251, 8385, 2934 });


            var timeExport = context.Container.TimeExport(StartDate.ToString(DateParametersFormat), EndDate.ToString(DateParametersFormat), null, null, null);
            timeExport = timeExport.AddQueryOption("api-version", "2.1");
            Program.WriteLogLine("Calling Timetracker API...");
            ExportItemViewModelApi[] timeExportResult = timeExport.ToArray();

            Program.WriteLogLine("Loading parents...");

            GroupedTimeRecords groupedItems = GroupItemsAndLoadParents(timeExportResult);

            Program.WriteLogLine("Exporting...");

            new ExcelExporter().ExportAsExcel(groupedItems);
            //Export(cmd.Format, rows);
            Program.WriteLogLine("Finished. Press ENTER to exit.");
            Console.ReadLine();
        }



        public static Dictionary<int, TrackedTimeNode> timeNodeByWorkItemId = new Dictionary<int, TrackedTimeNode>();

        private static GroupedTimeRecords GroupItemsAndLoadParents(IEnumerable<ExportItemViewModelApi> rows)
        {
            // Populate timeNodeByWorkItemId
            Program.WriteLogLine("Grouping by parents...");
            foreach (var row in rows)
            {
                if (row.TFSID == null)
                {
                    Program.WriteLogLine($"*** WARN : No WorkItemId for TimetrackerRowId={row.RowID} of {row.DurationInSeconds / 60} min on {row.RecordDate} .");
                    continue;
                }

                var workItemId = row.TFSID.Value;
                if (!timeNodeByWorkItemId.ContainsKey(workItemId))
                {
                    timeNodeByWorkItemId[workItemId] = new TrackedTimeNode { FirstTrackedTimeRow = row };
                }
                timeNodeByWorkItemId[workItemId].DirectTrackedTimeRows.Add(row);
            }

            LoadMissingParents();

            DefineHierarchyLinks();

            GroupUnparentedBugsAsAnEpic();

            var groupedByWI = timeNodeByWorkItemId.Values.Where(x => x.ParentId == null)
                              .OrderBy(x => x.Project + " " + x.WorkItemType).ThenBy(x => x.ParentId)
                              .ToList();


            Program.WriteLogLine("Grouping by members...");

            var members = rows.Select(x => x.TeamMember).Distinct().OrderBy(x => x).ToList();
            var groupedByTeamMember = new List<TeamMemberRecords>();
            foreach (var member in members)
            {
                var tmr = new TeamMemberRecords()
                {
                    TeamMember = member,
                    GroupedByWorkItem = GetCloneTimeNodesFilteredByMember(member, groupedByWI).ToList()
                };
                groupedByTeamMember.Add(tmr);
            }

            var result = new GroupedTimeRecords()
            {
                GroupedByWorkItem = groupedByWI,
                GroupedByTeamMember = groupedByTeamMember
            };
            return result;
        }

        private static IEnumerable<TrackedTimeNode> GetCloneTimeNodesFilteredByMember(string member, List<TrackedTimeNode> timeNodes)
        {
            foreach (var timeNode in timeNodes.Where(x => x.ContainsMember(member)))
            {
                var cloneTimeNode = new TrackedTimeNode()
                {
                    WorkItem = timeNode.WorkItem ?? new WorkItem(timeNode.FirstTrackedTimeRow),
                };
                cloneTimeNode.DirectTrackedTimeRows.AddRange(timeNode.DirectTrackedTimeRows.Where(x => x.TeamMember == member));
                cloneTimeNode.FirstTrackedTimeRow = cloneTimeNode.DirectTrackedTimeRows.FirstOrDefault();

                if (timeNode.Childs.Any(x => x.ContainsMember(member)))
                    cloneTimeNode.Childs.AddRange(GetCloneTimeNodesFilteredByMember(member, timeNode.Childs));
                yield return cloneTimeNode;
            }
        }

        private static void GroupUnparentedBugsAsAnEpic()
        {
            var unparentedBugs = timeNodeByWorkItemId.Values.Where(x => x.ParentId == null && x.WorkItemType == "Bug").ToList();
            var newParents = new Dictionary<string, TrackedTimeNode>();

            var artificialWorkItemId = 0;
            foreach (var bug in unparentedBugs)
            {
                var parentName = "Unparented bugs for " + bug.Project;
                if (!newParents.ContainsKey(parentName))
                {
                    artificialWorkItemId--;
                    newParents[parentName] = new TrackedTimeNode
                    {
                        WorkItem = new WorkItem(artificialWorkItemId)
                        {
                            WorkItemType = "Epic",
                            TeamProject = bug.Project,
                            Title = parentName,
                        }
                    };
                    timeNodeByWorkItemId[artificialWorkItemId] = newParents[parentName];
                }
                newParents[parentName].Childs.Add(bug);
                bug.ForceParentId(newParents[parentName].WorkItem.Id);
            }
        }

        private static void DefineHierarchyLinks()
        {
            foreach (var timeNode in timeNodeByWorkItemId.Values)
            {
                var parentId = timeNode.ParentId;
                if (parentId == null)
                    continue;
                timeNodeByWorkItemId[parentId.Value].Childs.Add(timeNode);
            }
        }

        private static void LoadMissingParents()
        {
            // LoadMissingParentsCore gère 1 seul niveau d'arborescence. Pour gérer plusieurs niveaux (le parent du parent qui est manquant),
            // il faut l'appeler consécutivement
            const int CALL_COUNT_LIMITS = 20;
            var i = 1;
            do
            {
                Program.WriteLogLine($"Loading missing parents... (step {i++})");
            } while (LoadMissingParentsCore() && i < CALL_COUNT_LIMITS);

            if (i >= CALL_COUNT_LIMITS)
            {
                Program.WriteLogLine("**** WARN : incomplete results because parents where not complete after 10 calls to VSTS.");
            }
            else
            {
                Program.WriteLogLine($"All parents found in {i} steps.");
            }
        }

        private static bool LoadMissingParentsCore()
        {
            // find missing workItem parents
            var missingParentIds = (from r in timeNodeByWorkItemId.Values
                                    where r.ParentId.HasValue
                                       && !timeNodeByWorkItemId.ContainsKey(r.ParentId.Value)
                                    select r.ParentId.Value).ToList();
            // get missing workItems from VSTS
            if (missingParentIds.Any())
            {
                var missingWorkItems = tfsExtender.GetMultipleTfsItemsDataWithoutCache(missingParentIds.Take(100)).ToList();
                foreach (var missingWorkItem in missingWorkItems)
                {
                    timeNodeByWorkItemId[missingWorkItem.Id] = new TrackedTimeNode()
                    {
                        WorkItem = missingWorkItem,
                    };
                }
                return true;
            }
            return false;
        }

        public static void WriteLogLine(string message)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} : {message}");
        }
    }

    [Serializable]
    public class ExtendedTimetrackerRow
    {

        public ExportItemViewModelApi TimetrackerRow { get; set; }
        public Dictionary<string, string> TfsData { get; set; }
    }
}
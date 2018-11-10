using ClosedXML.Excel;
using CommandLine;
using ExtendedXmlSerializer.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using TimetrackerOnline.BusinessLayer.Models;
using Formatting = Newtonsoft.Json.Formatting;

namespace TimetrackerOdataClient
{
    internal class Program
    {
        private const string DateParametersFormat = @"yyyy-MM-dd";

        private static TFSExtender tfsExtender;

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
                       .WithNotParsed(x => { Console.WriteLine("Check https://github.com/7pace/timetracker-api-samplecode to get samples of usage"); });

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

            //TODO: DEFINE DATE PERIOD HERE
            // Perform query for 3 last years
#warning TODO : adjust time
            var startDate = DateTime.Today.AddDays(-7).ToString(DateParametersFormat);
            var endDate = DateTime.Today.ToString(DateParametersFormat);

            // tests only
            //var workItems = tfsExtender.GetMultipleTfsItemsDataWithoutCache(new int[] { 11251, 8385, 2934 });


            var timeExport = context.Container.TimeExport(startDate, endDate, null, null, null);
            timeExport = timeExport.AddQueryOption("api-version", "2.1");
            Console.WriteLine("Calling Timetracker API...");
            ExportItemViewModelApi[] timeExportResult = timeExport.ToArray();
            Console.WriteLine($"Extending with TFS data for {timeExportResult.Length} results...");
            var rows = ExtendWithAdditionalFields(cmd, timeExportResult);
            // Print out the result
            foreach (var row in rows)
            {
                Console.WriteLine("{0:g} {1} {2}", row.TimetrackerRow.RecordDate, row.TimetrackerRow.TeamMember, row.TimetrackerRow.DurationInSeconds);
            }
            Console.WriteLine("Exporting...");

            var groupedItems = GroupItems(timeExportResult);

            ExportAsExcel(groupedItems);
            //Export(cmd.Format, rows);
            Console.WriteLine("Finished. Press ENTER to exit.");
            Console.ReadLine();
        }



        public static Dictionary<int, TrackedTimeNode> timeNodeByWorkItemId = new Dictionary<int, TrackedTimeNode>();

        private static IEnumerable<TrackedTimeNode> GroupItems(IEnumerable<ExportItemViewModelApi> rows)
        {
            // Populate timeNodeByWorkItemId
            foreach (var row in rows)
            {
                if (row.TFSID == null)
                {
                    Console.WriteLine($"*** WARN : No WorkItemId for TimetrackerRowId={row.RowID} of {row.DurationInSeconds / 60} min on {row.RecordDate} .");
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

            return timeNodeByWorkItemId.Values.Where(x => x.ParentId == null).OrderBy(x => x.ParentId);
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
            const int CALL_COUNT_LIMITS = 10;
            var i = 1;
            do
            {
                Console.WriteLine($"Loading missing parents... (step {i++})");
            } while (LoadMissingParentsCore() && i < CALL_COUNT_LIMITS);

            if (i >= CALL_COUNT_LIMITS)
            {
                Console.WriteLine("**** WARN : incomplete results because parents where not complete after 10 calls to VSTS.");
            }
            else
            {
                Console.WriteLine($"All parents found in {i} steps.");
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
                var missingWorkItems = tfsExtender.GetMultipleTfsItemsDataWithoutCache(missingParentIds).ToList();
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

        public static List<ExtendedTimetrackerRow> ExtendWithAdditionalFields(CommandLineOptions options, ExportItemViewModelApi[] timeExportResult)
        {
            var extender = tfsExtender;

            var extendedData = new List<ExtendedTimetrackerRow>();

            string[] tfsFields = new string[0];
            if (options.TfsFields != null)
            {
                tfsFields = options.TfsFields.ToArray();
            }
            var totalCount = timeExportResult.Count();
            var i = 0;
            foreach (var row in timeExportResult)
            {
                var extendedRow = new ExtendedTimetrackerRow
                {
                    TimetrackerRow = row
                };

                extendedData.Add(extendedRow);

                //non tfs
                if (row.TFSID == null)
                {
                    continue;
                }
                Console.WriteLine($"Extending result #{i++} (on {totalCount}) with TFS data...");
                //extendedRow.TfsData = extender.GetTfsItemData( row.TFSID.Value, tfsFields );
            }

            return extendedData;
        }

        public static void Export(string format, List<ExtendedTimetrackerRow> extendedData)
        {
            if (string.IsNullOrEmpty(format))
            {
                return;
            }

            //save here
            string location = System.Reflection.Assembly.GetExecutingAssembly().Location;

            //once you have the path you get the directory with:
            var directory = System.IO.Path.GetDirectoryName(location);

            if (format == "xml")
            {
                var serializer = new ConfigurationContainer()

                    // Configure...
                    .Create();

                var exportPath = directory + "/export.xml";

                var file = File.OpenWrite(exportPath);
                var settings = new XmlWriterSettings { Indent = true };

                var xmlTextWriter = new XmlTextWriter(file, Encoding.UTF8);
                xmlTextWriter.Formatting = System.Xml.Formatting.Indented;

                xmlTextWriter.Indentation = 4;

                serializer.Serialize(xmlTextWriter, extendedData);
                xmlTextWriter.Close();
                xmlTextWriter.Dispose();
                file.Close();
                file.Dispose();
            }
            else if (format == "json")
            {
                var json = JsonConvert.SerializeObject(extendedData, Formatting.Indented);
                var exportPath = directory + "/export.json";
                File.WriteAllText(exportPath, json);
            }
            else if (format == "xlsx")
            {
                ExportToExcel(extendedData);
            }
            else
            {
                throw new NotSupportedException("Provided format is not supported: " + format);
            }
        }

        private static void ExportToExcel(List<ExtendedTimetrackerRow> extendedData)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Timetracker Export");

                var frontCell = worksheet.Cell("A1");
                var currentCell = frontCell;
                currentCell.Value = "WorkItem ID";
                currentCell = currentCell.CellRight().SetValue("Type");
                currentCell = currentCell.CellRight().SetValue("TeamMember");

                currentCell = currentCell.CellRight().SetValue("Title");
                currentCell = currentCell.CellRight().SetValue("Duration (h)");

                frontCell = frontCell.CellBelow();
                foreach (var item in extendedData)
                {
                    var row = item.TimetrackerRow;
                    currentCell = frontCell;
                    currentCell.Value = row.TFSID;
                    currentCell = currentCell.CellRight().SetValue(row.WorkItemType);
                    currentCell = currentCell.CellRight().SetValue(row.TeamMember);
                    currentCell = currentCell.CellRight().SetValue(row.TFSTitle);
                    currentCell = currentCell.CellRight().SetValue(row.DurationInSeconds / 3600d);
                    frontCell = frontCell.CellBelow();

                }


                //worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                workbook.SaveAs($"Timetracker Export {DateTime.Now:yyyy-dd-MM_HH-mm-ss}.xlsx");


            }
        }


        private static void ExportAsExcel(IEnumerable<TrackedTimeNode> groupedItems)
        {
            const string DurationLabel = "Duration with children (h)";
            const string DurationWithoutChildrenLabel = "Duration without children (h)";
            const string WorkItemTitleLabel = "Title";
            const string TeamMemberLabel = "TeamMember";
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Timetracker Export");

                var frontCell = worksheet.Cell("B2");
                var firstCell = frontCell;
                var headerCell = frontCell;
                headerCell.Value = "WorkItem ID";
                headerCell = headerCell.CellRight().SetValue("Project");
                headerCell = headerCell.CellRight().SetValue("Type");
                headerCell = headerCell.CellRight().SetValue("ParentId");
                headerCell = headerCell.CellRight().SetValue("Parent (lev2)");
                headerCell = headerCell.CellRight().SetValue("Parent (lev3)");
                headerCell = headerCell.CellRight().SetValue("Parent (lev4+)");
                headerCell = headerCell.CellRight().SetValue(WorkItemTitleLabel);
                headerCell = headerCell.CellRight().SetValue(TeamMemberLabel);
                headerCell = headerCell.CellRight().SetValue(DurationLabel);
                headerCell = headerCell.CellRight().SetValue(DurationWithoutChildrenLabel);

                frontCell = frontCell.CellBelow();
                var lastDataCell = frontCell;
                AddWorkItems(groupedItems, ref frontCell, ref lastDataCell, 1);

                var excelTable = worksheet.Range(firstCell, lastDataCell).CreateTable();
                // Add the totals row
                excelTable.ShowTotalsRow = true;
                // Put the average on the field "Income"
                // Notice how we're calling the cell by the column name
                excelTable.Field(DurationWithoutChildrenLabel).TotalsRowFunction = XLTotalsRowFunction.Sum;
                // Put a label on the totals cell of the field "Title"
                excelTable.Field(DurationLabel).TotalsRowLabel = "Sum:";

                worksheet.Columns().AdjustToContents();
                //worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                workbook.SaveAs($"Timetracker Export {DateTime.Now:yyyy-dd-MM_HH-mm-ss}.xlsx");


            }
        }

        private static void AddWorkItems(IEnumerable<TrackedTimeNode> groupedItems, ref IXLCell frontCell, ref IXLCell lastDataCell, int level)
        {
            foreach (var item in groupedItems)
            {
                var firstCell = frontCell;
                var row = item;
                var currentCell = frontCell;
                if (item.FirstTrackedTimeRow == null)
                {
                    if (row.TotalDurationWithoutChildrenInMin > 0)
                        throw new InvalidOperationException("Etat inatendu : présence d'un temps saisi sans enregistrement Timetracker.");
                    // aucune saisie des temps directe :
                    currentCell = AddHeader(level, row, currentCell);
                    currentCell = currentCell.CellRight().SetValue(row.TimeForTeamMember);
                    currentCell = currentCell.CellRight().SetValue(row.TotalDurationWithChildrenInMin / 60d);
                    currentCell = currentCell.CellRight().SetValue(row.TotalDurationWithoutChildrenInMin / 60d);
                    lastDataCell = currentCell;
                    frontCell = frontCell.CellBelow();
                }
                else
                {
                    // plusieurs saisies des temps directes sur cet item :
                    foreach (var workItemTimes in row.DirectTrackedTimeRows)
                    {
                        currentCell = frontCell;
                        currentCell = AddHeader(level, row, currentCell);
                        currentCell = currentCell.CellRight().SetValue(workItemTimes.TeamMember);
                        currentCell = currentCell.CellRight().SetValue(workItemTimes.DurationInSeconds / 3600d);
                        currentCell = currentCell.CellRight().SetValue(workItemTimes.DurationInSeconds / 3600d);
                        lastDataCell = currentCell;
                        frontCell = frontCell.CellBelow();
                    }
                }
                AddWorkItems(row.Childs, ref frontCell, ref lastDataCell, level + 1);
                //if (firstCell.CellBelow() != frontCell)
                //{
                //    // regrouper car il y a plus d'une ligne sur cet item:
                //    firstCell.Worksheet.Rows(firstCell.Address.RowNumber, frontCell.Address.RowNumber).Group(level); // Create an outline
                //}
            }
        }

        private static IXLCell AddHeader(int level, TrackedTimeNode row, IXLCell currentCell)
        {
            currentCell.Value = row.WorkItemId;
            currentCell = currentCell.CellRight().SetValue(row.Project);
            currentCell = currentCell.CellRight().SetValue(row.WorkItemType);
            currentCell = currentCell.CellRight().SetValue(level == 2 ? row.ParentId : null);
            currentCell = currentCell.CellRight().SetValue(level == 3 ? row.ParentId : null);
            currentCell = currentCell.CellRight().SetValue(level == 4 ? row.ParentId : null);
            currentCell = currentCell.CellRight().SetValue(level > 4 ? row.ParentId : null);

            currentCell = currentCell.CellRight().SetValue(row.Title);
            return currentCell;
        }
    }

    [Serializable]
    public class ExtendedTimetrackerRow
    {

        public ExportItemViewModelApi TimetrackerRow { get; set; }
        public Dictionary<string, string> TfsData { get; set; }
    }
}
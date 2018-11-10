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

            // Create OData service context
            var context = cmd.IsWindowsAuth
                ? new TimetrackerOdataContext(cmd.ServiceUri)
                : new TimetrackerOdataContext(cmd.ServiceUri, cmd.Token);

            //TODO: DEFINE DATE PERIOD HERE
            // Perform query for 3 last years
#warning TODO : adjust time
            var startDate = DateTime.Today.AddDays(-7).ToString(DateParametersFormat);
            var endDate = DateTime.Today.ToString(DateParametersFormat);

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

            var groupedItems = GroupItems(rows);

            Export(cmd.Format, rows);
            Console.WriteLine("Finished. Press ENTER to exit.");
            Console.ReadLine();
        }

        public static Dictionary<int, TrackedTimeNode> timeNodeByWorkItemId = new Dictionary<int, TrackedTimeNode>();

        private static IEnumerable<TrackedTimeNode> GroupItems(List<ExtendedTimetrackerRow> rows)
        {
            // Populate timeNodeByWorkItemId
            foreach (var row in rows)
            {
                if (row.TimetrackerRow.TFSID == null)
                {
                    Console.WriteLine($"*** WARN : No WorkItemId for TimetrackerRowId={row.TimetrackerRow.RowID} of {row.TimetrackerRow.DurationInSeconds / 60} min on {row.TimetrackerRow.RecordDate} .");
                    continue;
                }

                var workItemId = row.TimetrackerRow.TFSID.Value;
                if (!timeNodeByWorkItemId.ContainsKey(workItemId))
                {
                    timeNodeByWorkItemId[workItemId] = new TrackedTimeNode { FirstRow = row };
                }
                timeNodeByWorkItemId[workItemId].Rows.Add(row);
                timeNodeByWorkItemId[workItemId].TotalDurationWithChildrenInMin += Convert.ToInt32(row.TimetrackerRow.DurationInSeconds / 60);
                timeNodeByWorkItemId[workItemId].TotalDurationWithoutChildrenInMin += Convert.ToInt32(row.TimetrackerRow.DurationInSeconds / 60);

            }

            // Complete timeNodeByWorkItemId with missing parent items
            foreach (var node in timeNodeByWorkItemId.Values.ToList())
            {
                var parentId = node.FirstRow.TimetrackerRow.ParentTFSID;
                if (parentId == null)
                    continue; // no parent for the item
                if (!timeNodeByWorkItemId.ContainsKey(parentId.Value))
                {
                    var parent = new TrackedTimeNode
                    {
                        
                    };


                 }
                // find parent in TFS
            }

            // TODO
            return null;
        }

        public static List<ExtendedTimetrackerRow> ExtendWithAdditionalFields(CommandLineOptions options, ExportItemViewModelApi[] timeExportResult)
        {
            var extender = new TFSExtender(options.TfsUrl, options.VstsToken);

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
    }

    [Serializable]
    public class ExtendedTimetrackerRow
    {
        


        public ExportItemViewModelApi TimetrackerRow { get; set; }
        public Dictionary<string, string> TfsData { get; set; }
    }
}
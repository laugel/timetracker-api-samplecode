using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace TimetrackerOdataClient
{
    class ExcelExporter
    {
        private static IXLWorksheet worksheet;

        public static void ExportAsExcel(IEnumerable<TrackedTimeNode> groupedItems)
        {
            const string TotalDurationLabel = "Total duration (with children) (h)";
            const string DirectDurationWithoutChildrenLabel = "Direct duration (without children) (h)";
            const string WorkItemTitleLabel = "Title";
            const string TeamMemberLabel = "TeamMember";
            using (var workbook = new XLWorkbook())
            {
                worksheet = workbook.Worksheets.Add("Timetracker Export");
                worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                worksheet.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Left;

                var frontCell = worksheet.Cell("B4");
                var firstCell = frontCell;
                var headerCell = frontCell;
                
                headerCell.Value = "WorkItem ID";
                headerCell = headerCell.CellRight().SetValue("Project");
                headerCell = headerCell.CellRight().SetValue("Type");
                headerCell = headerCell.CellRight().SetValue("ParentId");
                headerCell = headerCell.CellRight().SetValue("Parent (lev2)");
                var parentLvl2Column = headerCell.WorksheetColumn();
                headerCell = headerCell.CellRight().SetValue("Parent (lev3)");
                headerCell = headerCell.CellRight().SetValue("Parent (lev4+)");
                var parentLvl4Column = headerCell.WorksheetColumn();
                headerCell = headerCell.CellRight().SetValue(WorkItemTitleLabel);
                var titleColumn = headerCell.WorksheetColumn();
                headerCell = headerCell.CellRight().SetValue(TotalDurationLabel);
                headerCell = headerCell.CellRight().SetValue("SubTotal (Lvl2)");
                var subTotalLvl1Column = headerCell.WorksheetColumn();
                headerCell = headerCell.CellRight().SetValue("SubTotal (Lvl3)");
                headerCell = headerCell.CellRight().SetValue("SubTotal (Lvl4+)");
                var subTotalLvl4Column = headerCell.WorksheetColumn();
                headerCell = headerCell.CellRight().SetValue(DirectDurationWithoutChildrenLabel);
                headerCell = headerCell.CellRight().SetValue("Date");
                headerCell = headerCell.CellRight().SetValue(TeamMemberLabel);

                frontCell = frontCell.CellBelow();
                var lastDataCell = frontCell;
                AddWorkItems(groupedItems, ref frontCell, ref lastDataCell, 1);

                var excelTable = worksheet.Range(firstCell, lastDataCell).CreateTable();
                // Add the totals row
                excelTable.ShowTotalsRow = true;
                // Put the average on the field "Income"
                // Notice how we're calling the cell by the column name
                excelTable.Field(DirectDurationWithoutChildrenLabel).TotalsRowFunction = XLTotalsRowFunction.Sum;
                excelTable.Field(TotalDurationLabel).TotalsRowFunction = XLTotalsRowFunction.Sum;
                // Put a label on the totals cell of the field "Title"
                excelTable.Field(WorkItemTitleLabel).TotalsRowLabel = "Sum:";

                
                foreach (var action in excelGroupingActions)
                {
                    action();
                }

                worksheet.Columns().AdjustToContents();

                titleColumn.Width = 60;
                var titleCell = worksheet.Cell("B2").SetValue($"Time tracked between {Program.StartDate} and {Program.EndDate} (in hours)");
                titleCell.Style.Font.Bold = true;
                titleCell.Style.Font.FontSize = 16;

                // group nested parent columns under "parentId"
                worksheet.Columns(parentLvl2Column.ColumnNumber(), parentLvl4Column.ColumnNumber()).Group();
                worksheet.Columns(subTotalLvl1Column.ColumnNumber(), subTotalLvl4Column.ColumnNumber()).Group();

                worksheet.CollapseRows();
                worksheet.CollapseColumns();

                workbook.SaveAs($"Timetracker Export {DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx");


            }
        }

        private static Stack<Action> excelGroupingActions = new Stack<Action>();

        private static void AddWorkItems(IEnumerable<TrackedTimeNode> groupedItems, ref IXLCell frontCell, ref IXLCell lastDataCell, int level)
        {
            foreach (var item in groupedItems)
            {
                var firstCell = frontCell;
                var row = item;
                var currentCell = frontCell;
                // 3 lots par WorkItem :
                // 1/ Element récapitulatif pour le WI
                currentCell = AddValues(level, row, currentCell, null,
                                        row.TotalDurationWithChildrenInMin / 60d, null, null);
                lastDataCell = currentCell;
                frontCell = frontCell.CellBelow();

                // 3/ Liste des saisies directes sur le WI
                if (row.FirstTrackedTimeRow != null)
                {
                    // plusieurs saisies des temps directes sur cet item :
                    foreach (var workItemTimes in row.DirectTrackedTimeRows)
                    {
                        currentCell = frontCell;
                        currentCell = AddValues(level + 1, row, currentCell, workItemTimes.TeamMember,
                                                null, workItemTimes.DurationInSeconds / 3600d, workItemTimes.RecordDate.Date);
                        lastDataCell = currentCell;
                        frontCell = frontCell.CellBelow();
                    }
                }
                // 3/ Liste des WI enfants
                AddWorkItems(row.Childs, ref frontCell, ref lastDataCell, level + 1);


                if (firstCell.CellBelow() != frontCell)
                {
                    // regrouper car il y a plus d'une ligne sur cet item:
                    // cf https://stackoverflow.com/questions/25756741/closedxml-outline pour identifier quelles lignes utiliser pour appeler Group().
                    var rowsToGroup = firstCell.Worksheet.Rows(firstCell.CellBelow().Address.RowNumber, frontCell.CellAbove().Address.RowNumber);
                    excelGroupingActions.Push(() =>
                    {
                        rowsToGroup.Group(level); // Create an outline
                    });
                }
            }
        }

        private static IXLCell AddValues(int level, TrackedTimeNode row, IXLCell currentCell, string teamMember,
                                         double? wiTotalDurationInMin, double? wiDirectDurationInMin, DateTime? recordDate)
        {
            var firstCell = currentCell;
            currentCell.Value = row.WorkItemId;
            currentCell = currentCell.CellRight().SetValue(row.Project);
            currentCell = currentCell.CellRight().SetValue(row.WorkItemType);
            currentCell = currentCell.CellRight().SetValue(level == 2 ? row.ParentId : null);
            currentCell = currentCell.CellRight().SetValue(level == 3 ? row.ParentId : null);
            currentCell = currentCell.CellRight().SetValue(level == 4 ? row.ParentId : null);
            currentCell = currentCell.CellRight().SetValue(level > 4 ? row.ParentId : null);

            currentCell = currentCell.CellRight().SetValue(row.Title);
            currentCell = currentCell.CellRight().SetValue(level == 1 ? wiTotalDurationInMin : null);
            currentCell = currentCell.CellRight().SetValue(level == 2 ? wiTotalDurationInMin : null);
            currentCell = currentCell.CellRight().SetValue(level == 3 ? wiTotalDurationInMin : null);
            currentCell = currentCell.CellRight().SetValue(level >= 4 ? wiTotalDurationInMin : null);
            currentCell = currentCell.CellRight().SetValue(wiDirectDurationInMin);
            currentCell = currentCell.CellRight().SetValue(recordDate);
            currentCell = currentCell.CellRight().SetValue(teamMember);

            var range = worksheet.Range(firstCell, currentCell);
            if (level == 1 && wiDirectDurationInMin == null) // 1stl level WorkItem (not a time tracked record)
                range.Style.Fill.BackgroundColor = XLColor.FromArgb(216, 228, 188);
            else if (level > 1 && wiDirectDurationInMin == null) // next level WorkItem (not a time tracked record)
                range.Style.Fill.BackgroundColor = XLColor.FromArgb(235, 241, 222);
            else
            {
                // wiDirectDurationInMin != null (time tracked record)
                range.Style.Fill.BackgroundColor = XLColor.White;
            }
            return currentCell;
        }

    }
}

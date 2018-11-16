using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TimetrackerOdataClient
{
    class ExcelExporter
    {
        private IXLWorksheet worksheet;

        private const string TotalDurationLabel = "Total duration (with children) (h)";
        private const string DirectDurationWithoutChildrenLabel = "Direct duration (without children) (h)";

        private const string WorkItemTitleLabel = "Title";
        private const string TeamMemberLabel = "TeamMember";

        public void ExportAsExcel(GroupedTimeRecords groupedTimeNodes)
        {
            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                AddSheetWithGroupingByTopParents(groupedTimeNodes.GroupedByWorkItem, workbook);
                AddSheetWithGroupingByTeamMember(groupedTimeNodes, workbook);


                workbook.SaveAs($"Timetracker Export {DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx");
            }
        }

        private void AddSheetWithGroupingByTeamMember(GroupedTimeRecords groupedTimeNodes, XLWorkbook workbook)
        {
            worksheet = workbook.Worksheets.Add("Grouped by team members");
            worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
            worksheet.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Left;

            var frontCell = worksheet.Cell("A4");
            var firstCell = frontCell;
            var headerCell = frontCell;

            headerCell.Value = "Member";
            headerCell.WorksheetColumn().Width = 22;
            headerCell = headerCell.CellRight().SetValue("WorkItem ID");
            headerCell.WorksheetColumn().Width = 12;
            headerCell = headerCell.CellRight().SetValue("Project");
            headerCell.WorksheetColumn().Width = 18;
            headerCell = headerCell.CellRight().SetValue("Type");
            headerCell.WorksheetColumn().Width = 9;
            headerCell = headerCell.CellRight().SetValue("ParentId");
            headerCell.WorksheetColumn().Width = 8;
            headerCell = headerCell.CellRight().SetValue("Parent (lev2)");
            headerCell.WorksheetColumn().Width = 11;
            var parentLvl2Column = headerCell.WorksheetColumn();
            headerCell = headerCell.CellRight().SetValue("Parent (lev3)");
            headerCell.WorksheetColumn().Width = 11;
            headerCell = headerCell.CellRight().SetValue("Parent (lev4+)");
            headerCell.WorksheetColumn().Width = 11;
            var parentLvl4Column = headerCell.WorksheetColumn();
            headerCell = headerCell.CellRight().SetValue(WorkItemTitleLabel);
            var titleColumn = headerCell.WorksheetColumn();
            headerCell = headerCell.CellRight().SetValue(TotalDurationLabel);
            var totalDurationColumn = headerCell.WorksheetColumn();
            headerCell.WorksheetColumn().Width = 14;
            headerCell = headerCell.CellRight().SetValue("SubTotal (Lvl2)");
            headerCell.WorksheetColumn().Width = 14;
            var subTotalLvl1Column = headerCell.WorksheetColumn();
            headerCell = headerCell.CellRight().SetValue("SubTotal (Lvl3)");
            headerCell.WorksheetColumn().Width = 14;
            headerCell = headerCell.CellRight().SetValue("SubTotal (Lvl4+)");
            headerCell.WorksheetColumn().Width = 14;
            var subTotalLvl4Column = headerCell.WorksheetColumn();
            headerCell = headerCell.CellRight().SetValue(DirectDurationWithoutChildrenLabel);
            headerCell.WorksheetColumn().Width = 14;
            headerCell = headerCell.CellRight().SetValue("Date");
            headerCell.WorksheetColumn().Width = 14;
            headerCell = headerCell.CellRight().SetValue(TeamMemberLabel);
            headerCell.WorksheetColumn().Width = 15;

            frontCell = frontCell.CellBelow();
            var lastDataCell = frontCell;

            foreach (var member in groupedTimeNodes.GroupedByTeamMember)
            {
                var currentCell = frontCell;
                var firstMemberCell = currentCell;

                currentCell.SetValue(member.TeamMember);
                currentCell = currentCell.CellRight(); // WorkItemID
                currentCell = currentCell.CellRight(); // Project
                currentCell = currentCell.CellRight(); // WorkItemType
                currentCell = currentCell.CellRight();//.SetValue(level == 2 ? row.ParentId : null);
                currentCell = currentCell.CellRight();//.SetValue(level == 3 ? row.ParentId : null);
                currentCell = currentCell.CellRight();//.SetValue(level == 4 ? row.ParentId : null);
                currentCell = currentCell.CellRight();//.SetValue(level > 4 ? row.ParentId : null);

                currentCell = currentCell.CellRight(); // Title
                currentCell = currentCell.CellRight().SetValue(member.GroupedByWorkItem.Sum(x => x.TotalDurationWithChildrenInMin / 60d));
                currentCell = currentCell.CellRight();// (level == 2 ? wiTotalDurationInMin : null);
                currentCell = currentCell.CellRight();// (level == 3 ? wiTotalDurationInMin : null);
                currentCell = currentCell.CellRight();// (level >= 4 ? wiTotalDurationInMin : null);
                currentCell = currentCell.CellRight();// (wiDirectDurationInMin);
                currentCell = currentCell.CellRight();// (recordDate);
                currentCell = currentCell.CellRight();// (teamMember);

                frontCell = frontCell.CellBelow().CellRight();
                AddWorkItems(member.GroupedByWorkItem, ref frontCell, ref lastDataCell, 2);
                frontCell = frontCell.CellLeft();

                // regrouper
                // cf https://stackoverflow.com/questions/25756741/closedxml-outline pour identifier quelles lignes utiliser pour appeler Group().
                var rowsToGroup = firstMemberCell.Worksheet.Rows(firstMemberCell.CellBelow().Address.RowNumber, frontCell.CellAbove().Address.RowNumber);
                excelGroupingActions.Push(() =>
                {
                    rowsToGroup.Group(1); // Create an outline
                });
            }

            foreach (var action in excelGroupingActions)
            {
                action();
            }

            

            var excelTable = worksheet.Range(firstCell, lastDataCell).CreateTable();
            // Add the totals row
            //excelTable.ShowTotalsRow = true;

            // following lines corrupt the XLSX file :
            //excelTable.Field(DirectDurationWithoutChildrenLabel).TotalsRowFunction = XLTotalsRowFunction.Sum;
            //excelTable.Field(titleColumn.RangeAddress.FirstAddress.ColumnNumber).TotalsRowFunction = XLTotalsRowFunction.Sum;

            //excelTable.Field(WorkItemTitleLabel).TotalsRowLabel = "Sum :";
            var sumDurationCell = worksheet.Cell(lastDataCell.CellBelow().Address.RowNumber, totalDurationColumn.ColumnNumber());
            var firstDurationCell = totalDurationColumn.FirstCellUsed().CellBelow();
            var lastDurationCell = sumDurationCell.CellAbove();

            sumDurationCell.FormulaA1 = $"=SUM({firstDurationCell.Address}:{lastDurationCell.Address})";
            sumDurationCell.Style.Font.Bold = true;
            sumDurationCell.CellLeft().Value = "Sum : ";
            sumDurationCell.CellLeft().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            sumDurationCell.CellLeft().Style.Font.Bold = true;

            titleColumn.Width = 60;
            var titleCell = worksheet.Cell("B2").SetValue($"Time tracked between {Program.StartDate:yyyy-MM-dd} and {Program.EndDate:yyyy-MM-dd} (in hours)");
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 16;

            // group nested parent columns under "parentId"
            worksheet.Columns(parentLvl2Column.ColumnNumber(), parentLvl4Column.ColumnNumber()).Group();
            worksheet.Columns(subTotalLvl1Column.ColumnNumber(), subTotalLvl4Column.ColumnNumber()).Group();

            worksheet.CollapseRows();
            worksheet.CollapseColumns();
        }

        private void AddSheetWithGroupingByTopParents(IEnumerable<TrackedTimeNode> groupedItems, XLWorkbook workbook)
        {
            worksheet = workbook.Worksheets.Add("Grouped by parents");
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
            var titleCell = worksheet.Cell("B2").SetValue($"Time tracked between {Program.StartDate:yyyy-MM-dd} and {Program.EndDate:yyyy-MM-dd} (in hours)");
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 16;

            // group nested parent columns under "parentId"
            worksheet.Columns(parentLvl2Column.ColumnNumber(), parentLvl4Column.ColumnNumber()).Group();
            worksheet.Columns(subTotalLvl1Column.ColumnNumber(), subTotalLvl4Column.ColumnNumber()).Group();

            worksheet.CollapseRows();
            worksheet.CollapseColumns();
        }

        private Stack<Action> excelGroupingActions = new Stack<Action>();

        private void AddWorkItems(IEnumerable<TrackedTimeNode> groupedItems, ref IXLCell frontCell, ref IXLCell lastDataCell, int level)
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

        private IXLCell AddValues(int level, TrackedTimeNode row, IXLCell currentCell, string teamMember,
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

            string indent = new string(' ', Math.Max(0, 4 * (level - 2))); // for a "tree" visualization
            currentCell = currentCell.CellRight().SetValue(indent + row.Title); 
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DataImportAPI.Models;
using DataImportAPI.Models.DailyReport;
using DataImportAPI.Utilities;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using OpenXmlPowerTools;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace DataImportAPI.Utilities.ExcelUtilities.ExcelWriterUtility
{
    public class DailyReportWriter : IDailyReportWriter
    {
        public List<RowDfn> Rows { get; set; }
        public WorkbookDfn GenerateExcelSheet(DailyReportData dailyReportSheetData)
        {

            WorkbookDfn wb = new WorkbookDfn();
            WorksheetDfn ws = new WorksheetDfn();
            Rows = new List<RowDfn>();
            SetWorksheetNames(ws,dailyReportSheetData.SheetName);
            SetColumnHeaders(ws);
            SetDailyReportInfo(ws, dailyReportSheetData.ReportInfo);
            Rows.Add(new RowDfn{Cells = new CellDfn[]{new CellDfn{Value=""}}});
            SetActivityLogHeader(ws, dailyReportSheetData.ActivityLogHeaders);
            SetActivityLogRows(ws, dailyReportSheetData.ActivityLog);
            Rows.Add(new RowDfn{Cells = new CellDfn[]{new CellDfn{Value=""}}});
            SetReportBudgetHeader(ws, dailyReportSheetData.ReportBudgetHeaders);
            SetReportBudgetRows(ws, dailyReportSheetData.ReportBudget);
            ws.Rows=Rows;
            List<WorksheetDfn> worksheetDfns = new List<WorksheetDfn> { ws };
            wb.Worksheets = worksheetDfns;
            return wb;
        }

        private void SetWorksheetNames(WorksheetDfn worksheet,string sheetName)
        {
            worksheet.Name = sheetName;
            worksheet.TableName = sheetName;
        }
        private void SetColumnHeaders(WorksheetDfn worksheet)
        {
            List<CellDfn> columnHeaders = new List<CellDfn>();
            for (int i=0;i<13;i++)
            {
                CellDfn cell = new CellDfn();

                if (i==6)
                {
                    cell.Value = "DAILY REPORT SHEET";
                    cell.Bold = true;
                    columnHeaders.Add(cell);

                }
                else
                {
                    
                    cell.Value = "-";                   
                    columnHeaders.Add(cell);
                }
                
                
            }
            worksheet.ColumnHeadings = columnHeaders;
        }
        private void SetDailyReportInfo(WorksheetDfn worksheet, List<DailyReportInfo> dailyReportInfo)
        {

            RowDfn row = new RowDfn();
            List<CellDfn> cells = new List<CellDfn>();

            for (int i = 1; i <= dailyReportInfo.Count; i++)
            {
                CellDfn cell1 = new CellDfn();
                cell1.Value = dailyReportInfo[i-1].Header;
                cell1.Bold=true;
                cells.Add(cell1);

                CellDfn cell2 = new CellDfn();
                cell2.Value = dailyReportInfo[i-1].Value;
                cells.Add(cell2);

                CellDfn cell3 = new CellDfn();
                cell3.Value = "";
                cells.Add(cell3);

                if (i % 3 == 0)
                {
                    row.Cells = cells.ToList();
                    Rows.Add(row);
                    cells = new List<CellDfn>();
                    row = new RowDfn();
                }
            }
        }
        private void SetActivityLogHeader(WorksheetDfn worksheet, List<string> activityLogHeaders)
        {
            RowDfn row = new RowDfn();
            List<CellDfn> cells = new List<CellDfn>();

            foreach (var header in activityLogHeaders)
            {
                CellDfn cell = new CellDfn();
                cell.Value = header;
                cell.Bold = true;
                cells.Add(cell);
            }
            row.Cells = cells;
            Rows.Add(row);

        }
        private void SetActivityLogRows(WorksheetDfn worksheet, List<ActivityLogEntryData> activityLogEntry)
        {
            RowDfn row;
            List<CellDfn> cells;

            foreach (var activityLog in activityLogEntry)
            {
                cells = new List<CellDfn>();
                row = new RowDfn();

                CellDfn from = new CellDfn();
                from.Value = activityLog.From;
                cells.Add(from);

                CellDfn to = new CellDfn();
                to.Value = activityLog.To;
                cells.Add(to);

                CellDfn elapsedTime = new CellDfn();
                elapsedTime.Value = activityLog.ElapsedTime;
                cells.Add(elapsedTime);

                CellDfn cumTime = new CellDfn();
                cumTime.Value = activityLog.CumulativeTime;
                cells.Add(cumTime);

                CellDfn depth = new CellDfn();
                depth.Value = activityLog.Depth;
                cells.Add(depth);

                CellDfn mudWeight = new CellDfn();
                mudWeight.Value = activityLog.MudWeight;
                cells.Add(mudWeight);

                CellDfn activity = new CellDfn();
                activity.Value = activityLog.Activity;
                cells.Add(activity);

                CellDfn milestone = new CellDfn();
                milestone.Value = activityLog.Milestone;
                cells.Add(milestone);

                CellDfn unplanned_planned = new CellDfn();
                unplanned_planned.Value = activityLog.Unplanned_Planned;
                cells.Add(unplanned_planned);

                CellDfn hasNPT = new CellDfn();
                hasNPT.Value = activityLog.HasNpt;
                cells.Add(hasNPT);

                CellDfn NPTReference = new CellDfn();
                NPTReference.Value = activityLog.NptReference;
                cells.Add(NPTReference);

                CellDfn descrption = new CellDfn();
                descrption.Value = activityLog.Description;
                cells.Add(descrption);

                row.Cells = cells.ToList();
                Rows.Add(row);

            }
        }

        private void SetReportBudgetHeader(WorksheetDfn worksheet, List<string> reportBudgetHeaders)
        {
            RowDfn row = new RowDfn();
            List<CellDfn> cells = new List<CellDfn>();

            foreach (var header in reportBudgetHeaders)
            {
                CellDfn cell = new CellDfn();
                cell.Value = header;
                cell.Bold = true;
                cells.Add(cell);
            }

            row.Cells = cells;
            Rows.Add(row);
        }

        private void SetReportBudgetRows(WorksheetDfn worksheet, List<DailyReportBudgetData> reportBudgetData)
        {
            RowDfn row;
            List<CellDfn> cells;

            foreach (var budgetData in reportBudgetData)
            {
                cells = new List<CellDfn>();
                row = new RowDfn();


                CellDfn cell = new CellDfn();
                cell.Value = budgetData.LineItem;
                cells.Add(cell);

                foreach (var milestone in budgetData.Milestones)
                {
                    CellDfn milestoneCell = new CellDfn();
                    milestoneCell.Value = milestone;
                    cells.Add(milestoneCell);
                }
                row.Cells = cells.ToList();
                Rows.Add(row);
            }
        }
    }
}
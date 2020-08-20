using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DataImportAPI.Models;
using DataImportAPI.Models.DailyReport;
using DataImportAPI.Utilities;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using OpenXmlPowerTools;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace DataImportAPI.Utilities.ExcelUtilities.ExcelWriterUtility
{
    public interface IDailyReportWriterWithStyleSheet
    {
        byte[] GenerateExcelSheet(DailyReportData dailyReportSheetData);
    }

    public class DailyReportWriterWithStyleSheet : IDailyReportWriterWithStyleSheet
    {
        public byte[] GenerateExcelSheet(DailyReportData dailyReportSheetData)
        {
            var stream = new MemoryStream();
            var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            var workbookpart = document.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            WorkbookStylesPart stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = Utility.GenerateStlyeSheet();
            stylesPart.Stylesheet.Save();

            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();

            worksheetPart.Worksheet = new Worksheet(sheetData);

            var sheets = document.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = dailyReportSheetData.SheetName ?? "Sheet 1" };
            sheets.AppendChild(sheet);

            SetColumnHeaders(sheetData);
            SetDailyReportInfo(sheetData, dailyReportSheetData.ReportInfo);
            SetEmptyRow(sheetData);
            SetActivityLogHeader(sheetData, dailyReportSheetData.ActivityLogHeaders);
            SetActivityLogRows(sheetData, dailyReportSheetData.ActivityLog);
            SetEmptyRow(sheetData);
            SetReportBudgetHeader(sheetData, dailyReportSheetData.ReportBudgetHeaders);
            SetReportBudgetRows(sheetData, dailyReportSheetData.ReportBudget);

            //SetColumnProperties(worksheetPart.Worksheet);

            workbookpart.Workbook.Save();

            document.Close();
            
            stream.Position=0;

            return stream.ToArray();
        }

        private void SetColumnProperties(Worksheet worksheet)
        {
            Columns columns = new Columns(
                new Column
                {
                    Min = 12,
                    Max = 12,
                    Width = 12,
                    CustomWidth = true
                }
            );
            worksheet.AppendChild(columns);
        }

        private void SetEmptyRow(SheetData sheetData){
            Row row=new Row();
            row.Append(
                CreateCell("",CellValues.String)
            );
            sheetData.Append(row);
        }
        private void SetReportBudgetRows(SheetData sheetData, List<DailyReportBudgetData> reportBudgetData)
        {
            Row row;
            List<CellDfn> cells;

            foreach (var budgetData in reportBudgetData)
            {
                cells = new List<CellDfn>();
                row = new Row();

                row.Append(CreateCell(budgetData.LineItem, CellValues.String, 1));

                foreach (var milestone in budgetData.Milestones)
                {
                    row.Append(CreateCell(milestone, CellValues.String, 1));

                }
                sheetData.Append(row);
            }
        }

        private void SetReportBudgetHeader(SheetData sheetData, List<string> reportBudgetHeaders)
        {
            Row row = new Row();

            foreach (var header in reportBudgetHeaders)
            {
                row.Append(CreateCell(header, CellValues.String, 2));
            }

            sheetData.Append(row);
        }

        private void SetActivityLogRows(SheetData sheetData, List<ActivityLogEntryData> activityLogEntry)
        {
            Row row;


            foreach (var activityLog in activityLogEntry)
            {
                row = new Row();
                row.Append(
                    CreateCell(activityLog.From, CellValues.String, 1),
                    CreateCell(activityLog.To, CellValues.String, 1),
                    CreateCell(activityLog.ElapsedTime, CellValues.String, 1),
                    CreateCell(activityLog.CumulativeTime, CellValues.String, 1),
                    CreateCell(activityLog.Depth, CellValues.String, 1),
                    CreateCell(activityLog.MudWeight, CellValues.String, 1),
                    CreateCell(activityLog.Activity, CellValues.String, 1),
                    CreateCell(activityLog.Milestone, CellValues.String, 1),
                    CreateCell(activityLog.Unplanned_Planned, CellValues.String, 1),
                    CreateCell(activityLog.HasNpt, CellValues.String, 1),
                    CreateCell(activityLog.NptReference, CellValues.String, 1),
                    CreateCell(activityLog.Description, CellValues.String, 1)
                );

                sheetData.Append(row);
            }
        }

        private void SetActivityLogHeader(SheetData sheetData, List<string> activityLogHeaders)
        {
            Row row = new Row();
            List<CellDfn> cells = new List<CellDfn>();

            foreach (var header in activityLogHeaders)
            {
                row.Append(CreateCell(header, CellValues.String, 2));
            }
            sheetData.Append(row);
        }

        private Cell CreateCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }
        private void SetColumnHeaders(SheetData sheetData)
        {
            Row row = new Row();
            for (int i = 0; i < 12; i++)
            {
                Cell cell = null;

                if (i == 6)
                {
                    cell = CreateCell("DAILY REPORT SHEET", CellValues.String, 2);
                    row.Append(cell);

                }
                else
                {
                    cell = CreateCell("", CellValues.String, 3);
                    row.Append(cell);
                }

            }
            sheetData.Append(row);
        }

        private void SetDailyReportInfo(SheetData sheetData, List<DailyReportInfo> dailyReportInfo)
        {

            Row row = new Row();
            List<CellDfn> cells = new List<CellDfn>();

            for (int i = 1; i <= dailyReportInfo.Count; i++)
            {
                
                row.Append(
                    CreateCell(dailyReportInfo[i - 1].Header, CellValues.String, 4),
                    CreateCell(dailyReportInfo[i - 1].Value, CellValues.String, 1),
                    CreateCell("", CellValues.String, 1)
                );


                if (i % 4 == 0)
                {
                    sheetData.Append(row);
                    row = new Row();
                }
            }
        }
    }
}
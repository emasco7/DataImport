using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DataImportAPI.Models;
using DataImportAPI.Utilities;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using OpenXmlPowerTools;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace DataImportAPI.Utilities.ExcelUtilities.ExcelWriterUtility
{

    public class ProductionDataSheetWriter : IProductionDataSheetWriter
    {
        public WorkbookDfn GenerateExcelSheet(ProductionSheetData productionSheetData)
        {

            WorkbookDfn wb = new WorkbookDfn();
            WorksheetDfn ws = new WorksheetDfn();
            SetWorksheetNames(ws);
            SetColumnHeaders(ws, productionSheetData);
            SetRowData(ws, productionSheetData);
            List<WorksheetDfn> worksheetDfns = new List<WorksheetDfn> { ws };
            wb.Worksheets = worksheetDfns;
            return wb;
        }

        private void SetColumnHeaders(WorksheetDfn worksheet, ProductionSheetData productionSheetData)
        {
            List<CellDfn> columnHeaders = new List<CellDfn>();
            foreach (var header in productionSheetData.ColumnHeaders)
            {
                CellDfn cell = new CellDfn();
                cell.Value = header;
                cell.Bold = true;
                columnHeaders.Add(cell);
            }
            worksheet.ColumnHeadings = columnHeaders;
        }

        private void SetWorksheetNames(WorksheetDfn worksheet)
        {
            worksheet.Name = "production data";
            worksheet.TableName = "Production Data";
        }
        private void SetRowData(WorksheetDfn worksheet, ProductionSheetData productionSheetData)
        {
            List<RowDfn> rows = new List<RowDfn>();
            foreach (var dataRow in productionSheetData.DataRows)
            {
                RowDfn row = new RowDfn();
                List<CellDfn> cells = new List<CellDfn>();

                foreach (var cellData in dataRow)
                {
                    CellDfn cell = new CellDfn();
                    cell.Value = cellData;
                    cells.Add(cell);
                }

                row.Cells = cells;
                rows.Add(row);
            }
            worksheet.Rows = rows;
        }
    }
}
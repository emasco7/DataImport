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

namespace DataImportAPI.Utilities.ExcelUtilities
{
    class BudgetSheetReader : IBudgetSheetReader
    {

        public string[] GetSheetNames(string filePath)
        {
            var sheets = SmlDataRetriever.SheetNames(filePath);
            return sheets;
        }

        public BudgetSheetData RetrieveExcelSheet(ExcelWorkSheetInfo excelWorkSheet)
        {
            var data = ReadExcel(excelWorkSheet.SheetName, excelWorkSheet.WorkBookFilePath);
            return data;
        }

        public BudgetSheetData ReadExcel(string sheetName, string filePath)
        {
            var data = new BudgetSheetData();
            WorkbookPart workbookPart;
            List<Row> rows;

            try
            {
                SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false);
                workbookPart = document.WorkbookPart;
                var sheets = workbookPart.Workbook.Descendants<Sheet>();
                var sheet = sheets.Where(x => x.Name == sheetName).FirstOrDefault();

                data.SheetName = sheet.Name;

                var workSheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
                var sheetData = workSheet.Elements<SheetData>().First();
                rows = sheetData.Elements<Row>().ToList();
                // XElement sheetData1 = SmlDataRetriever.RetrieveSheet(document,data.SheetName);
                // var rowz = sheetData1.Elements(X.Row).ToList();               
            }
            catch (System.Exception)
            {
                data.Status = "Unable to open the file";
                return data;
            }

            data.DataRows = GetDataRow(rows, workbookPart);
            return data;
        }

        public List<List<string>> GetDataRow(List<Row> rows, WorkbookPart workbookPart)
        {
            List<List<string>> DataRows = new List<List<string>>();
            if (rows.Count > 0)
            {
                for (var i = 0; i < rows.Count; i++)
                {
                    var dataRow = new List<string>();
                    DataRows.Add(dataRow);
                    var row = rows[i];
                    var RowsCells = Utility.GetRowCells(row);

                    foreach (var cell in RowsCells)
                    {
                        var text = Utility.ReadExcelCell(cell, workbookPart).Trim();
                        dataRow.Add(text);
                    }

                }
            }
            return DataRows;
        }


    }
}



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DataImportAPI.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using OpenXmlPowerTools;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace DataImportAPI.Utilities.ExcelUtilities
{
    public class ExcelReader : IExcelReader
    {

        private string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);
            return match.Value;
        }

        private int ConvertColumnNameToNumber(string columnName)
        {
            Regex alpha = new Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName)) throw new ArgumentException();
            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);
            int convertedValue = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                char letter = colLetters[i];
                int current = i == 0 ? letter - 65 : letter - 64;
                convertedValue += current * (int)Math.Pow(26, i);
            }
            return convertedValue;
        }

        public string[] GetSheetNames(string filePath)
        {
            var sheets = SmlDataRetriever.SheetNames(filePath);
            return sheets;
        }

        private IEnumerable<Cell> GetRowCells(Row row)
        {
            int currentCount = 0;
            foreach (Cell cell in row.Descendants<Cell>())
            {
                string columnName = GetColumnName(cell.CellReference);
                int currentColumnIndex = ConvertColumnNameToNumber(columnName);
                for (; currentCount < currentColumnIndex; currentCount++)
                {
                    yield return new Cell();
                }
                yield return cell;
                currentCount++;
            }
        }

        private string ReadExcelCell(Cell cell, WorkbookPart workbookPart)
        {
            var cellValue = cell.CellValue;
            var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                text = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(
                        Convert.ToInt32(cell.CellValue.Text)).InnerText;
            }
            return (text ?? string.Empty).Trim();
        }

        public ExcelData RetrieveExcelSheet(ExcelWorkSheetInfo excelWorkSheet)
        {
            var data = ReadExcel(excelWorkSheet.SheetName, excelWorkSheet.WorkBookFilePath);
            return data;
        }

       

       
        private ExcelData ReadExcel(string sheetName, string filePath)
        {
            var data = new ExcelData();
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

            data.ColumnHeaders = GetHeaderRow(rows, workbookPart);
            data.DataRows = GetDataRow(rows, workbookPart);

            return data;
        }

        private List<List<string>> GetDataRow(List<Row> rows, WorkbookPart workbookPart)
        {
            List<List<string>> DataRows = new List<List<string>>();
            if (rows.Count > 1)
            {
                for (var i = 1; i < rows.Count; i++)
                {
                    var dataRow = new List<string>();
                    DataRows.Add(dataRow);
                    var row = rows[i];
                    var RowsCells = GetRowCells(row);

                    foreach (var cell in RowsCells)
                    {
                        var text = ReadExcelCell(cell, workbookPart).Trim();
                        dataRow.Add(text);
                    }

                }
            }
            return DataRows;
        }

        private List<string> GetHeaderRow(List<Row> rows, WorkbookPart workbookPart)
        {
            List<string> columnHeaders = new List<string>();
            if (rows.Count > 0)
            {
                var row = rows[0];
                var RowsCells = GetRowCells(row);

                foreach (var cell in RowsCells)
                {
                    var text = ReadExcelCell(cell, workbookPart).Trim();
                    columnHeaders.Add(text);
                }
            }
            return columnHeaders;
        }
    }
}
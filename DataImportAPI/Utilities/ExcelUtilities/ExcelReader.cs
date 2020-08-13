using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using OpenXmlPowerTools;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace DataImportAPI.Utilities.ExcelUtilities
{
    public class ExcelReader
    {
        
        private string GetColumnName(string cellReference){
            // Create a regular expression to match the column name portion of the cell name.
            var regex = new Regex("[A-Za-z]+");
            var match= regex.Match(cellReference);
            return match.Value;
        }

         public static int ConvertColumnNameToNumber(string columnName)
        {
            Regex alpha = new Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName)) throw new ArgumentException();

            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);

            int convertedValue = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                char letter = colLetters[i];
                int current = i == 0 ? letter - 65 : letter - 64; // ASCII 'A' = 65
                convertedValue += current * (int)Math.Pow(26, i);
            }

            return convertedValue;
        }

        private  IEnumerable<Cell> GetRowCells(Row row)
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

        public ExcelData ReadExcel(IFormFile file){
            var data = new ExcelData();

            if (file==null)
            {
                data.Status="null file";
                return data;
            }
            if (file.Length <=0)
            {
                data.Status = Utility.EmptyWorkbook;
                return data;
            }
            if (file.ContentType!= Utility.OpenXmlFormats)
            {
                data.Status = Utility.InvalidFormat;
                return data;
            }

            WorkbookPart workbookPart;
            List<Row> rows;

            try
            {
                string filePath = @"C:\Users\EMASCO\Documents\PECON Prod DB.xlsx";
                SpreadsheetDocument myDoc = SpreadsheetDocument.Open(file.OpenReadStream(), false);
                // SpreadsheetDocument myDoc = SpreadsheetDocument.Open(filePath, false);

                workbookPart = myDoc.WorkbookPart;

                var sheets = workbookPart.Workbook.Descendants<Sheet>();
                var sheet = sheets.First();
                data.SheetName = sheet.Name;
	
                var workSheet = ((WorksheetPart)workbookPart
                    .GetPartById(sheet.Id)).Worksheet;
	
                var sheetData = workSheet.Elements<SheetData>().First();
                rows = sheetData.Elements<Row>().ToList();

                //var sheets= SmlDataRetriever.SheetNames(myDoc);
                //data.SheetName = sheets.First();

                
                //XElement sheetData = SmlDataRetriever.RetrieveSheet(myDoc,data.SheetName);
                //rows = sheetData.Elements<Row>().ToList();
            }
            catch (System.Exception)
            {
                
                data.Status = "Unable to open the file";
                return data;
            }

            if (rows.Count > 0)
            {
                var row = rows[0];
                var RowsCells = GetRowCells(row);
               
                foreach (var cell in RowsCells)
                {
                    var text = ReadExcelCell(cell, workbookPart).Trim();
                    data.ColumnHeaders.Add(text);
                }
            }

             if (rows.Count > 1)
            {
                for (var i = 1; i < rows.Count; i++)
                {
                    var dataRow = new List<string>();
                    data.DataRows.Add(dataRow);
                    var row = rows[i];
                    var RowsCells = GetRowCells(row);

                    
                    foreach (var cell in RowsCells)
                    {
                        var text = ReadExcelCell(cell, workbookPart).Trim();
                        dataRow.Add(text);
                    }
                   
                }
            }
	
            return data;
        }

       
    }

    
}
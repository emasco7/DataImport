using System.Collections.Generic;
using DataImportAPI.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataImportAPI.Utilities.ExcelUtilities
{
    public interface IBudgetSheetReader
    {
        List<List<string>> GetDataRow(List<Row> rows, WorkbookPart workbookPart);
        string[] GetSheetNames(string filePath);
        BudgetSheetData ReadExcel(string sheetName, string filePath);
        BudgetSheetData RetrieveExcelSheet(ExcelWorkSheetInfo excelWorkSheet);
    }
}
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataImportAPI.Utilities.ExcelUtilities
{
    public class ExcelData
    {
        public string Status { get; set; }
        public List<string> ColumnHeaders { get; set; }= new List<string>();
        public List<List<string>> DataRows { get; set; }= new List<List<string>>();
        public string SheetName { get; set; }


    }
}
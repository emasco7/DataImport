using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataImportAPI.Models
{
    public class ProductionSheetData
    {
        
        public List<string> ColumnHeaders { get; set; }= new List<string>();
        public List<List<string>> DataRows { get; set; }= new List<List<string>>();
        public string SheetName { get; set; }
    }
}
using System.Collections.Generic;

namespace DataImportAPI.Models
{
    public class BudgetSheetData
    {
        public string Status { get; set; }
        public List<List<string>> DataRows { get; set; }= new List<List<string>>();
        public string SheetName { get; set; }

    }
}
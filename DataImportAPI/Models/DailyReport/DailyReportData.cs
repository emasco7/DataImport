using System.Collections.Generic;

namespace DataImportAPI.Models.DailyReport
{
    public class DailyReportData
    {
        public string SheetName {get;set;}
        public List<DailyReportInfo> ReportInfo {get;set;}
        public List<DailyReportBudgetData> ReportBudget { get; set; } 
        public List<ActivityLogEntryData> ActivityLog { get; set; }
        public List<string> ActivityLogHeaders { get; set; }
        public List<string> ReportBudgetHeaders { get; set; }       
    }
   
}
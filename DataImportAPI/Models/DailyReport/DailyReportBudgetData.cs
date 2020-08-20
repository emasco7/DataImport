using System.Collections.Generic;

namespace DataImportAPI.Models.DailyReport
{
    public class DailyReportBudgetData
    {
      public string LineItem { get; set; }
      public List<string> Milestones {get; set;}   
    }
}
namespace DataImportAPI.Models.DailyReport
{
    public class ActivityLogEntryData
    {
       public string From { get; set; }
       public string To { get; set; }
       public string ElapsedTime { get; set; }
       public string CumulativeTime { get; set; }
       public string Depth { get; set; }
       public string MudWeight { get; set; }
       public string Activity { get; set; }
       public string Milestone { get; set; }
       public string Unplanned_Planned { get; set; }
       public string HasNpt { get; set; }
       public string NptReference { get; set; }
       public string Description { get; set; }
      
    }
}
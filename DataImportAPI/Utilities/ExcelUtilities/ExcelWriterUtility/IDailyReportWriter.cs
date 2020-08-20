using DataImportAPI.Models.DailyReport;
using OpenXmlPowerTools;

namespace DataImportAPI.Utilities.ExcelUtilities.ExcelWriterUtility
{
    public interface IDailyReportWriter
    {
        WorkbookDfn GenerateExcelSheet(DailyReportData dailyReportSheetData);
    }
}
using System;
using System.IO;
using System.Threading.Tasks;
using DataImportAPI.Models.DailyReport;
using DataImportAPI.Utilities.ExcelUtilities.ExcelWriterUtility;
using Microsoft.AspNetCore.Mvc;
using OpenXmlPowerTools;

namespace DataImportAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DailyReportDataExportController:ControllerBase
    {
        private readonly IDailyReportWriter dailyReportWriter;

        public readonly IDailyReportWriterWithStyleSheet dailyReportWriterWithStyle;

        public DailyReportDataExportController(IDailyReportWriter dailyReportWriter,IDailyReportWriterWithStyleSheet dailyReportWriterWithStyle)
        {
            this.dailyReportWriter = dailyReportWriter;
            this.dailyReportWriterWithStyle = dailyReportWriterWithStyle;      
        }

        [HttpPost]
        [Route("ExportDocument")]
        public async Task<IActionResult> DownloadProductionDataAsync(DailyReportData dailyReportData){
            if (dailyReportData==null)
            {
                throw new ArgumentNullException(nameof(dailyReportData));
            }
            else
            {
                string fileName = Path.GetTempFileName();
                var data = dailyReportWriter.GenerateExcelSheet(dailyReportData);
                SpreadsheetWriter.Write(fileName, data);

                var memory = new MemoryStream();

                using (var fileStream = new FileStream(fileName, FileMode.Open))
                {
                    await fileStream.CopyToAsync(memory);
                }  
                
                memory.Position=0;
                
                Response.Headers.Add("Content-Disposition",
                "attachment; filename=DailyReport.xlsx");
                return File(memory,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            }
        } 

        [HttpPost]
        [Route("ExportDocumentWithStyle")]
        public  IActionResult DownloadDailyReportDataAsync(DailyReportData dailyReportData){
            if (dailyReportData==null)
            {
                throw new ArgumentNullException(nameof(dailyReportData));
            }
            else
            {
               
                var data = dailyReportWriterWithStyle.GenerateExcelSheet(dailyReportData);
                
                Response.Headers.Add("Content-Disposition",
                "attachment; filename=DailyReport.xlsx");
                return File(data,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            }
        } 

        [HttpGet]
        public IActionResult GetDataModel(){
            
            return Ok(new DailyReportData());
        } 
    }
}
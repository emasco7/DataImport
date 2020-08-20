using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using DataImportAPI.Models;
using DataImportAPI.Utilities;
using DataImportAPI.Utilities.ExcelUtilities;
using DataImportAPI.Utilities.ExcelUtilities.ExcelWriterUtility;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OpenXmlPowerTools;

namespace DataImportAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProductionDataExportController:ControllerBase
    {
        private readonly IProductionDataSheetWriter productionDataSheetWriter;

        public ProductionDataExportController(IProductionDataSheetWriter productionDataSheetWriter)
        {
            this.productionDataSheetWriter=productionDataSheetWriter;
        }

        [HttpPost]
        [Route("ExportDocument")]
        public async Task<IActionResult> DownloadProductionDataAsync(ProductionSheetData productiondata){
            if (productiondata==null)
            {
                throw new ArgumentNullException(nameof(productiondata));
            }
            else
            {
                string fileName = Path.GetTempFileName();
                var data = productionDataSheetWriter.GenerateExcelSheet(productiondata);
                SpreadsheetWriter.Write(fileName, data);

                var memory = new MemoryStream();

                

                using (var fileStream = new FileStream(fileName, FileMode.Open))
                {
                    await fileStream.CopyToAsync(memory);
                }  
                
                memory.Position=0;
                
                Response.Headers.Add("Content-Disposition",
                "attachment; filename=ExcelFile.xlsx");
                return File(memory,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            }
        } 
    }
}
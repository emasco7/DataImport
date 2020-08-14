using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using DataImportAPI.Utilities;
using DataImportAPI.Utilities.ExcelUtilities;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;


namespace DataImportAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelDataImportController:ControllerBase
    {
        public ExcelDataImportController(IExcelReader excelReader)
        {
            this.excelReader = excelReader;      
        }
      
        private readonly IExcelReader excelReader;

        [HttpPost]
        [Route("UploadDocument")]
        public async Task<IActionResult> CreateCommandAsync([FromForm]ExcelFormData excelFormData) {

            if (excelFormData.File == null || excelFormData.File.Length == 0)
                return Content("File Not Selected");

            string fileExtension = Path.GetExtension(excelFormData.File.FileName);

            if (fileExtension == ".xls" || fileExtension == ".xlsx")
            {       
                string fileName = Path.GetTempFileName();

                //Console.WriteLine("TEMP file created at: " + fileName);
                using (var fileStream = new FileStream(fileName, FileMode.Create))
                {
                    await excelFormData.File.CopyToAsync(fileStream);
                }         
                var sheets = excelReader.GetSheetNames(fileName);
    
                return Ok(new{sheets,fileName});      
            }
            else
            {
                throw new FileFormatException();
            }
             
            
        }

        [HttpPost]
        [Route("RetrieveDocument")]
        public ActionResult <ExcelData> CreateExcelData(ExcelWorkSheetInfo excelWorkSheetInfo){
            if (excelWorkSheetInfo.SheetName==null||excelWorkSheetInfo.WorkBookFilePath==null)
            {
                throw new ArgumentNullException(nameof(excelWorkSheetInfo));
            }
            else
            {
              var data = excelReader.RetrieveExcelSheet(excelWorkSheetInfo);
              return Ok(data);
            }
        } 
    }
    public class ExcelFormData
    {      
        public IFormFile File { get; set; }
    }

    

}


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
    public class BudgetSheetImportController:ControllerBase
    {
        public BudgetSheetImportController(IBudgetSheetReader budgetSheetReader)
        {
            this.budgetSheetReader = budgetSheetReader;      
        }
      
        private readonly IBudgetSheetReader budgetSheetReader;

        [HttpPost]
        [Route("UploadDocument")]
        public async Task<IActionResult> CreateCommandAsync([FromForm]ExcelFormData excelFormData) {

            if (excelFormData.File == null || excelFormData.File.Length == 0)
                return Content("File Not Selected");

            string fileExtension = Path.GetExtension(excelFormData.File.FileName);

            if (fileExtension == ".xls" || fileExtension == ".xlsx")
            {       
                string fileName = Path.GetTempFileName();

                using (var fileStream = new FileStream(fileName, FileMode.Create))
                {
                    await excelFormData.File.CopyToAsync(fileStream);
                }         
                var sheets = budgetSheetReader.GetSheetNames(fileName);
    
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
              var data = budgetSheetReader.RetrieveExcelSheet(excelWorkSheetInfo);
              return Ok(data);
            }
        } 
    }
    

    

}


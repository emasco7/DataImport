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
    //[ApiController]
    public class ExcelDataImportController:ControllerBase
    {
        ExcelReader excelReader;
        ExcelData data;
        public ExcelDataImportController()
        {

        }
        [HttpPost]
        public ActionResult <ExcelData> CreateCommand([FromForm]ExcelFormData excelFormData) {

            excelReader = new ExcelReader();
            data = excelReader.ReadExcel(excelFormData.File);
            HttpContext.Session.SetObject(Utility.SessionData,data);
            return CreatedAtRoute(nameof(GetDataById), new {Id=1}, data);
        }

        // [HttpPost]
        // public async Task<IActionResult> CreateCommandAsync([FromForm]ExcelFormData excelFormData) {

        //     if (excelFormData.File == null || excelFormData.File.Length == 0)
        //         return Content("File Not Selected");
        //     string fileExtension = Path.GetExtension(excelFormData.File.FileName);
        //     if (fileExtension == ".xls" || fileExtension == ".xlsx")
        //     {
        //         var rootFolder = @"C:\Files";
        //         var fileName = excelFormData.File.FileName;
        //         var filePath = Path.Combine(rootFolder, fileName);
        //         var fileLocation = new FileInfo(filePath);

        //         using (var fileStream = new FileStream(filePath, FileMode.Create))
        //         {
        //             await excelFormData.File.CopyToAsync(fileStream);
        //         }

                 
        //     }
        //     excelReader = new ExcelReader();
        //     data = excelReader.ReadExcel(excelFormData.File);
        //     return CreatedAtRoute(nameof(GetCommandById), new {Id=1}, data);
        // }



        [HttpGet("{id}", Name="GetDataById")]
        public ActionResult <ExcelData> GetDataById(int id){
            //creat inappmemory
            
                var data = HttpContext.Session.GetObject<ISession>(Utility.SessionData);
                return Ok(data);

            
        }
    }
    public class ExcelFormData
    {
        
        public IFormFile File { get; set; }
    }
    public static class SessionExtensions
    {
        public static void SetObject(this ISession session, string key, object value)
        {
            session.SetString(key, JsonSerializer.Serialize(value));
        }

        public static T GetObject<T>(this ISession session, string key)
        {
            var value = session.GetString(key);
            return value == null ? default(T) : JsonSerializer.Deserialize<T>(value);
        }
    }
}


namespace DataImportAPI.Utilities.ExcelUtilities
{
    public interface IExcelReader
    {
        string[] GetSheetNames(string filePath);
        ExcelData RetrieveExcelSheet(ExcelWorkSheetInfo excelWorkSheet);
    }
}
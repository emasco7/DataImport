using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DataImportAPI.Models;
using DataImportAPI.Utilities;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using OpenXmlPowerTools;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace DataImportAPI.Utilities.ExcelUtilities.ExcelWriterUtility
{
    public interface IProductionDataSheetWriter
    {
        WorkbookDfn GenerateExcelSheet(ProductionSheetData productionSheetData);
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataImportAPI.Utilities
{
    public class Utility
    {
        public static string OpenXmlFormats="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public static string EmptyWorkbook="You uploaded an empty workbook";
        public static string InvalidFormat = "please upload a valid excel format";
        public static string SessionData = "excel data";
        public static string UnableToOpenFile = "Unable to open file";

        public static string GetColumnName(string cellReference)
        {
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);
            return match.Value;
        }
        public static int ConvertColumnNameToNumber(string columnName)
        {
            Regex alpha = new Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName)) throw new ArgumentException();
            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);
            int convertedValue = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                char letter = colLetters[i];
                int current = i == 0 ? letter - 65 : letter - 64;
                convertedValue += current * (int)Math.Pow(26, i);
            }
            return convertedValue;
        }
        public static IEnumerable<Cell> GetRowCells(Row row)
        {
            int currentCount = 0;
            foreach (Cell cell in row.Descendants<Cell>())
            {
                string columnName = GetColumnName(cell.CellReference);
                int currentColumnIndex = ConvertColumnNameToNumber(columnName);
                for (; currentCount < currentColumnIndex; currentCount++)
                {
                    yield return new Cell();
                }
                yield return cell;
                currentCount++;
            }
        }
        public static string ReadExcelCell(Cell cell, WorkbookPart workbookPart)
        {
            var cellValue = cell.CellValue;
            var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                text = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(
                        Convert.ToInt32(cell.CellValue.Text)).InnerText;
            }
            return (text ?? string.Empty).Trim();
        }

        public static Stylesheet GenerateStlyeSheet(){
            Stylesheet stylesheet =null;
            Fonts fonts = new Fonts(
        new Font( // Index 0 - default
            new FontSize() { Val = 10 }

        ),
        new Font( // Index 1 - header
            new FontSize() { Val = 10 },
            new Bold(),
            new Color() { Rgb = "FFFFFF" }

        ),
        new Font( // Index 0 - default
            new FontSize() { Val = 10 },
            new Bold()
        ));
            
            Fills fills = new Fills(
                new Fill(new PatternFill(){PatternType=PatternValues.None}),
                new Fill(new PatternFill(){PatternType=PatternValues.Gray125}),
                new Fill(new PatternFill(new ForegroundColor{Rgb=new HexBinaryValue(){Value="66666666"}}){PatternType=PatternValues.Solid})
            );

            Borders borders = new Borders(
                new Border(), // index 0 default
                new Border( // index 1 black border
                    new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder()
                )
            );

            CellFormats cellFormats = new CellFormats(
                new CellFormat(), // default
                new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true }, // header
                new CellFormat { FontId = 1, FillId = 2, BorderId = 0, ApplyFill = true },// header
                new CellFormat { FontId = 2, FillId = 0, BorderId = 1, ApplyBorder = true }// header
            );

            stylesheet = new Stylesheet(fonts,fills,borders,cellFormats);
            return stylesheet;
        }
    }
}
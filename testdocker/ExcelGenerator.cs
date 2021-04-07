using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;

namespace OpenXmlPocDocker
{
    public static class ExcelGenerator
    {
        private const string FlagUrl = "https://www.countryflags.com/";
        public static void CreateExcelFile(string outPutFileDirectory)
        { 
            List<Country> data = DataGenerator.GetCountries().Result;

            var dateTime = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            var fileName = Path.Combine(outPutFileDirectory, $"output_{dateTime}.xlsx");

            // Create a spreadsheet document using the file name.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            
            // Add a WorkbookPart with a workbook object to the document.  
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook(); 
            
            // Add a WorksheetPart with a worksheet to the WorkbookPart. 
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>(); 
            worksheetPart.Worksheet = new Worksheet(); 
            
            // Create a Sheets object and append it to workbook.  
            Sheets sheets = workbookpart.Workbook.AppendChild(new Sheets()); 
            
            // Create a Sheet object and append it to Sheets.
            Sheet sheet = new Sheet() { Id = workbookpart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet-1" };
            sheets.Append(sheet);
            
            //Create Column object to set properties like column width
            Columns columns = new Columns();
            columns.Append(new Column() { Min = 1, Max = 1, Width = 15, CustomWidth = true });
            columns.Append(new Column() { Min = 2, Max = 3, Width = 40, CustomWidth = true });
            worksheetPart.Worksheet.Append(columns);
            
            //Create a StyleSheet object to set properties like fonts, fills, bordres
            WorkbookStylesPart stylesPart  = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = GenerateStyleSheet();
            stylesPart.Stylesheet.Save();
            
            // Get the sheetData cell table.  
            SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
            GenerateSheetData(sheetData, data);

            // Close the document.  
            spreadsheetDocument.Close(); 


        }

        private static Stylesheet GenerateStyleSheet()
        {
            Stylesheet stylesheet = new Stylesheet();
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            
            var fonts = new Fonts() { Count = 2U};
            var fills = new Fills() { Count = 5U };
            var borders = new Borders() {Count = 1};
            var cellFormats = new CellFormats() { Count = 4U };
            
            // Create Default Row Font : Verdana Black 12
            Font rowFont = new Font();
            rowFont.Append(new FontSize() { Val = 12D });
            rowFont.Append(new Color() { Rgb = "FF000000"});
            rowFont.Append(new FontName() { Val = "Verdana" });

            // Create Header Font : Calibri White 16 Bold
            Font headerFont = new Font();
            headerFont.Append(new Bold());
            headerFont.Append(new FontSize() { Val = 16D });
            headerFont.Append(new Color() { Rgb = "FFFFFFFF"});
            headerFont.Append(new FontName() { Val = "Calibri" });
            
            fonts.Append(rowFont);
            fonts.Append(headerFont);
            
            // Create Header Fill : Dark Grey
            PatternFill headerPatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor {Rgb = "FF4D4D4D"},
                BackgroundColor = new BackgroundColor {Indexed = 64}
            };

            // Create Row Odd Fill : Light Grey
            PatternFill oddRowPatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor {Rgb = "FFEAEAEA"},
                BackgroundColor = new BackgroundColor {Indexed = 64}
            };

            // Create Row Even Fill : White
            PatternFill evenRowPatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor {Rgb = "FFFFFFFF"},
                BackgroundColor = new BackgroundColor {Indexed = 64}
            };

            fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required
            fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required
            fills.AppendChild(new Fill { PatternFill = headerPatternFill });
            fills.AppendChild(new Fill { PatternFill = oddRowPatternFill });
            fills.AppendChild(new Fill { PatternFill = evenRowPatternFill });
            
            // Create default border
            Border border1 = new Border();
            border1.Append(new LeftBorder());
            border1.Append(new RightBorder());
            border1.Append(new TopBorder());
            border1.Append(new BottomBorder());
            border1.Append(new DiagonalBorder());

            borders.Append(border1);

            cellFormats.AppendChild(new CellFormat());
            cellFormats.AppendChild(new CellFormat { FontId = 1, FillId = 2, BorderId = 0, ApplyFill = true }); // 1.header cell format
            cellFormats.AppendChild(new CellFormat { FontId = 0, FillId = 3, BorderId = 0, ApplyFill = true }); // 2.odd row cell format
            cellFormats.AppendChild(new CellFormat { FontId = 0, FillId = 4, BorderId = 0, ApplyFill = true }); // 3.even row cell format

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellFormats);

            return stylesheet;
        }
        private static void GenerateSheetData(SheetData sheetData, List<Country> data)
        {
            Row titleRow = CreatTitleeRow();
            sheetData.Append(titleRow);
            
            Row headerRow = CreateHeaderRow();
            sheetData.Append(headerRow);

            UInt32Value index = headerRow.RowIndex;
            foreach (Country country in data)
            {
                index++;
                country.name = country.name.Split('(')[0].Trim().Replace(' ', '-');
                Row row = CreateRow(index, country.name);
                sheetData.Append(row);
            }
        }
        
        private static Row CreatTitleeRow()
        {
            Row titleRow = new Row(){ RowIndex = 1 };
            titleRow.Append(CreateCellWithReference(FlagUrl, "F1", 3U));
            return titleRow;
        }
        private static Row CreateHeaderRow()
        {
            Row headerRow = new Row(){ RowIndex = 1 };
            headerRow.Append(CreateCell("Country", 1U));
            headerRow.Append(CreateCell("Flag URL By Formula", 1U));
            headerRow.Append(CreateCell("Flag URL By Code", 1U));
            return headerRow;
        }
        
        private static Row CreateRow(UInt32Value index, string text)
        {
            UInt32Value styleIndex = 2U;
            if (index % 2 != 0)
            {
                styleIndex = 3U;
            }
            Row row = new Row() { RowIndex = index };
            row.Append(CreateCell(text, styleIndex));
            row.Append(CreateCellWithFormula("LIEN_HYPERTEXTE(CONCAT(F1,A"+index+"))", styleIndex));
            row.Append(CreateCellWithFormula("LIEN_HYPERTEXTE(\""+FlagUrl+text+"\")", styleIndex));
            return row;
        }
        private static Cell CreateCell(string text, UInt32Value styleIndex)
        {
            Cell cell = new Cell
            {
                CellValue = new CellValue(text),
                DataType = ResolveCellDataTypeOnValue(text),
                StyleIndex = styleIndex
            };
            return cell;
        }
        
        private static Cell CreateCellWithFormula(string formula, UInt32Value styleIndex)
        {
            Cell cell = CreateCell("0", styleIndex);
            cell.CellFormula = new CellFormula(formula);
            return cell;
        }

        private static Cell CreateCellWithReference(string text, string reference, UInt32Value styleIndex)
        {
            Cell cell = CreateCell(text, styleIndex);
            cell.CellReference = reference;
            return cell;
        }
        private static EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            if (int.TryParse(text, out _) || double.TryParse(text, out _))
            {
                return CellValues.Number;
            }
            return CellValues.String;
        }
    }
}

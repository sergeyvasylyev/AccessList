using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Data;

namespace AccessList
{
    class ExcelData
    {
        public void ExportDataSet(DataTable table, string excelFileName, string SheetName, Boolean NewFile)
        {
            if (NewFile == true)
            {
                try
                {
                    using (var workbook = SpreadsheetDocument.Create(excelFileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                    {
                        ExportDataCommon(table, workbook, SheetName, true);
                    };
                }
                catch (Exception Ex)
                {
                    Console.WriteLine(Ex.Message);
                }
            }
            else
            {
                try
                {
                    using (var workbook = SpreadsheetDocument.Open(excelFileName, true))
                    {
                        ExportDataCommon(table, workbook, SheetName, false);
                    };
                }
                catch (Exception Ex)
                {
                    Console.WriteLine(Ex.Message);
                }
            };            
        }

        private void ExportDataCommon(DataTable table, SpreadsheetDocument workbook, string SheetName, Boolean newDoc)
        {
            if (newDoc == true)
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();
            };
            
            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

            uint sheetId = 1;
            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = SheetName };
            sheets.Append(sheet);

            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

            
            //DocumentFormat.OpenXml.Spreadsheet.Column columns2 = new DocumentFormat.OpenXml.Spreadsheet.Column();
            /*
            //Columns columns2 = new Columns();
            Columns columns1 = sheet.GetFirstChild<Columns>();
            Column column1 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 16D, CustomWidth = true };
            columns1.Append(column1);
            /*
            columns2.Append(new Column() { Min = 1, Max = 3, Width = 100, CustomWidth = true });
            columns2.Append(new Column() { Min = 4, Max = 4, Width = 100, CustomWidth = true });
            sheetData.AppendChild(columns2);
            */

            // Construct column names 
            List<String> columns = new List<string>();
            foreach (System.Data.DataColumn column in table.Columns)
            {
                columns.Add(column.ColumnName);
                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                headerRow.AppendChild(cell);
            }
            // Add the row values to the excel sheet 
            sheetData.AppendChild(headerRow);

            foreach (System.Data.DataRow dsrow in table.Rows)
            {
                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                foreach (String col in columns)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString());
                    newRow.AppendChild(cell);
                }

                sheetData.AppendChild(newRow);
            }

            /*
            DocumentFormat.OpenXml.Spreadsheet.Columns cs = sheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
            if (cs != null)
            {
                IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Column> ic = cs.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>();
                DocumentFormat.OpenXml.Spreadsheet.Column c = ic.First();
                c.Width = 100;
            }    
            */
        }
    }
}

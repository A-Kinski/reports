using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Reports
{
    class ExcelReader
    {
        //двумерный Dicrionary - содержит инфу из листа
        Dictionary<string, Dictionary<string, string>> excelSheet;
        List<List<string>> excelBuffer;

        public ExcelReader(string fileName, string sheetName = "Лист 1")
        {
            excelBuffer = new List<List<string>>();
            SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false);
            WorkbookPart wbPart = document.WorkbookPart;
            Workbook workbook = document.WorkbookPart.Workbook;
            //IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().First();//Where(s => s.Name == sheetName).FirstOrDefault();
            Sheet sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().First();
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.Id);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var excelRows = sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().ToList();

            foreach (Row r in excelRows)
            {
                List<string> buffer = new List<string>();
                foreach (Cell c in r)
                {
                    if (c.DataType != null)
                    {
                        if (c.DataType == CellValues.SharedString)
                        {
                            int id = -1;

                            if (Int32.TryParse(c.InnerText, out id))
                            {
                                SharedStringItem item = wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);

                                if (item.Text != null)
                                {
                                    buffer.Add(item.Text.Text);
                                }
                                //else if (item.InnerText != null)
                                //{
                                //    buffer.Add(item.InnerText);
                                //}
                                //else if (item.InnerXml != null)
                                //{
                                //    buffer.Add(item.InnerXml);
                                //}
                            }
                        }
                    }
                    else
                    {
                        buffer.Add(c.CellValue.Text);
                    }
                }
                excelBuffer.Add(buffer);
            }
            makeDictionary(excelBuffer);
            document.Close();
        }

        private void makeDictionary(List<List<string>> list)
        {
            excelSheet = new Dictionary<string, Dictionary<string, string>>();
            List<string> row = new List<string>();
            for (int i = 1; i < list[0].Count; i++)
                row.Add(list[0][i]);
            
            List<string> column = new List<string>();
            for (int i = 1; i < list.Count; i++)
                column.Add(list[i][0]);

            for (int i = 0; i < row.Count; i++)
            {
                Dictionary<string, string> buffer = new Dictionary<string, string>();
                for (int j = 0; j < column.Count; j++)
                {                    
                    buffer.Add(column[j], list[j+1][i+1]);
                }
                excelSheet.Add(row[i], buffer);
            }
        }
    }
}

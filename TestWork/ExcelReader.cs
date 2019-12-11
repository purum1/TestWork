using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MyExcelReader.Elements;

namespace MyExcelReader
{
    class ExcelReader
    {
        private List<Product> products;
        private Product product;
        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
        public List<Product> openExcel()
        {
            products = new List<Product>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("C:\\Price.xlsx", false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                bool result1;
                int code = 0;
                int circle = 1;
                foreach (Row r in sheetData.Elements<Row>())
                {
                    product = new Product();
                    foreach (Cell cell in r.Elements<Cell>())
                    {
                        code = 0;
                        result1 = int.TryParse(cell.InnerText, out code);
                        if (cell.DataType != null && circle == 1)
                            break;
                        if (circle <= 6)
                        {
                            string cellValue = "";
                            
                                if (cell.CellValue == null)
                                    break;
                                if (cell.DataType != null)
                                {
                                    if (cell.DataType == CellValues.SharedString)
                                    {
                                        int id = -1;
                                        if (Int32.TryParse(cell.InnerText, out id))
                                        {
                                            SharedStringItem item = GetSharedStringItemById(workbookPart, id);
                                            if (item.Text != null)
                                            {
                                                cellValue = item.Text.Text;
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                cellValue = item.InnerText;
                                            }
                                            else if (item.InnerXml != null)
                                            {
                                                cellValue = item.InnerXml;
                                            }
                                        }
                                    }
                                }
                            
                            if (cell.DataType == null)
                                cellValue = code.ToString();
                            switch (circle)
                            {
                                case 1:
                                    product.Code = cellValue;
                                    break;
                                case 2:
                                    product.Article = cellValue;
                                    break;
                                case 3:
                                    product.Name = cellValue;
                                    break;
                                case 4:
                                    product.Factory = cellValue;
                                    break;
                                case 5:
                                    product.Unit = cellValue;
                                    break;
                                case 6:
                                    product.Coast = cellValue;
                                    break;
                            }
                            circle++;
                        }
                        
                    }
                    if (product.Code != null)
                        products.Add(product);    
                    circle = 1;
                }
            }
            Console.WriteLine("End...");
            return products;
        }
    }
}

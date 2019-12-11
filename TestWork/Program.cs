using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyExcelReader.Elements;

namespace MyExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelReader excelReader = new ExcelReader();
            List<Product> products = excelReader.openExcel();
            CreaterTXT createrTXT = new CreaterTXT();
            createrTXT.createTXT(products);

        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using MyExcelReader.Elements;

namespace MyExcelReader
{
    class CreaterTXT
    {
        public CreaterTXT()
        {

        }
        public void createTXT(List<Product> products)
        {
            string writePath = @"C:\EndParsing.txt";
            try
            {
                using (StreamWriter sw = new StreamWriter(writePath, false, System.Text.Encoding.Default))
                {
                    foreach (var p in products)
                    {
                        sw.WriteLine("Код: " + p.Code + ", Артикул: " + p.Article + ", Наименование: " + p.Name + ", Пр-ль: " + p.Factory + ", Ед.изм.: " + p.Unit + ", Цена: " + p.Coast + "");
                    }
                }
                Console.WriteLine("Запись выполнена");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}

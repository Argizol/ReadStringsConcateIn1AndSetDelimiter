using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadAllLinesConcatInOneStringAndSetDelimiter
{
    internal class ReadExcel
    {
        public static void GetDataFromExcel()
        {           

            string path = @"D:/Загрузки работа/";

            Console.WriteLine($"Insert name of Excel with extension \n >>");

            path = Path.Combine(path, Console.ReadLine()!);

            using StreamWriter sw = new StreamWriter(@"D:/Загрузки работа/data/Текстовый документ.txt");
            FileInfo file = new FileInfo(path);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int limiter = worksheet.Dimension.End.Row;

                for (int i = 2; i < limiter; i++)
                {
                    var PC = worksheet.Cells[$"AV{i}"].Value is not null ? worksheet.Cells[$"AV{i}"].Value : 0.0;
                    var PPDL = worksheet.Cells[$"X{i}"].Value is not null ? worksheet.Cells[$"X{i}"].Value : 0.0;
                    var PPD = worksheet.Cells[$"Y{i}"].Value is not null ? worksheet.Cells[$"Y{i}"].Value : 0.0;
                    var IsGM = worksheet.Cells[$"AY{i}"].Value;

                    //Пишем id промо в файл если соответствует условиям
                    if ((double)PC < 20.0 && (double)PPD > (double)PPDL && IsGM.Equals("yes"))
                    {
                        sw.WriteLine(worksheet.Cells[$"A{i}"].Value);                       
                    }                    
                }
            }
        }
    }
}

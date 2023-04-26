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
            int count = 2;
            int PC = 0;
            int PPDL = 0;
            int PPD = 0;

            string path = @"D:/Загрузки работа/data/Promo Approval Status_Q3.xlsx";

            //Console.WriteLine($"Insert name of Excel with extension \n >>");

            //path = Path.Combine(path, Console.ReadLine()!);

            using StreamWriter sw = new StreamWriter(@"D:/Загрузки работа/data/Текстовый документ.txt");
            FileInfo file = new FileInfo(path);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                while (worksheet.Cells[$"A{count}"] is not null) {


                    //Чтение ячейки листа
                    if(Int32.TryParse(worksheet.Cells[$"AV{count}"].Value.ToString(), out int PCvalue))
                    {
                        PC = PCvalue;
                    }
                    if (Int32.TryParse(worksheet.Cells[$"X{count}"].Value.ToString(), out int PPDLvalue))
                    {
                        PPDL = PPDLvalue;
                    }
                    if (Int32.TryParse(worksheet.Cells[$"Y{count}"].Value.ToString(), out int PPDvalue))
                    {
                        PPD = PPDvalue;
                    }
                    
                    string IsGM = worksheet.Cells[$"AY{count}"].Value.ToString()!;

                    //Пишем id промо в файл если соответствует условиям
                    if (PC > 20 && PPD > PPDL && IsGM.Equals("yes"))
                    {
                        sw.WriteLine(worksheet.Cells[$"A{count}"].Value);
                    }
                    count++;
                }     
            }
        }
    }
}

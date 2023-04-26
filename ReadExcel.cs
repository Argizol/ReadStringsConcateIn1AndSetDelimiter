using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

            Console.WriteLine("Insert path to the Excel file with name and extension");
            Console.Write(">> ");

            path = Path.Combine(path,Console.ReadLine());
           
            using StreamWriter sw = new StreamWriter(@"result.txt");
            StringBuilder sb = new StringBuilder();
            FileInfo file = new FileInfo(path);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet is not null)
                {
                    int limiter = worksheet.Dimension.End.Row;
                    for (int i = 2; i < limiter; i++)
                    {
                        //Получаем значения, если для числового значения null меняем на 0

                        double PC = (double?)worksheet.Cells[$"AV{i}"].Value ?? 0.0;
                        double PPDL = (double?)worksheet.Cells[$"X{i}"].Value ?? 0.0;
                        double PPD = (double?)worksheet.Cells[$"Y{i}"].Value ?? 0.0;

                        string IsGM = worksheet.Cells[$"AY{i}"].Value.ToString() ?? " ";

                        //Собираем строку id промо если соответствует условиям
                        if (PC >= 20.0 && PPD <= PPDL && IsGM.Equals("yes"))
                        {
                            sb.Append(worksheet.Cells[$"A{i}"].Value).Append(';');
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No worksheets in file");
                    return;
                }

                //Пишем строку с id промо в файл, разделитель ;
                sw.Write(sb.ToString().TrimEnd(';'));                
            }
        }
    }
}

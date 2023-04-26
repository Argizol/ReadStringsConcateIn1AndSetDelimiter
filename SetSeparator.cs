using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadAllLinesConcatInOneStringAndSetDelimiter
{
    internal class SetSeparator
    {
        public static void SetSeparatorForDataFromTxt() {
            string path = @"D:/Загрузки работа/data/Текстовый документ.txt";
            using StreamReader sr = new StreamReader(path);
            using StreamWriter sw = new StreamWriter(@"D:/Загрузки работа/data/result.txt");
            StringBuilder sb = new StringBuilder();
            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine()!;
                if (line != null)
                    sb.Append(line + ';');
                else continue;

            }
            sw.Write(sb.ToString().TrimEnd(';'));
        }
    }
}

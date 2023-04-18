using System.Text;

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

using OfficeOpenXml;
using System.Text;
using System.IO;



namespace ReadAllLinesConcatInOneStringAndSetDelimiter
{

    internal class ReadExcel
    {

        public static void GetDataFromExcel()
        {

            string path = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = @"c:\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    path = openFileDialog.FileName;

                }
            }

            using StreamWriter sw = new StreamWriter(@"result.txt");
            StringBuilder sb = new StringBuilder();
            FileInfo file = new FileInfo(path);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet is null)
                {
                    Console.WriteLine("No worksheets in file");
                    return;
                }

                int limiter = worksheet.Dimension.End.Row;
                for (int i = 2; i < limiter; i++)
                {
                    //Получаем значения, если для числового значения null меняем на 0

                    double PC = (double?)worksheet.Cells[$"AW{i}"].Value ?? 0.0;
                    double PPDL = (double?)worksheet.Cells[$"X{i}"].Value ?? 0.0;
                    double PPD = (double?)worksheet.Cells[$"Y{i}"].Value ?? 0.0;

                    var IsGM = (string)worksheet.Cells[$"AZ{i}"].Value;

                    //Собираем строку id промо если соответствует условиям
                    if ((PC >= 20.0 && PPD <= PPDL) && IsGM.Equals("yes"))
                    {
                        sb.Append(worksheet.Cells[$"A{i}"].Value).Append(';');
                    }
                }

                //Пишем строку с id промо в файл, разделитель ;
                sw.Write(sb.ToString().TrimEnd(';'));
            }
        }
    }
}

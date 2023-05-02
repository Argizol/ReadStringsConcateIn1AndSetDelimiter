using ReadAllLinesConcatInOneStringAndSetDelimiter;
using System.Diagnostics;



internal class Program
{
    [STAThread] 
    private static void Main(string[] args)
    {
        ReadExcel.GetDataFromExcel();
        Process.Start(new ProcessStartInfo { FileName = @"result.txt", UseShellExecute = true });
    }
}
using ReadAllLinesConcatInOneStringAndSetDelimiter;
using System.Diagnostics;



ReadExcel.GetDataFromExcel();
Process.Start(new ProcessStartInfo {FileName = @"result.txt", UseShellExecute = true });


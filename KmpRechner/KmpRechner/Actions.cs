using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace KmpRechner;

public static class Actions
{
    [SupportedOSPlatform("windows")]
    public static void ConvertToXlsx()
    {
        Console.WriteLine("Start");
        string userPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        Excel.Application xlApp = new();
        Excel.Workbook book = xlApp.Workbooks.Open(@$"{userPath}\Desktop\Posturographie.xlsx");
        Excel.Worksheet sheet = book.Worksheets[1];

        string[] files = Directory.GetFiles(@$"{userPath}\Desktop").Where(f => Regex.IsMatch(f, @"_xy\.txt$")).OrderBy(c => c).ToArray();
        foreach (string file in files)
        {
            string[] content = File.ReadAllLines(file);
            List<(double X, double Y)> values = new();
            foreach (string line in content)
            {
                string[] nContent = line.Split('\t').Skip(6).ToArray();
                double x = double.Parse(nContent[0].Replace(',', '.'));
                double y = double.Parse(nContent[1].Replace(',', '.'));
                values.Add(new(x, y));
            }

            double result = 0;
            for (int i = 1; i < values.Count; i++)
            {
                result += Math.Pow(Math.Pow(values[i].X - values[i - 1].X, 2) + Math.Pow(values[i].Y - values[i - 1].Y, 2), 0.5);
            }

            bool breakNow = false;
            for (int i = 20; i <= 1159; i++)
            {
                for (int j = 11; j <= 13; j++)
                {
                    if (string.IsNullOrEmpty(sheet.Cells[i, j]?.Value?.ToString()))
                    {
                        sheet.Cells[i, j] = result.ToString(CultureInfo.InvariantCulture);
                        breakNow = true;
                        break;
                    }
                }

                if (breakNow)
                {
                    break;
                }
            }
        }

        book.Save();

        book.Close();
        xlApp.Quit();

        Marshal.ReleaseComObject(sheet);
        Marshal.ReleaseComObject(book);
        Marshal.ReleaseComObject(xlApp);
        Console.WriteLine("Finished");
    }
}

using System.Globalization;
using System.IO;
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
        string userPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        object misValue = System.Reflection.Missing.Value;
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook book = xlApp.Workbooks.Open(@$"{userPath}\Desktop\Posturographie.xlsx");
        Excel.Worksheet sheet = book.Worksheets[2];

        foreach (string file in Directory.GetFiles(@$"{userPath}\Desktop").OrderBy(c => c))
        {
            if (Regex.IsMatch(file, @"_xy\.txt$"))
            {
                string[] content = File.ReadAllLines(file);
                for (int i = 0; i < content.Length; i++)
                {
                    string[] nContent = content[i].Split("\t").Skip(6).ToArray();
                    sheet.Cells[i + 3, 3] = nContent[0].Replace(',', '.');
                    sheet.Cells[i + 3, 4] = nContent[1].Replace(',', '.');
                }
            }

            string result = sheet.Cells[4, 7].Calculate().ToString();
            Console.WriteLine(sheet.Cells[4, 7].Text.ToString());
            Excel.Worksheet sheet2 = book.Worksheets[1];
            for (int i = 20; i <= 1159; i++)
            {
                for (int j = 11; j <= 13; j++)
                {
                    if (string.IsNullOrEmpty(sheet2.Cells[i, j]?.Value?.ToString()))
                    {
                        sheet2.Cells[j, i] = result;
                    }
                }
            }
        }

        book.Save();
        
        book.Close();
        xlApp.Quit();

        Marshal.ReleaseComObject(sheet);
        Marshal.ReleaseComObject(book);
        Marshal.ReleaseComObject(xlApp);
    }
}
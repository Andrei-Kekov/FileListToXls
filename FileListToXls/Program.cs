using OfficeOpenXml;
using System.IO;
using System.Linq;

namespace FileListToXls
{
    public static class Program
    {
        private static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Files");
            var files = Directory.GetFiles(args[0]);

            for (int i = 0, j = 1; i < files.Length; i++, j++)
            {
                sheet.Cells[j, 1].Value = Path.GetFileName(files[i]);
            }

            Console.WriteLine("Saving...");
            Save(package);
            Console.WriteLine("Saved.");
            Console.WriteLine($"{files.Length} entries total.");
        }

        private static void Save(ExcelPackage package)
        {
            byte[] data = package.GetAsByteArray();
            File.WriteAllBytes("Files.xlsx", data);
        }
    }
    
}
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"C:\Users\mackn\Documents\C#\ExcelDemo\demo.xlsx");
            var people = GetSetupData();
            await SaveExcelFileAsync(people, file);
            List<Person> peopleFromExcel = await LoadExcelFile(file);
            foreach(var p in peopleFromExcel)
            {
                Console.WriteLine($"{p.Id} {p.FirstName} {p.LastName}");
            }
            // graphs
            // imaging
            // charts
        }

        private static async Task<List<Person>> LoadExcelFile(FileInfo file)
        {
            List<Person> output = new();
            // todo check file
            // make sure it's xls or xlsx
            // then...

            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];
            int row = 3;
            int col = 1;

            while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
            {
                Person p = new();
                p.Id = int.Parse(ws.Cells[row, col].Value.ToString());
                p.FirstName = ws.Cells[row, col + 1].Value.ToString();
                p.LastName = ws.Cells[row, col + 2].Value.ToString();
                output.Add(p);
                row += 1;
            }

            return output;
        }

        private static async Task SaveExcelFileAsync(List<Person> people, FileInfo file)
        {
            DeleteIfExists(file);
            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("MainReport");
            var range = ws.Cells["A2"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            // Formats header
            ws.Cells["A1"].Value = "Our Cool Report";
            ws.Cells["A1:C1"].Merge = true;
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

            ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(2).Style.Font.Bold = true;
            ws.Column(3).Width = 20;

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists) file.Delete();
        }

        private static List<Person> GetSetupData()
        {
            List<Person> people = new()
            {
                new() { Id = 1, FirstName = "Mackenzie", LastName = "Weaver" },
                new() { Id = 2, FirstName = "Joe", LastName = "Weaver" },
                new() { Id = 3, FirstName = "Dorinda", LastName = "Weaver" }
            };

            return people;
        }
    }
}

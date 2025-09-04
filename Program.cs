using System.Globalization;
using System.IO;
using CsvHelper;
using ClosedXML.Excel;
using CsvHelper.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;

class Program
{
    static void Main(string[] args)
    {

        var appConfig = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

        string csvPath = appConfig["CsvPath"];
        string excelPath = appConfig["ExcelPath"];
        string sheetName = appConfig["SheetName"];


        using var reader = new StreamReader(csvPath);
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = ";",
            HasHeaderRecord = true,
            IgnoreBlankLines = true,
            BadDataFound = null
        };
        using var csv = new CsvReader(reader, config);
        var records = csv.GetRecords<dynamic>();

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);

        int row = 1;
        foreach (var record in records)
        {
            var dict = (IDictionary<string, object>)record;
            int col = 1;
            foreach (var kvp in dict)
            {
                if (row == 1)
                    worksheet.Cell(row, col).Value = kvp.Key;
                worksheet.Cell(row + 1, col).Value = kvp.Value?.ToString() ?? "";
                col++;
            }
            row++;
        }

        workbook.SaveAs(excelPath);
        Console.WriteLine($"File Excel salvato in: {excelPath}");
    }
}
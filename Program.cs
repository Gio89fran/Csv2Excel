using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;

class Program
{
    static void Main(string[] args)
    {
        var appConfig = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

        string csvPath = appConfig["CsvPath"] ?? "";
        string excelPath = appConfig["ExcelPath"] ?? "";
        string sheetName = appConfig["SheetName"] ?? "";
        string onlineFileName = appConfig["OnlineFileName"] ?? "";
        string iQuoteConsole = appConfig["IQuoteConsole"] ?? "";

        //prima scarico il file csv
        var downloadedFile = new List<string>();

        Program.DownloadAsync(onlineFileName, csvPath, downloadedFile, null);

        //se non ho scaricato nulla, esco
        if (downloadedFile.Count == 0)
        {
            Console.WriteLine("Nessun file scaricato. Terminazione del programma.");
            return;
        }

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

        //una volta salvato il file, posso eliminare il file csv
        try
        {
            csv.Dispose();
            File.Delete(csvPath);
        }
        catch (Exception)
        {

        }
        //quando finito, lancio la procedura di importazione
        try
        {
            ProcessStartInfo psi = new ProcessStartInfo() { FileName = iQuoteConsole, Arguments = "Remote -j ImportData Hendi", Verb = "runas", UseShellExecute = false };
            Process p = new Process() { StartInfo = psi };
            p.Start();
            p.WaitForExit();
        }
        catch (Exception)
        {
            
        }
    }

    internal static void DownloadAsync(string url, string localPath, List<string> downloadedFile, Guid? threadId)
    {
        using (var client = new WebClient())
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                client.DownloadFile(url, localPath);
                downloadedFile.Add(localPath);
            }
            catch (Exception)
            {
                //if (threadId != null)
                //    new DataExchange.DataExchangeWriteLog(threadId) { ProcessName = "Downloading File: " + Path.GetFileName(localPath), ErrorID = 5495, ErrorMessage = ex.Message }.WriteToLog();
            }

        }
    }
}
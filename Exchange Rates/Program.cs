using GemBox.Spreadsheet;
using System.Text.Json;

internal class Program
{
    class ExchangeRate
    {
        public string Cur_Abbreviation { get; set; }
        public decimal Cur_OfficialRate { get; set; }
    }

    private static readonly string baseUrl = "https://api.nbrb.by/";
    private static readonly string[] currencyAbbreviations = { "USD", "EUR" };

    static async Task Main(string[] args)
    {
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Console.WriteLine("Welcome! Here you can find exchange rates for USD and EUR for specific dates.");

        while (true)
        {
            Console.WriteLine("\nPlease select an option:\n1. Input date manually in console.\n2. Read dates from an Excel file.");

            string option = Console.ReadLine();

            switch (option)
            {
                case "1":
                    await InputDateAndDisplayExchangeRates();
                    break;
                case "2":
                    await InputExcelPathAndWriteExchangeRates();
                    break;
                default:
                    Console.WriteLine("Please select a valid option.");
                    break;
            }
        }
    }

    private static async Task<List<ExchangeRate>?> GetExchangeRateForDate(string inputDate)
    {
        using (HttpClient client = new HttpClient())
        {
            // URL to retrieve every currency exchange rate on the given date
            string completeUrl = $"{baseUrl}ExRates/Rates?onDate={inputDate}&Periodicity=0";

            try
            {
                HttpResponseMessage response = await client.GetAsync(completeUrl);
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();

                try
                {
                    List<ExchangeRate>? responseExchangeRates = JsonSerializer.Deserialize<List<ExchangeRate>>(responseBody);
                    
                    // Filter exchange rates to selected currencies
                    List<ExchangeRate>? filteredExchangeRates = responseExchangeRates?
                        .Where(rate => currencyAbbreviations.Contains(rate.Cur_Abbreviation))
                        .ToList();
                    return filteredExchangeRates;
                }
                catch (JsonException ex)
                {
                    Console.WriteLine($"JSON deserialization error: {ex.Message}");
                }
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
        return null;
    }

    private static async Task InputDateAndDisplayExchangeRates()
    {
        string inputDate = InputDate();
        List<ExchangeRate>? exchangeRates = await GetExchangeRateForDate(inputDate);
        DisplayExchangeRates(exchangeRates);
    }

    private static async Task InputExcelPathAndWriteExchangeRates()
    {
        Console.Write("Enter the full path to the Excel file: ");
        string excelFilePath = Console.ReadLine();

        if (File.Exists(excelFilePath))
        {
            await ReadDatesFromExcelAndWriteExchangeRates(excelFilePath);
        }
        else
        {
            Console.WriteLine("Path is not valid.");
        }
    }

    private static async Task ReadDatesFromExcelAndWriteExchangeRates(string path)
    {
        ExcelFile file = ExcelFile.Load(path);
        ExcelWorksheet ws = file.Worksheets[0];

        // Assuming that the dates are in the first column
        int dateColumnIndex = 0;

        // Offset for writing in the exchange rates
        int columnOffset = 1;

        // Insert header row with content names (assuming the file has only dates)
        ws.Rows.InsertEmpty(0, 1);
        ws.Cells[0, dateColumnIndex].Value = "Date";
        for (int i = 0; i < currencyAbbreviations.Length; i++)
        {
            ws.Cells[0, i + columnOffset].Value = currencyAbbreviations[i];
        }

        // Start from 1st row to skip header
        for (int rowIndex = 1; rowIndex < ws.Rows.Count; rowIndex++)
        {
            string dateValue = ws.Cells[rowIndex, dateColumnIndex].Value?.ToString();
            string parsedDate = ParseDate(dateValue);

            // If the date is invalid, skip the row
            if (!string.IsNullOrEmpty(parsedDate))
            {
                List<ExchangeRate>? exchangeRates = await GetExchangeRateForDate(parsedDate);

                if (exchangeRates != null)
                {
                    foreach (var exRate in exchangeRates)
                    {
                        decimal exchangeRate = exRate.Cur_OfficialRate;

                        // Get column index to insert the value
                        int columnIndex = Array.IndexOf(currencyAbbreviations, exRate.Cur_Abbreviation) + columnOffset;

                        ws.Cells[rowIndex, columnIndex].Value = exchangeRate;
                    }
                }
            }
        }

        file.Save(path);
        Console.WriteLine("File modified successfully.");
    }

    private static void DisplayExchangeRates(List<ExchangeRate>? exchangeRates)
    {
        foreach (var exRate in exchangeRates)
        {
            Console.WriteLine($"Exchange rate for {exRate.Cur_Abbreviation} is {exRate.Cur_OfficialRate}");
        }
    }

    private static string InputDate()
    {
        while (true)
        {
            Console.Write("Enter a date: ");
            string date = ParseDate(Console.ReadLine());
            if (!string.IsNullOrEmpty(date))
            {
                return date;
            }
            else
            {
                Console.WriteLine("Invalid input. Please try again.");
            }
        }
    }

    private static string ParseDate(string inputDate)
    {
        if (DateTime.TryParse(inputDate, out DateTime parsedDate))
        {
            return parsedDate.ToUniversalTime().ToString("yyyy-MM-dd");
        }
        else
        {
            // Parsing failed
            return string.Empty;
        }
    }
}
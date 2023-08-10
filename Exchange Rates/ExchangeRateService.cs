using GemBox.Spreadsheet;
using System.Text.Json;

namespace ExchangeRates
{
    public class ExchangeRateService
    {
        private readonly string _baseUrl = "https://api.nbrb.by/";
        private readonly string[] _currencyAbbreviations = { "USD", "EUR" };

        public async Task<List<ExchangeRate>?> GetExchangeRateForDate(string inputDate)
        {
            using (HttpClient client = new HttpClient())
            {
                // URL to retrieve every currency exchange rate on the given date
                var completeUrl = $"{_baseUrl}ExRates/Rates?onDate={inputDate}&Periodicity=0";

                try
                {
                    HttpResponseMessage response = await client.GetAsync(completeUrl);
                    response.EnsureSuccessStatusCode();

                    var responseBody = await response.Content.ReadAsStringAsync();
                    List<ExchangeRate>? responseExchangeRates = JsonSerializer.Deserialize<List<ExchangeRate>>(responseBody);

                    // Filter exchange rates to selected currencies
                    List<ExchangeRate>? filteredExchangeRates = responseExchangeRates?
                        .Where(rate => _currencyAbbreviations.Contains(rate.Abbreviation))
                        .ToList();

                    return filteredExchangeRates;
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
                catch (JsonException ex)
                {
                    Console.WriteLine($"JSON deserialization error: {ex.Message}");
                }
            }
            return null;
        }

        public async Task<List<DateExchangeRate>> ReadExcelAndGetExchangeRatesWithDates(string path)
        {
            ExcelFile file = ExcelFile.Load(path);
            ExcelWorksheet ws = file.Worksheets[0];

            List<DateExchangeRate> dateExchangeRates = new List<DateExchangeRate>();

            // Assuming the dates are in the first column
            foreach (var cell in ws.Columns[0].Cells)
            {
                var date = ParseDate(cell.Value?.ToString());
                if (!string.IsNullOrEmpty(date))
                {
                    var exchangeRates = await GetExchangeRateForDate(date);

                    // Retrieve official rates for currencies
                    var usdRate = exchangeRates?.FirstOrDefault(rate => rate.Abbreviation == "USD")?.OfficialRate ?? 0.0m;
                    var eurRate = exchangeRates?.FirstOrDefault(rate => rate.Abbreviation == "EUR")?.OfficialRate ?? 0.0m;

                    dateExchangeRates.Add(new DateExchangeRate { Date = date, EURRate = eurRate, USDRate = usdRate });
                }
            }
            return dateExchangeRates;
        }

        public void WriteExchangeRatesToExcelFile(string path, List<DateExchangeRate> dateExchangeRates)
        {
            ExcelFile file = ExcelFile.Load(path);

            // Create a new worksheet or clear it if it already exists
            var worksheetName = "Exchange Rates";
            var ws = file.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
            if (ws == null)
            {
                ws = file.Worksheets.Add(worksheetName);
            }
            else
            {
                ws.Clear();
            }

            // Define column indexes for specific contents
            int dateIndex = 0;
            int usdIndex = 1;
            int eurIndex = 2;

            // Add header row
            ws.Rows.InsertEmpty(0, 1);
            ws.Cells[0, dateIndex].Value = "Date";
            ws.Cells[0, usdIndex].Value = "USD";
            ws.Cells[0, eurIndex].Value = "EUR";

            // Write data to Excel file
            for (int i = 1; i < dateExchangeRates.Count + 1; i++)
            {
                var rate = dateExchangeRates[i - 1];
                ws.Cells[i, dateIndex].Value = rate.Date;
                ws.Cells[i, usdIndex].Value = rate.USDRate;
                ws.Cells[i, eurIndex].Value = rate.EURRate;
            }

            file.Save(path);
        }

        public string ParseDate(string inputDate)
        {
            if (DateTime.TryParse(inputDate, out DateTime parsedDate))
            {
                return parsedDate.ToString("yyyy-MM-dd");
            }
            else
            {
                // Parsing failed
                return string.Empty;
            }
        }
    }
}

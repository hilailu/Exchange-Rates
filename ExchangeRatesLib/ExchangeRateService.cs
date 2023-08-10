using GemBox.Spreadsheet;
using System.Text.Json;

namespace ExchangeRatesLib
{
    public class ExchangeRateService
    {
        private readonly string[] _currencyAbbreviations = { Constants.USD, Constants.EUR };

        public async Task<List<ExchangeRate>?> GetExchangeRateForDate(string inputDate)
        {
            using (HttpClient client = new HttpClient())
            {
                // URL to retrieve every currency exchange rate on the given date
                var completeUrl = $"{Constants.BaseApiUrl}ExRates/Rates?onDate={inputDate}&Periodicity=0";

                try
                {
                    var response = await client.GetAsync(completeUrl);
                    response.EnsureSuccessStatusCode();

                    var responseBody = await response.Content.ReadAsStringAsync();
                    var responseExchangeRates = JsonSerializer.Deserialize<List<ExchangeRate>>(responseBody);

                    // Filter exchange rates to selected currencies
                    var filteredExchangeRates = responseExchangeRates?
                        .Where(rate => _currencyAbbreviations.Contains(rate.Abbreviation))
                        .ToList();

                    return filteredExchangeRates;
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"HTTP error: {ex.Message}");
                }
                catch (JsonException ex)
                {
                    Console.WriteLine($"JSON deserialization error: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occured: {ex.Message}");
                }
            }
            return null;
        }

        public async Task<List<DateExchangeRate>> ReadExcelAndGetExchangeRatesWithDates(string path)
        {
            List<DateExchangeRate> dateExchangeRates = new List<DateExchangeRate>();

            if (!IsValidExcelFile(path))
            {
                return dateExchangeRates;
            }

            var file = ExcelFile.Load(path);
            var ws = file.Worksheets[0];

            // Assuming the dates are in the first column
            foreach (var cell in ws.Columns[(int)Indexes.Date].Cells)
            {
                var date = ParseDate(cell.Value?.ToString());
                if (!string.IsNullOrEmpty(date))
                {
                    var exchangeRates = await GetExchangeRateForDate(date);

                    // Retrieve official rates for currencies
                    var usdRate = exchangeRates?.FirstOrDefault(rate => rate.Abbreviation == Constants.USD)?.OfficialRate ?? 0.0m;
                    var eurRate = exchangeRates?.FirstOrDefault(rate => rate.Abbreviation == Constants.EUR)?.OfficialRate ?? 0.0m;

                    dateExchangeRates.Add(new DateExchangeRate { Date = date, EURRate = eurRate, USDRate = usdRate });
                }
            }
            return dateExchangeRates;
        }

        public bool IsWriteExchangeRatesToExcelFileSuccessful(string path, List<DateExchangeRate> dateExchangeRates)
        {
            if (!IsValidExcelFile(path) || dateExchangeRates.Count == 0)
            {
                return false;
            }

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

            // Add header row
            ws.Rows.InsertEmpty(0, 1);
            ws.Cells[0, (int)Indexes.Date].Value = Constants.Date;
            ws.Cells[0, (int)Indexes.USD].Value = Constants.USD;
            ws.Cells[0, (int)Indexes.EUR].Value = Constants.EUR;

            // Write data to Excel file
            for (int i = 1; i < dateExchangeRates.Count + 1; i++)
            {
                var rate = dateExchangeRates[i - 1];
                ws.Cells[i, (int)Indexes.Date].Value = rate.Date;
                ws.Cells[i, (int)Indexes.USD].Value = rate.USDRate;
                ws.Cells[i, (int)Indexes.EUR].Value = rate.EURRate;
            }

            file.Save(path);
            return true;
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

        public bool IsValidExcelFile(string path, bool isTemp = false)
            => (File.Exists(path) || isTemp) && Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase);
    }
}

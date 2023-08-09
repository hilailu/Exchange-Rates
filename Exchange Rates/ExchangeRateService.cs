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

        public async Task ReadDatesFromExcelAndWriteExchangeRates(string path)
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
            for (int i = 0; i < _currencyAbbreviations.Length; i++)
            {
                ws.Cells[0, i + columnOffset].Value = _currencyAbbreviations[i];
            }

            // Start from 1st row to skip header
            for (int rowIndex = 1; rowIndex < ws.Rows.Count; rowIndex++)
            {
                var dateValue = ws.Cells[rowIndex, dateColumnIndex].Value?.ToString();
                var parsedDate = ParseDate(dateValue);

                // If the date is invalid, skip the row
                if (!string.IsNullOrEmpty(parsedDate))
                {
                    List<ExchangeRate>? exchangeRates = await GetExchangeRateForDate(parsedDate);

                    if (exchangeRates != null)
                    {
                        foreach (var exRate in exchangeRates)
                        {
                            var exchangeRate = exRate.OfficialRate;

                            // Get column index to insert the value
                            var columnIndex = Array.IndexOf(_currencyAbbreviations, exRate.Abbreviation) + columnOffset;

                            ws.Cells[rowIndex, columnIndex].Value = exchangeRate;
                        }
                    }
                }
            }

            file.Save(path);
        }

        public string ParseDate(string inputDate)
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
}

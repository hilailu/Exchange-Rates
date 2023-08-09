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
                    // Excel logic
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
            if (DateTime.TryParse(Console.ReadLine(), out DateTime parsedDate))
            {
                return parsedDate.ToUniversalTime().ToString("yyyy-MM-dd");
            }
            else
            {
                Console.WriteLine("Invalid input. Please try again.");
            }
        }
    }
}
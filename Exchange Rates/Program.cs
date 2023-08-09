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

        string inputDate = InputDate();

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
                    List<ExchangeRate> responseExchangeRates = JsonSerializer.Deserialize<List<ExchangeRate>>(responseBody);

                    foreach (var exRate in responseExchangeRates)
                    {
                        var currencyAbbreviation = exRate.Cur_Abbreviation;

                        // Show data only for selected currencies
                        if (currencyAbbreviations.Contains(currencyAbbreviation))
                        {
                            Console.WriteLine($"Exchange rate for {currencyAbbreviation} is {exRate.Cur_OfficialRate}");
                        }
                    }
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
internal class Program
{
    private static readonly string baseUrl = "https://api.nbrb.by/";
    private static readonly string[] currencyCodes = { "USD", "EUR" };

    static async Task Main(string[] args)
    {
        Console.WriteLine("Welcome! Here you can find exchange rates for USD and EUR for specific dates.");

        string inputDate = InputDate();

        using (HttpClient client = new HttpClient())
        {
            // separate requests for currencies
            foreach (string code in currencyCodes)
            {
                string completeUrl = $"{baseUrl}ExRates/Rates/{code}?onDate={inputDate}&ParamMode=2";

                try
                {
                    HttpResponseMessage response = await client.GetAsync(completeUrl);
                    response.EnsureSuccessStatusCode();

                    string responseBody = await response.Content.ReadAsStringAsync();
                    Console.WriteLine(responseBody);
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }
        }
    }

    private static string InputDate()
    {
        Console.Write("Enter a date: ");
        bool isDateValid = DateTime.TryParse(Console.ReadLine(), out DateTime parsedDate);

        if (!isDateValid)
        {
            Console.WriteLine("Invalid input. Please try again.");
            InputDate();
        }

        return parsedDate.ToUniversalTime().ToString("yyyy-MM-dd");
    }
}
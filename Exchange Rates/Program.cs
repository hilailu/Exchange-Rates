using GemBox.Spreadsheet;

namespace ExchangeRates
{
    internal class Program
    {
        private static ExchangeRateService _exchangeRateService = new ExchangeRateService();

        static async Task Main(string[] args)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            Console.WriteLine("Welcome! Here you can find exchange rates for USD and EUR for specific dates.");

            while (true)
            {
                Console.WriteLine("\nPlease select an option:\n1. Input date manually in console.\n2. Read dates from an Excel file.");

                var option = Console.ReadLine();

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

        private static async Task InputDateAndDisplayExchangeRates()
        {
            var inputDate = InputDate();
            List<ExchangeRate>? exchangeRates = await _exchangeRateService.GetExchangeRateForDate(inputDate);
            DisplayExchangeRates(exchangeRates);
        }

        private static async Task InputExcelPathAndWriteExchangeRates()
        {
            Console.Write("Enter the full path to the Excel file: ");
            var excelFilePath = Console.ReadLine();

            if (File.Exists(excelFilePath) && Path.GetExtension(excelFilePath).Equals(".xlsx"))
            {
                var dateExchangeRates = await _exchangeRateService.ReadExcelAndGetExchangeRatesWithDates(excelFilePath);
                _exchangeRateService.WriteExchangeRatesToExcelFile(excelFilePath, dateExchangeRates);
                Console.WriteLine("File modified successfully.");
            }
            else
            {
                Console.WriteLine("Invalid path or file format.");
            }
        }

        private static void DisplayExchangeRates(List<ExchangeRate>? exchangeRates)
        {
            foreach (var exRate in exchangeRates)
            {
                Console.WriteLine($"Exchange rate for {exRate.Abbreviation} is {exRate.OfficialRate}");
            }
        }

        private static string InputDate()
        {
            while (true)
            {
                Console.Write("Enter a date: ");
                var date = _exchangeRateService.ParseDate(Console.ReadLine());
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
    }
}
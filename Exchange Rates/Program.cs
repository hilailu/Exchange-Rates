using ExchangeRatesLib;
using GemBox.Spreadsheet;

namespace ExchangeRates
{
    internal class Program
    {
        private static ExchangeRateService _exchangeRateService = new ExchangeRateService();

        static async Task Main(string[] args)
        {
            SpreadsheetInfo.SetLicense(Constants.GemBoxKey);

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
            var exchangeRates = await _exchangeRateService.GetExchangeRateForDate(inputDate);
            DisplayExchangeRates(exchangeRates);
        }

        private static async Task InputExcelPathAndWriteExchangeRates()
        {
            Console.Write("Enter the full path to the Excel file: ");
            var excelFilePath = Console.ReadLine();

            if (_exchangeRateService.IsValidExcelFile(excelFilePath))
            {
                var dateExchangeRates = await _exchangeRateService.ReadExcelAndGetExchangeRatesWithDates(excelFilePath);
                var isWriteSuccessful = _exchangeRateService.IsWriteExchangeRatesToExcelFileSuccessful(excelFilePath, dateExchangeRates);
                Console.WriteLine(isWriteSuccessful ? "File modified successfully." : Constants.ErrorNoDateEntries);
            }
            else
            {
                Console.WriteLine(Constants.ErrorInvalidPathOrFormat);
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
                    Console.WriteLine();
                }
            }
        }
    }
}
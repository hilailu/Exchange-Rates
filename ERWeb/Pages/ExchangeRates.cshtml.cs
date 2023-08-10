using ERWeb.Data;
using ExchangeRates;
using GemBox.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;

namespace ERWeb.Pages
{
    public class ExchangeRatesModel : PageModel
    {
        private readonly ExchangeRateService _exchangeRateService;
        private readonly ApplicationDbContext _dbContext;
        public string ErrorMessage { get; private set; }

        public ExchangeRatesModel(ExchangeRateService exchangeRateService, ApplicationDbContext dbContext)
        {
            _exchangeRateService = exchangeRateService;
            _dbContext = dbContext;
        }

        [BindProperty]
        public IFormFile UploadedFile { get; set; }

        public List<DateExchangeRateModel> DateExchangeRates { get; set; }

        public async Task OnPostAsync()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            if (UploadedFile != null && UploadedFile.Length > 0 && Path.GetExtension(UploadedFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                using (var memoryStream = new MemoryStream())
                {
                    await UploadedFile.CopyToAsync(memoryStream);

                    var workbook = ExcelFile.Load(memoryStream, LoadOptions.XlsxDefault);
                    var worksheet = workbook.Worksheets[0];

                    DateExchangeRates = new List<DateExchangeRateModel>();

                    // Assuming first row contains headers
                    for (int row = 1; row < worksheet.Rows.Count; row++)
                    {
                        var excelDate = worksheet.Cells[row, 0].Value?.ToString();
                        var parsedDate = _exchangeRateService.ParseDate(excelDate);
                        if (!string.IsNullOrEmpty(parsedDate))
                        {
                            List<ExchangeRate> exchangeRates = await _exchangeRateService.GetExchangeRateForDate(parsedDate);

                            var usdRate = exchangeRates?.FirstOrDefault(rate => rate.Abbreviation == "USD")?.OfficialRate ?? 0.0m;
                            var eurRate = exchangeRates?.FirstOrDefault(rate => rate.Abbreviation == "EUR")?.OfficialRate ?? 0.0m;

                            var dateExRate = new DateExchangeRateModel { Id = Guid.NewGuid(), Date = parsedDate, USDRate = usdRate, EURRate = eurRate };

                            DateExchangeRates.Add(dateExRate);
                            _dbContext.DateExchangeRates.Add(dateExRate);
                        }
                    }
                }

                try
                {
                    await _dbContext.SaveChangesAsync();
                }
                catch (Exception ex)
                {
                    ErrorMessage = "Some entries weren't saved to the database. We don't store duplicate dates.";
                }
            }
            else
            {
                ErrorMessage = "Please provide an *.xlsx file.";
            }
        }
    }

    public class DateExchangeRateModel
    {
        [Key]
        public Guid Id { get; set; }

        public string Date { get; set; }

        public decimal USDRate { get; set; }
        public decimal EURRate { get; set; }
    }

}
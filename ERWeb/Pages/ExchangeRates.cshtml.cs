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

            if (UploadedFile != null && UploadedFile.Length > 0)
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
                            var exchangeRate = new DateExchangeRateModel { Id = Guid.NewGuid(), Date = parsedDate, ExchangeRates = exchangeRates };
                            DateExchangeRates.Add(exchangeRate);

                            _dbContext.DateExchangeRates.Add(exchangeRate);
                        }
                    }
                }
                await _dbContext.SaveChangesAsync();
            }
        }
    }

    public class DateExchangeRateModel
    {
        [Key]
        public Guid Id { get; set; }

        public string Date { get; set; }
        public List<ExchangeRate> ExchangeRates { get; set; }
    }

}
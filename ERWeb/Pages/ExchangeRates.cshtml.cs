using ERWeb.Data;
using ExchangeRates;
using GemBox.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

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

        public List<DateExchangeRate> DateExchangeRates { get; set; }

        public async Task OnPostAsync()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            if (UploadedFile != null && UploadedFile.Length > 0 && Path.GetExtension(UploadedFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                // Get temporary file path
                var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(UploadedFile.FileName));
                using (var stream = new FileStream(tempFilePath, FileMode.Create))
                {
                    await UploadedFile.CopyToAsync(stream);
                }

                // Retrieve info from excel file
                DateExchangeRates = await _exchangeRateService.ReadExcelAndGetExchangeRatesWithDates(tempFilePath);
                DateExchangeRates.ForEach(rate => rate.Id = Guid.NewGuid());
                _dbContext.DateExchangeRates.AddRange(DateExchangeRates);        

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
}
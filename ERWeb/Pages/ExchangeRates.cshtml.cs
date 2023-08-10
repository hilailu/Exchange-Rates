using ERWeb.Data;
using ExchangeRatesLib;
using GemBox.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;

namespace ERWeb.Pages
{
    public class ExchangeRatesModel : PageModel
    {
        private readonly ExchangeRateService _exchangeRateService;
        private readonly ApplicationDbContext _dbContext;

        [BindProperty]
        public IFormFile UploadedFile { get; set; }
        public List<DateExchangeRate> DateExchangeRates { get; set; }
        public string ErrorMessage { get; private set; }

        public ExchangeRatesModel(ExchangeRateService exchangeRateService, ApplicationDbContext dbContext)
        {
            SpreadsheetInfo.SetLicense(Constants.GemBoxKey);

            _exchangeRateService = exchangeRateService;
            _dbContext = dbContext;
        }

        public async Task OnPostAsync()
        {
            if (UploadedFile != null && UploadedFile.Length > 0)
            {
                // Get temporary file path
                var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(UploadedFile.FileName));

                if (!_exchangeRateService.IsValidExcelFile(tempFilePath, true))
                {
                    ErrorMessage = Constants.ErrorInvalidPathOrFormat;
                    return;
                }

                using (var stream = new FileStream(tempFilePath, FileMode.Create))
                {
                    await UploadedFile.CopyToAsync(stream);
                }

                // Retrieve info from excel file
                DateExchangeRates = await _exchangeRateService.ReadExcelAndGetExchangeRatesWithDates(tempFilePath);

                if (DateExchangeRates.Count == 0)
                {
                    ErrorMessage = Constants.ErrorNoDateEntries;
                    return;
                }

                DateExchangeRates.ForEach(rate => rate.Id = Guid.NewGuid());

                try
                {
                    // Add processed data to database
                    _dbContext.DateExchangeRates.AddRange(DateExchangeRates);
                    await _dbContext.SaveChangesAsync();
                }
                catch (DbUpdateException ex) when (IsDuplicateKeyException(ex))
                {
                    ErrorMessage = "Some entries weren't saved to the database. We don't store duplicate dates.";
                }
                catch (Exception ex)
                {
                    ErrorMessage = "Error saving data to database. Try again later.";
                }
            }
        }

        private bool IsDuplicateKeyException(DbUpdateException ex)
            => ex.InnerException is SqlException sqlException && (sqlException.Number == 2601 || sqlException.Number == 2627);
    }
}
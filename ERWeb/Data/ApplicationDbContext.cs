using ERWeb.Pages;
using ExchangeRates;
using Microsoft.EntityFrameworkCore;

namespace ERWeb.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }

        public DbSet<DateExchangeRateModel> DateExchangeRates { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<DateExchangeRateModel>()
                .HasKey(d => d.Id);

            modelBuilder.Entity<DateExchangeRateModel>()
                .OwnsMany(d => d.ExchangeRates, e =>
                {
                    e.Property(exchangeRate => exchangeRate.Abbreviation);
                    e.Property(exchangeRate => exchangeRate.OfficialRate);
                });
        }
    }
}

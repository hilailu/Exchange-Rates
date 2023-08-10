using ExchangeRates;
using Microsoft.EntityFrameworkCore;

namespace ERWeb.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }

        public DbSet<DateExchangeRate> DateExchangeRates { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<DateExchangeRate>()
                .HasKey(d => d.Id);

            modelBuilder.Entity<DateExchangeRate>()
                .HasIndex(d => d.Date)
                .IsUnique();
        }
    }
}

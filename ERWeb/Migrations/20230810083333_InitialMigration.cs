using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ERWeb.Migrations
{
    /// <inheritdoc />
    public partial class InitialMigration : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "DateExchangeRates",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    Date = table.Column<string>(type: "nvarchar(450)", nullable: false),
                    USDRate = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    EURRate = table.Column<decimal>(type: "decimal(18,2)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DateExchangeRates", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_DateExchangeRates_Date",
                table: "DateExchangeRates",
                column: "Date",
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "DateExchangeRates");
        }
    }
}

using System.Text.Json.Serialization;

namespace ExchangeRatesLib
{
    public class ExchangeRate
    {
        [JsonPropertyName("Cur_Abbreviation")]
        public string Abbreviation { get; set; }

        [JsonPropertyName("Cur_OfficialRate")]
        public decimal OfficialRate { get; set; }
    }

    public class DateExchangeRate
    {
        public Guid Id { get; set; }

        public string Date { get; set; }

        public decimal USDRate { get; set; }
        public decimal EURRate { get; set; }
    }
}

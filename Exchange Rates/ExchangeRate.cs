using System.Text.Json.Serialization;

namespace ExchangeRates
{
    public class ExchangeRate
    {
        [JsonPropertyName("Cur_Abbreviation")]
        public string Abbreviation { get; set; }

        [JsonPropertyName("Cur_OfficialRate")]
        public decimal OfficialRate { get; set; }
    }
}

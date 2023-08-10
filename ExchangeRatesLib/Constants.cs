namespace ExchangeRatesLib
{
    public static class Constants
    {
        public static readonly string Date = "Date";
        public static readonly string USD = "USD";
        public static readonly string EUR = "EUR";

        public static readonly string BaseApiUrl = "https://api.nbrb.by/";
        public static readonly string GemBoxKey = "FREE-LIMITED-KEY";

        public static readonly string ErrorNoDateEntries = "No date entries found.";
        public static readonly string ErrorInvalidPathOrFormat = "Invalid path or file format. Please provide an *.xlsx file.";
    }

    enum Indexes
    {
        Date,
        USD,
        EUR,
    }
}

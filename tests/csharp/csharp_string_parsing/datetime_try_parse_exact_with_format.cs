// vybe-test: csharp/csharp_string_parsing/datetime_try_parse_exact_with_format
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool ok=System.DateTime.TryParseExact("2024-01-15","yyyy-MM-dd",
    System.Globalization.CultureInfo.InvariantCulture,
    System.Globalization.DateTimeStyles.None,out var dt);
__Check((ok).ToString(), "True"); __Check((dt.Day).ToString(), "15");

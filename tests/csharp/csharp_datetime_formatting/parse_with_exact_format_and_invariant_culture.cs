// vybe-test: csharp/csharp_datetime_formatting/parse_with_exact_format_and_invariant_culture
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = System.DateTime.ParseExact("2024-03-21","yyyy-MM-dd",
    System.Globalization.CultureInfo.InvariantCulture);
__Check((d.Year).ToString(), "2024"); __Check((d.Month).ToString(), "3"); __Check((d.Day).ToString(), "21");

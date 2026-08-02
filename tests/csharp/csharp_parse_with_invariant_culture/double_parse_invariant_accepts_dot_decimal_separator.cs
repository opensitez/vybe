// vybe-test: csharp/csharp_parse_with_invariant_culture/double_parse_invariant_accepts_dot_decimal_separator
// origin: languages/csharp/tests/csharp/test_csharp_parse_with_invariant_culture.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double value = double.Parse("3.5", System.Globalization.CultureInfo.InvariantCulture);
__Check((value).ToString(), "3.5");

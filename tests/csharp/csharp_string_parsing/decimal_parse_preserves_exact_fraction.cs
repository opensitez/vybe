// vybe-test: csharp/csharp_string_parsing/decimal_parse_preserves_exact_fraction
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d=decimal.Parse("0.1",System.Globalization.CultureInfo.InvariantCulture);
__Check((d+0.2m==0.3m).ToString(), "True");

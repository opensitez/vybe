// vybe-test: csharp/csharp_parsing_formatting/int_try_parse_reports_true_for_valid_digits
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ok = int.TryParse("42", out var value); __Check((ok).ToString(), "True"); __Check((value).ToString(), "42");

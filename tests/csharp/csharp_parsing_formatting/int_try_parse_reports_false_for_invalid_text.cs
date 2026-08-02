// vybe-test: csharp/csharp_parsing_formatting/int_try_parse_reports_false_for_invalid_text
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ok = int.TryParse("4x", out var value); __Check((ok).ToString(), "False"); __Check((value).ToString(), "0");

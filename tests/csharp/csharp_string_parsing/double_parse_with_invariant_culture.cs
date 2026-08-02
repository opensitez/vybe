// vybe-test: csharp/csharp_string_parsing/double_parse_with_invariant_culture
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d=double.Parse("3.14",System.Globalization.CultureInfo.InvariantCulture);
__Check((d).ToString(), "3.14");

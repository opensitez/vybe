// vybe-test: csharp/csharp_parsing_formatting/trim_then_parse_allows_surrounding_whitespace
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((int.Parse(" 12 ".Trim())).ToString(), "12");

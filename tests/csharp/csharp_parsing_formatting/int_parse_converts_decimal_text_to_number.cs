// vybe-test: csharp/csharp_parsing_formatting/int_parse_converts_decimal_text_to_number
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((int.Parse("42") + 1).ToString(), "43");

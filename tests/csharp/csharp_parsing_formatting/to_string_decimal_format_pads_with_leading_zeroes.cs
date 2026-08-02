// vybe-test: csharp/csharp_parsing_formatting/to_string_decimal_format_pads_with_leading_zeroes
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((7.ToString("D4")).ToString(), "0007");

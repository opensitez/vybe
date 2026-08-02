// vybe-test: csharp/csharp_parsing_formatting/to_string_hex_format_outputs_uppercase_hex_digits
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((255.ToString("X")).ToString(), "FF");

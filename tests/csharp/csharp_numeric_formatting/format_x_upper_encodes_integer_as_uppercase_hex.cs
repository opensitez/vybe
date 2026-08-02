// vybe-test: csharp/csharp_numeric_formatting/format_x_upper_encodes_integer_as_uppercase_hex
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((255.ToString("X")).ToString(), "FF");

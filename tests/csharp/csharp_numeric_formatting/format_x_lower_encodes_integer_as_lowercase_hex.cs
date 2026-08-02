// vybe-test: csharp/csharp_numeric_formatting/format_x_lower_encodes_integer_as_lowercase_hex
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((255.ToString("x")).ToString(), "ff");

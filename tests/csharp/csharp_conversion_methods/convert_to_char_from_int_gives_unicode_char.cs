// vybe-test: csharp/csharp_conversion_methods/convert_to_char_from_int_gives_unicode_char
// origin: languages/csharp/tests/csharp/test_csharp_conversion_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// conversion_methods
__Check((System.Convert.ToChar(65)).ToString(), "A");

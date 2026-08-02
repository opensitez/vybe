// vybe-test: csharp/csharp_string_format/string_format_single_placeholder_interpolates_argument
// origin: languages/csharp/tests/csharp/test_csharp_string_format.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("Hello {0}!", "world")).ToString(), "Hello world!");

// vybe-test: csharp/csharp_strings/string_chained_methods
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("  Hello World  ".Trim().ToUpper()).ToString(), "HELLO WORLD");

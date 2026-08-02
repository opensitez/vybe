// vybe-test: csharp/strings/chained_string_methods
// origin: languages/csharp/tests/csharp/test_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("  Hello World  ".Trim().ToUpper()).ToString(), "HELLO WORLD");

// vybe-test: csharp/csharp_string_methods/trim_removes_leading_and_trailing_whitespace
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("  hello  ".Trim()).ToString(), "hello");

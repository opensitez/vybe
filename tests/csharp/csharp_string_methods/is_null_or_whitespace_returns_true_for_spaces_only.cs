// vybe-test: csharp/csharp_string_methods/is_null_or_whitespace_returns_true_for_spaces_only
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.IsNullOrWhiteSpace("   ")).ToString(), "True");

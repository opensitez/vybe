// vybe-test: csharp/csharp_string_comparison/starts_with_case_insensitive_returns_true_for_different_case_prefix
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("HELLO".StartsWith("hell",System.StringComparison.OrdinalIgnoreCase)).ToString(), "True");

// vybe-test: csharp/csharp_string_comparison/ordinal_comparison_is_case_sensitive_by_default
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Compare("A","a",System.StringComparison.Ordinal) != 0).ToString(), "True");

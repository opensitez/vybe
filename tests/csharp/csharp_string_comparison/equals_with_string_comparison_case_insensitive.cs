// vybe-test: csharp/csharp_string_comparison/equals_with_string_comparison_case_insensitive
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("ABC".Equals("abc",System.StringComparison.OrdinalIgnoreCase)).ToString(), "True");

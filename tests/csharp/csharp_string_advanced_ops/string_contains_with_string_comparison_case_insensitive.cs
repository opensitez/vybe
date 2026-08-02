// vybe-test: csharp/csharp_string_advanced_ops/string_contains_with_string_comparison_case_insensitive
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("Hello World".Contains("world",System.StringComparison.OrdinalIgnoreCase)).ToString(), "True");

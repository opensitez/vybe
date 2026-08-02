// vybe-test: csharp/csharp_string_comparison/index_of_with_string_comparison_finds_case_insensitive
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("fooBAR".IndexOf("bar",System.StringComparison.OrdinalIgnoreCase)).ToString(), "3");

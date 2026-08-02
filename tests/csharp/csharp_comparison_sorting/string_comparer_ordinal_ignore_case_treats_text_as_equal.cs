// vybe-test: csharp/csharp_comparison_sorting/string_comparer_ordinal_ignore_case_treats_text_as_equal
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.StringComparer.OrdinalIgnoreCase.Equals("abc", "ABC")).ToString(), "True");

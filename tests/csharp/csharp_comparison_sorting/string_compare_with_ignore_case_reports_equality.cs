// vybe-test: csharp/csharp_comparison_sorting/string_compare_with_ignore_case_reports_equality
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Compare("abc", "ABC", true)).ToString(), "0");

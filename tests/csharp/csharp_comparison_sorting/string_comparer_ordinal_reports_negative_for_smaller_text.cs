// vybe-test: csharp/csharp_comparison_sorting/string_comparer_ordinal_reports_negative_for_smaller_text
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.StringComparer.Ordinal.Compare("a", "b")).ToString(), "-1");

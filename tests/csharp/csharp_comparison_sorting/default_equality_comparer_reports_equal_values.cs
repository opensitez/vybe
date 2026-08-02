// vybe-test: csharp/csharp_comparison_sorting/default_equality_comparer_reports_equal_values
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Collections.Generic.EqualityComparer<int>.Default.Equals(4, 4)).ToString(), "True");

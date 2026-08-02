// vybe-test: csharp/csharp_comparison_sorting/compareto_on_integer_reports_positive_for_larger_value
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((9.CompareTo(3)).ToString(), "1");

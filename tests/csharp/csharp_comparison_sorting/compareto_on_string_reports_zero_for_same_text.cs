// vybe-test: csharp/csharp_comparison_sorting/compareto_on_string_reports_zero_for_same_text
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("abc".CompareTo("abc")).ToString(), "0");

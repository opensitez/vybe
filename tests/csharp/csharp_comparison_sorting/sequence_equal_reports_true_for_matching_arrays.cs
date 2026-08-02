// vybe-test: csharp/csharp_comparison_sorting/sequence_equal_reports_true_for_matching_arrays
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq; __Check((new[] { 1, 2 }.SequenceEqual(new[] { 1, 2 })).ToString(), "True");

// vybe-test: csharp/csharp_comparison_sorting/sequence_equal_reports_false_for_different_arrays
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq; __Check((new[] { 1, 2 }.SequenceEqual(new[] { 2, 1 })).ToString(), "False");

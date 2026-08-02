// vybe-test: csharp/csharp_comparison_sorting/max_returns_largest_value_from_sequence
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq; __Check((new[] { 2, 9, 4 }.Max()).ToString(), "9");

// vybe-test: csharp/csharp_linq_materialization/linq_min_returns_smallest_element_by_default_comparer
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
__Check((new[] { 3, 9, 4 }.Min()).ToString(), "3");

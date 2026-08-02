// vybe-test: csharp/csharp_linq_materialization/linq_max_returns_greatest_element_by_default_comparer
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
__Check((new[] { 3, 9, 4 }.Max()).ToString(), "9");

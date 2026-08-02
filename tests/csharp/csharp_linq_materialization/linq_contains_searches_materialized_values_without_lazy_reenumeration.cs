// vybe-test: csharp/csharp_linq_materialization/linq_contains_searches_materialized_values_without_lazy_reenumeration
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
var data = new[] { "a", "b" };
__Check((data.Contains("b")).ToString(), "True");

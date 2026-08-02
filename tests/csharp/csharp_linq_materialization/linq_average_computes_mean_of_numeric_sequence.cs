// vybe-test: csharp/csharp_linq_materialization/linq_average_computes_mean_of_numeric_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
__Check((new[] { 2, 4, 6 }.Average()).ToString(), "4");

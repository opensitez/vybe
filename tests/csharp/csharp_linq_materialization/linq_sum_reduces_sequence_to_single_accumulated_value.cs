// vybe-test: csharp/csharp_linq_materialization/linq_sum_reduces_sequence_to_single_accumulated_value
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
__Check((new[] { 1, 2, 3, 4 }.Sum()).ToString(), "10");

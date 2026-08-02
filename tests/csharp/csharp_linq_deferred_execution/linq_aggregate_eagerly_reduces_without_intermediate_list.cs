// vybe-test: csharp/csharp_linq_deferred_execution/linq_aggregate_eagerly_reduces_without_intermediate_list
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
int sum = new[] { 1, 2, 3, 4 }.Aggregate(0, (acc, x) => acc + x);
__Check((sum).ToString(), "10");

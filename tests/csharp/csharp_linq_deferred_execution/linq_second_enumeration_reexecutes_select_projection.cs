// vybe-test: csharp/csharp_linq_deferred_execution/linq_second_enumeration_reexecutes_select_projection
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
int projections = 0;
var query = new[] { 5 }.Select(x => { projections++; return x + 1; });
__Check((query.First()).ToString(), "6");
__Check((projections).ToString(), "1");
__Check((query.First()).ToString(), "6");
__Check((projections).ToString(), "2");

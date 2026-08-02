// vybe-test: csharp/csharp_linq_materialization/linq_distinct_to_list_collapses_duplicates_during_materialization
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
var unique = new[] { 1, 1, 2, 2, 3 }.Distinct().ToList();
__Check((unique.Count).ToString(), "3");

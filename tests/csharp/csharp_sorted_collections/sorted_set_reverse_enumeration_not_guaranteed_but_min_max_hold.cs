// vybe-test: csharp/csharp_sorted_collections/sorted_set_reverse_enumeration_not_guaranteed_but_min_max_hold
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 4, 1, 7 }; __Check((ss.Min).ToString(), "1"); __Check((ss.Max).ToString(), "7");

// vybe-test: csharp/csharp_sorted_collections/sorted_set_singleton_min_equals_max
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 42 }; __Check((ss.Min).ToString(), "42"); __Check((ss.Max).ToString(), "42");

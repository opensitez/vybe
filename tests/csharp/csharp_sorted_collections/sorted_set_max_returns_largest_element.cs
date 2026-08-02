// vybe-test: csharp/csharp_sorted_collections/sorted_set_max_returns_largest_element
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 8, 2, 5 }; __Check((ss.Max).ToString(), "8");

// vybe-test: csharp/csharp_sorted_collections/sorted_set_overlaps_false_when_disjoint
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new SortedSet<int> { 1, 2 }; var b = new SortedSet<int> { 5, 6 }; __Check((a.Overlaps(b)).ToString(), "False");

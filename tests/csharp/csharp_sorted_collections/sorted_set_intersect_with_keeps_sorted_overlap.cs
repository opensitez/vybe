// vybe-test: csharp/csharp_sorted_collections/sorted_set_intersect_with_keeps_sorted_overlap
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new SortedSet<int> { 1, 2, 3, 4 }; a.IntersectWith(new[] { 3, 4, 5 }); __Check((a.Count).ToString(), "2"); __Check((a.Min).ToString(), "3");

// vybe-test: csharp/csharp_hashset_set_algebra/intersect_with_empty_other_clears_set
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.IntersectWith(new int[] { }); __Check((a.Count).ToString(), "0");

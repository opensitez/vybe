// vybe-test: csharp/csharp_hashset_set_algebra/intersect_with_self_is_identity
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 7, 8 }; a.IntersectWith(a); __Check((a.Count).ToString(), "2");

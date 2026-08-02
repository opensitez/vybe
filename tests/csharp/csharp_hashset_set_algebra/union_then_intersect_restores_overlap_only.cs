// vybe-test: csharp/csharp_hashset_set_algebra/union_then_intersect_restores_overlap_only
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.UnionWith(new[] { 2, 3 }); a.IntersectWith(new[] { 2, 5 }); __Check((a.Count).ToString(), "1"); __Check((a.Contains(2)).ToString(), "True");

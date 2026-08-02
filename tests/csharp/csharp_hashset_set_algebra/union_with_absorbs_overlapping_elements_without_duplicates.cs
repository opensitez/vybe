// vybe-test: csharp/csharp_hashset_set_algebra/union_with_absorbs_overlapping_elements_without_duplicates
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.UnionWith(new[] { 3, 4 }); __Check((a.Count).ToString(), "4");

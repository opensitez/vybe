// vybe-test: csharp/csharp_hashset_set_algebra/union_with_enumerable_adds_all_new_items
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1 }; var extra = new List<int> { 2, 3 }; a.UnionWith(extra); __Check((a.Contains(3)).ToString(), "True");

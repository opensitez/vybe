// vybe-test: csharp/csharp_hashset_set_algebra/union_with_after_union_accumulates_third_batch
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1 }; a.UnionWith(new[] { 2 }); a.UnionWith(new[] { 3 }); __Check((a.Count).ToString(), "3");

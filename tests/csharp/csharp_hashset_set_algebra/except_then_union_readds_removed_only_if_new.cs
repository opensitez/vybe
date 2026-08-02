// vybe-test: csharp/csharp_hashset_set_algebra/except_then_union_readds_removed_only_if_new
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.ExceptWith(new[] { 2 }); a.UnionWith(new[] { 2, 4 }); __Check((a.Contains(2)).ToString(), "True"); __Check((a.Contains(4)).ToString(), "True");

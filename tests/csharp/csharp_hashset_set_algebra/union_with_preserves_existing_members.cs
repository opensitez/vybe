// vybe-test: csharp/csharp_hashset_set_algebra/union_with_preserves_existing_members
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 10 }; a.UnionWith(new[] { 20 }); __Check((a.Contains(10)).ToString(), "True");

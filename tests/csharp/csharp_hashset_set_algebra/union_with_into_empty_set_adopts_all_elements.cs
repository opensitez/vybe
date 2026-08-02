// vybe-test: csharp/csharp_hashset_set_algebra/union_with_into_empty_set_adopts_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int>(); a.UnionWith(new[] { 7, 8 }); __Check((a.Contains(7)).ToString(), "True"); __Check((a.Count).ToString(), "2");

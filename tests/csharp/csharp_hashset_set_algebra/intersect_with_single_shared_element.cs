// vybe-test: csharp/csharp_hashset_set_algebra/intersect_with_single_shared_element
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.IntersectWith(new[] { 3, 9 }); __Check((a.Contains(3)).ToString(), "True"); __Check((a.Contains(1)).ToString(), "False");

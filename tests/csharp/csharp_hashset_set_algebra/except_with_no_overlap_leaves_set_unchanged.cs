// vybe-test: csharp/csharp_hashset_set_algebra/except_with_no_overlap_leaves_set_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 5, 6 }; a.ExceptWith(new[] { 1, 2 }); __Check((a.Count).ToString(), "2");

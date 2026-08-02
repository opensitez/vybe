// vybe-test: csharp/csharp_hashset_set_algebra/except_with_removes_all_when_other_is_superset
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.ExceptWith(new[] { 1, 2, 3 }); __Check((a.Count).ToString(), "0");

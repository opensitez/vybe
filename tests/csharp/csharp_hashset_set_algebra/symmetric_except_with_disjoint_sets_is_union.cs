// vybe-test: csharp/csharp_hashset_set_algebra/symmetric_except_with_disjoint_sets_is_union
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.SymmetricExceptWith(new[] { 3, 4 }); __Check((a.Count).ToString(), "4");

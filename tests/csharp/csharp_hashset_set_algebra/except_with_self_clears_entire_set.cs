// vybe-test: csharp/csharp_hashset_set_algebra/except_with_self_clears_entire_set
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.ExceptWith(a); __Check((a.Count).ToString(), "0");

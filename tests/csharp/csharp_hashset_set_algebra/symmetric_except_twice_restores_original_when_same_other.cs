// vybe-test: csharp/csharp_hashset_set_algebra/symmetric_except_twice_restores_original_when_same_other
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; var other = new[] { 2, 4 }; a.SymmetricExceptWith(other); a.SymmetricExceptWith(other); __Check((a.SetEquals(new HashSet<int> { 1, 2, 3 })).ToString(), "True");

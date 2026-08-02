// vybe-test: csharp/csharp_hashset_set_algebra/symmetric_except_with_keeps_elements_in_exactly_one_set
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.SymmetricExceptWith(new[] { 2, 3, 4 }); __Check((a.Contains(1)).ToString(), "True"); __Check((a.Contains(4)).ToString(), "True"); __Check((a.Contains(2)).ToString(), "False");

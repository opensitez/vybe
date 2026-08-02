// vybe-test: csharp/csharp_hashset_set_algebra/except_with_single_element_removal
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 10, 20, 30 }; a.ExceptWith(new[] { 20 }); __Check((a.Contains(20)).ToString(), "False");

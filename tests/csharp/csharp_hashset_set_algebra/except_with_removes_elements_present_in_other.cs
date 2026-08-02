// vybe-test: csharp/csharp_hashset_set_algebra/except_with_removes_elements_present_in_other
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.ExceptWith(new[] { 2, 4 }); __Check((a.Count).ToString(), "2"); __Check((a.Contains(1)).ToString(), "True");

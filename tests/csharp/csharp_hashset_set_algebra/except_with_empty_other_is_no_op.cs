// vybe-test: csharp/csharp_hashset_set_algebra/except_with_empty_other_is_no_op
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 9 }; a.ExceptWith(new int[] { }); __Check((a.Count).ToString(), "1");

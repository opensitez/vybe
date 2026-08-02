// vybe-test: csharp/csharp_hashset_set_algebra/set_equals_self_is_true
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 9, 8 }; __Check((a.SetEquals(a)).ToString(), "True");

// vybe-test: csharp/csharp_hashset_set_algebra/set_equals_empty_sets
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int>(); var b = new HashSet<int>(); __Check((a.SetEquals(b)).ToString(), "True");

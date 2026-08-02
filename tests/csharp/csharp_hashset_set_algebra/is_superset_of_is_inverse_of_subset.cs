// vybe-test: csharp/csharp_hashset_set_algebra/is_superset_of_is_inverse_of_subset
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; var b = new HashSet<int> { 1, 2 }; __Check((a.IsSupersetOf(b)).ToString(), "True");

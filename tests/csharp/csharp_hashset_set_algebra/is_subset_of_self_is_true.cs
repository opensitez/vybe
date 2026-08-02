// vybe-test: csharp/csharp_hashset_set_algebra/is_subset_of_self_is_true
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 4, 5 }; __Check((a.IsSubsetOf(a)).ToString(), "True");

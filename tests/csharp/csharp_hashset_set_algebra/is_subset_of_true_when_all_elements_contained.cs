// vybe-test: csharp/csharp_hashset_set_algebra/is_subset_of_true_when_all_elements_contained
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; var b = new HashSet<int> { 1, 2, 3 }; __Check((a.IsSubsetOf(b)).ToString(), "True");

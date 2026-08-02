// vybe-test: csharp/csharp_hashset_set_algebra/is_subset_of_empty_set_only_for_empty
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var empty = new HashSet<int>(); var nonempty = new HashSet<int> { 1 }; __Check((empty.IsSubsetOf(nonempty)).ToString(), "True"); __Check((nonempty.IsSubsetOf(empty)).ToString(), "False");

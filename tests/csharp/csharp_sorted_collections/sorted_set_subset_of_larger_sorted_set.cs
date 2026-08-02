// vybe-test: csharp/csharp_sorted_collections/sorted_set_subset_of_larger_sorted_set
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var small = new SortedSet<int> { 2, 3 }; var big = new SortedSet<int> { 1, 2, 3, 4 }; __Check((small.IsSubsetOf(big)).ToString(), "True");

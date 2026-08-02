// vybe-test: csharp/csharp_sorted_collections/sorted_set_remove_min_then_new_min
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3 }; ss.Remove(1); __Check((ss.Min).ToString(), "2");

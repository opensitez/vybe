// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_count_reflects_insertions
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var sd = new SortedDictionary<int, int> { [1] = 10, [2] = 20, [3] = 30 }; __Check((sd.Count).ToString(), "3");

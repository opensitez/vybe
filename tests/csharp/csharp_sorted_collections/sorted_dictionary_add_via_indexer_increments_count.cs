// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_add_via_indexer_increments_count
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var sd = new SortedDictionary<string, int>(); sd["x"] = 1; sd["y"] = 2; __Check((sd.Count).ToString(), "2");

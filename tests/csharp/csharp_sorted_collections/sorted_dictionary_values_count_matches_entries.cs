// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_values_count_matches_entries
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "a", [2] = "b", [3] = "c" }; __Check((sd.Values.Count).ToString(), "3");

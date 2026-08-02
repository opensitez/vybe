// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_clear_resets_count
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "a" }; sd.Clear(); __Check((sd.Count).ToString(), "0");

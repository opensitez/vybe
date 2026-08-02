// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_indexer_overwrites_existing_value
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "old" }; sd[1] = "new"; __Check((sd[1]).ToString(), "new");

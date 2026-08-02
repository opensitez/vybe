// vybe-test: csharp/csharp_map_set_collections/sorted_dictionary_indexer_retrieves_inserted_value
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new SortedDictionary<int, string>(); map[2] = "two"; __Check((map[2]).ToString(), "two");

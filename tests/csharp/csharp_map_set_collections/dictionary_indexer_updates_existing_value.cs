// vybe-test: csharp/csharp_map_set_collections/dictionary_indexer_updates_existing_value
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map["a"] = 4; __Check((map["a"]).ToString(), "4");

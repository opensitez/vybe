// vybe-test: csharp/csharp_dictionary_contracts/indexer_assignment_updates_existing_entry_without_growing_count
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var map = new Dictionary<string, int> { ["x"] = 1 };
map["x"] = 9;
__Check((map["x"]).ToString(), "9");
__Check((map.Count).ToString(), "1");

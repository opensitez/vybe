// vybe-test: csharp/csharp_dictionary_contracts/indexer_read_returns_stored_value_for_existing_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var map = new Dictionary<string, int> { ["pi"] = 3 };
__Check((map["pi"]).ToString(), "3");

// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_after_indexer_insert_finds_new_entry
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); map["fresh"] = 42; __Check((map.TryGetValue("fresh", out int v)).ToString(), "True"); __Check((v).ToString(), "42");

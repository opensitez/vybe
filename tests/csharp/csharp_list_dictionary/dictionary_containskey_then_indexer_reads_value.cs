// vybe-test: csharp/csharp_list_dictionary/dictionary_containskey_then_indexer_reads_value
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 7 }; if (map.ContainsKey("a")) __Check((map["a"]).ToString(), "7");

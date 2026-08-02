// vybe-test: csharp/csharp_dictionary_tryget_patterns/indexer_write_then_containskey_sees_new_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); map["newkey"] = 77; __Check((map.ContainsKey("newkey")).ToString(), "True");

// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_true_before_indexer_read
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["safe"] = 3 }; if (map.ContainsKey("safe")) __Check((map["safe"]).ToString(), "3");
